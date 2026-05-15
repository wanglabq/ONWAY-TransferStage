import os
import time
import csv
from datetime import datetime, timedelta, timezone
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from flask import Flask, request
from flask_restful import Api, Resource
from pymodbus.client import ModbusSerialClient as ModbusClient
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
import json
import configparser
from serial import SerialException
# ──────────────────────────────────────────────────────────────
#  Global configuration -- everything now comes from INI
# ──────────────────────────────────────────────────────────────
CFG_PATH = os.path.join(os.path.dirname(__file__), "TempConfig.ini")


cfg = configparser.ConfigParser()
if not cfg.read(CFG_PATH, encoding="utf-8"):
    raise FileNotFoundError(f"[Config] Cannot read {CFG_PATH}")

# ---- security & logging ----
SECRET_TOKEN = cfg.get("Security", "secret_token", fallback="")
LOG_ENCODING = cfg.get("Logging", "encoding",     fallback="utf-8-sig")

# ---- fonts / UI ----
BASE_FONT_SIZE  = cfg.getint("UI", "base_font_size",  fallback=20)
BIG_FONT_SIZE   = cfg.getint("UI", "big_font_size",   fallback=82)
UNIT_FONT_SIZE  = cfg.getint("UI", "unit_font_size",  fallback=32)
ENTRY_FONT_SIZE = cfg.getint("UI", "entry_font_size", fallback=18)

BASE_FONT       = ("Arial", BASE_FONT_SIZE)
BASE_FONT_BOLD  = ("Arial", BASE_FONT_SIZE + 2, "bold")
BIG_FONT        = ("Arial", BIG_FONT_SIZE,  "bold")
UNIT_FONT       = ("Arial", UNIT_FONT_SIZE)
ENTRY_FONT      = ("Arial", ENTRY_FONT_SIZE)

ICON_PATH = cfg.get("UI", "icon_path", fallback="")

# ---- serial / Modbus ----  (section renamed to [Serial] in INI)
SERIAL_OPTS = {
    "port":     cfg.get("Serial", "port",     fallback="COM4"),
    "baudrate": cfg.getint("Serial", "baudrate", fallback=9600),
    "parity":   cfg.get("Serial", "parity",   fallback="N"),
    "stopbits": cfg.getint("Serial", "stopbits", fallback=1),
    "bytesize": cfg.getint("Serial", "bytesize", fallback=8),
}

# ---- other originals that remain constants ----
PARAM_REG_ADDRESSES = {18506, 18507, 18508, 18509, 18523, 18501, 2036}

# Flask & Modbus client globals
app = Flask(__name__)
api = Api(app)
client = None
stop_threads = False

# Shared data stores
data_store = {
    "Temperature": None, "SetTemperature": None, "Power": None, "PowerLimit": None,
    "Segment": None, "SegmentLeft": None, "P": None, "I": None, "D": None,
    "Cycle": None, "Correction": None, "Filter": None, "OvertempAlarm": None,
    "manual_override": False
}
last_param_values = {k: None for k in ("P", "I", "D", "Cycle", "Correction", "Filter", "OvertempAlarm")}

# --- Token check ---
def check_token():
    tok = request.headers.get("Authorization")
    return tok == SECRET_TOKEN if tok else True

# --- Modbus read/write with retry & scaling ---
def read_register(cli, address, slave=10, retries=5):
    try:
        for _ in range(retries):
            rsp = cli.read_holding_registers(address, count=1, slave=slave)
            if not rsp.isError():
                val = rsp.registers[0] / 10
                return val * 10 if address in [2036, 18523, 2092,
                                               18506, 18507, 18508,
                                               18509, 18501] else val
            time.sleep(0.1)
    except SerialException as e:
        print("[Modbus] serial error:", e)
    return None
def write_register(client, address, value, slave=10):
    # these registers need the /10 before writing
    if address in [18506, 18507, 18508, 18509, 18523, 18501, 2036]:
        value = value / 10
    client.write_register(address, int(value * 10), slave=slave)

# --- Timestamp & CSV logging ---
def _timestamp_parts():
    now = datetime.now()
    return now.strftime("%Y-%m-%d %H:%M:%S"), now.strftime("%Y-%m"), now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

def _log_csv(prefix, header, row, log_dir):
    ts, month, date_s, time_s = _timestamp_parts()
    fn = os.path.join(log_dir, f"{prefix}_{month}.csv")
    is_new = not os.path.exists(fn)
    try:
        with open(fn, "a", newline="", encoding=LOG_ENCODING) as f:
            w = csv.writer(f)
            if is_new: w.writerow(header)
            w.writerow([ts, date_s, time_s] + row)
    except PermissionError:
        print(f"[Warning] Cannot write to '{fn}' – file is open?")

# --- Modbus polling loop ---
def update_modbus_values_loop(app_ref):
    global stop_threads
    try:
        while not stop_threads and app_ref.running:
            temp = read_register(client, 18504)
            power = read_register(client, 2036)
            new_params = {
                "P": read_register(client, 18506),
                "I": read_register(client, 18507),
                "D": read_register(client, 18508),
                "Cycle": read_register(client, 18509),
                "Correction": read_register(client, 18550),
                "Filter": read_register(client, 18501),
                "OvertempAlarm": read_register(client, 2490),
            }
            data_store.update(Temperature=temp, Power=power, **new_params)

            if app_ref.data_save_var.get() and temp is not None and power is not None:
                _log_csv("TEMP_LOG", ["Timestamp","Date","Time","Temperature","Power"], [temp, power], app_ref.log_dir.get())
                if any(new_params[k] is not None and last_param_values[k] != new_params[k] for k in new_params):
                    _log_csv(
                        "PARAMETER_LOG",
                        ["Timestamp","Date","Time"] + list(new_params.keys()),
                        list(new_params.values()),
                        app_ref.log_dir.get()
                    )
                    last_param_values.update(new_params)
            time.sleep(0.1)
    except Exception as e:
        print("[Modbus] loop stopped:", e)


# --- REST API resources ---
class DataAPI(Resource):
    def get(self):
        return data_store

    def post(self):
        val = request.json.get("SetTemperature")
        if val is not None:
            write_register(client, 3000, float(val))
            data_store["SetTemperature"] = float(val)
            return {"message": "Set Temperature updated", "SetTemperature": val}
        return {"error": "Invalid input"}, 400

class SetpointAPI(Resource):
    def post(self):
        if not check_token():
            return {"error": "Unauthorized"}, 401
        val = request.json.get("SetTemperature")
        if val is not None:
            write_register(client, 3000, float(val))
            data_store["SetTemperature"] = float(val)

            print(f"[✅] Setpoint manually updated to {val}°C (manual override)")
            return {"message": "Set Temperature updated", "SetTemperature": val}
        return {"error": "Invalid input"}, 400

api.add_resource(DataAPI, "/api/data")
api.add_resource(SetpointAPI, "/setpoint")

def run_flask_app(port=5000):
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)


class MQTTManager:
    """
    MQTT manager for versioned telemetry, Home Assistant discovery, availability,
    and remote setpoint commands.

    The payload keeps the original flat keys for compatibility and adds stable
    metadata plus a normalized "metrics" object for long-term storage systems.
    """

    SCHEMA_NAME = "wanglab.controller.telemetry"
    SCHEMA_VERSION = 2

    def __init__(self, *, host, port, topic_pub,
                 client_id, username, password,
                 setpoint_cmd_topic, discovery_prefix,
                 qos=0, retain=False, enable=True,
                 gui_ref=None, device_type="onway",
                 device_id="", device_name="",
                 manufacturer="", model="",
                 entity_prefix_name="",
                 publish_interval=1.0,
                 expire_after=300,
                 history_topic=""):
        self.enable = bool(enable)
        self.client = None
        if not self.enable:
            return

        import os
        import paho.mqtt.client as mqtt

        self.topic_pub = str(topic_pub).rstrip("/")
        self.setpoint_cmd_topic = str(setpoint_cmd_topic or f"{self.topic_pub}/set")
        self.discovery_prefix = str(discovery_prefix or "homeassistant").rstrip("/")
        self.qos = int(qos)
        self.retain = bool(retain)
        self.gui_ref = gui_ref
        self.device_type = str(device_type or "controller").lower()
        self.device_id = self._sanitize_id(device_id or self._cfg_get("Device", "identifiers", self.device_type))
        self.device_name = device_name or self._cfg_get("Device", "name", self.device_id)
        self.manufacturer = manufacturer or self._cfg_get("Device", "manufacturer", "")
        self.model = model or self._cfg_get("Device", "model", "")
        self.entity_prefix_name = entity_prefix_name or self._default_entity_prefix()
        self.publish_interval = max(0.0, float(publish_interval or 0.0))
        self.expire_after = int(expire_after or 0)
        self.history_topic = str(history_topic or "").strip().rstrip("/")
        self.availability_topic = f"{self.topic_pub}/availability"
        self.schema_topic = f"{self.topic_pub}/schema"
        self.ha_status_topic = f"{self.discovery_prefix}/status"
        self.sensor_definitions = self._sensor_definitions()
        self._last_publish_monotonic = 0.0
        self._last_payload = None
        self._sequence = 0

        if not client_id:
            client_id = f"{self._sanitize_id(self.device_id)}_publisher_{os.getpid()}"

        self.client = mqtt.Client(client_id=client_id, clean_session=True)
        if username:
            self.client.username_pw_set(username, password or None)

        self.client.will_set(self.availability_topic, payload="offline", qos=self.qos, retain=True)
        self.client.on_connect = self._on_connect
        self.client.on_message = self._on_message
        self.client.on_disconnect = lambda c, u, *args: print("[MQTT] disconnected")

        try:
            self.client.connect(host, int(port), keepalive=60)
            self.client.loop_start()
        except Exception as e:
            print(f"[MQTT] connect error: {e}")
            self.enable = False

    def _cfg_get(self, section, option, fallback=""):
        try:
            if self.gui_ref is not None and hasattr(self.gui_ref, "cfg"):
                return self.gui_ref.cfg.get(section, option, fallback=fallback)
        except Exception:
            pass
        return fallback

    def _sanitize_id(self, value):
        import re
        text = str(value or "controller").strip()
        if "," in text:
            text = text.split(",", 1)[0].strip()
        text = re.sub(r"[^A-Za-z0-9_-]+", "_", text)
        return text.strip("_") or "controller"

    def _default_entity_prefix(self):
        if self.device_type == "furnace":
            return "Furnace"
        if self.device_type == "onway":
            return "Onway"
        return "Controller"

    def _device_info(self):
        info = {
            "identifiers": [self.device_id],
            "name": self.device_name,
        }
        if self.manufacturer:
            info["manufacturer"] = self.manufacturer
        if self.model:
            info["model"] = self.model
        return info

    def _sensor_definitions(self):
        defs = {
            "Temperature": {
                "name": "Temperature",
                "metric": "temperature_c",
                "unit": "°C",
                "device_class": "temperature",
                "state_class": "measurement",
                "precision": 1,
            },
            "Power": {
                "name": "Power Output",
                "metric": "power_pct",
                "unit": "%",
                "state_class": "measurement",
                "precision": 1,
            },
            "Setpoint": {
                "name": "Setpoint",
                "metric": "setpoint_c",
                "unit": "°C",
                "device_class": "temperature",
                "state_class": "measurement",
                "precision": 1,
            },
            "LastUpdate": {
                "name": "Last Update",
                "metric": None,
                "device_class": "timestamp",
                "template": "{{ value_json.ts }}",
            },
        }

        if self.device_type == "furnace":
            defs.update({
                "Remaining": {
                    "name": "Remaining Time",
                    "metric": "remaining_time",
                },
                "RemainingSeconds": {
                    "name": "Remaining Seconds",
                    "metric": "remaining_seconds",
                    "unit": "s",
                    "device_class": "duration",
                    "precision": 0,
                },
                "PatternStep": {
                    "name": "Pattern Step",
                    "metric": "pattern_step",
                },
                "Pattern": {
                    "name": "Pattern",
                    "metric": "pattern",
                    "precision": 0,
                },
                "Step": {
                    "name": "Step",
                    "metric": "step",
                    "precision": 0,
                },
            })
        else:
            defs.update({
                "P": {"name": "PID P", "metric": "pid_p", "entity_category": "diagnostic", "enabled_by_default": False},
                "I": {"name": "PID I", "metric": "pid_i", "entity_category": "diagnostic", "enabled_by_default": False},
                "D": {"name": "PID D", "metric": "pid_d", "entity_category": "diagnostic", "enabled_by_default": False},
                "Cycle": {"name": "Cycle", "metric": "cycle", "entity_category": "diagnostic", "enabled_by_default": False},
                "Correction": {"name": "Correction", "metric": "correction", "entity_category": "diagnostic", "enabled_by_default": False},
                "Filter": {"name": "Filter", "metric": "filter", "entity_category": "diagnostic", "enabled_by_default": False},
                "OvertempAlarm": {"name": "Overtemperature Alarm", "metric": "overtemp_alarm", "entity_category": "diagnostic", "enabled_by_default": False},
            })
        return defs

    def publish_discovery(self):
        if not self.enable or not self.client:
            return

        device_info = self._device_info()
        for field, conf in self.sensor_definitions.items():
            object_id = f"{self.device_id}_{self._sanitize_id(field).lower()}"
            topic = f"{self.discovery_prefix}/sensor/{object_id}/config"
            payload = {
                "name": f"{self.entity_prefix_name} {conf['name']}",
                "state_topic": self.topic_pub,
                "value_template": conf.get("template") or f"{{{{ value_json.{field} }}}}",
                "unique_id": object_id,
                "availability_topic": self.availability_topic,
                "payload_available": "online",
                "payload_not_available": "offline",
                "force_update": True,
                "device": device_info,
            }
            if self.expire_after > 0:
                payload["expire_after"] = self.expire_after
            if conf.get("unit") is not None:
                payload["unit_of_measurement"] = conf.get("unit")
            if conf.get("device_class"):
                payload["device_class"] = conf.get("device_class")
            if conf.get("state_class"):
                payload["state_class"] = conf.get("state_class")
            if conf.get("precision") is not None:
                payload["suggested_display_precision"] = conf.get("precision")
            if conf.get("entity_category"):
                payload["entity_category"] = conf.get("entity_category")
            if conf.get("enabled_by_default") is not None:
                payload["enabled_by_default"] = bool(conf.get("enabled_by_default"))
            if conf.get("icon"):
                payload["icon"] = conf.get("icon")
            self._publish_json(topic, payload, retain=True)

        number_topic = f"{self.discovery_prefix}/number/{self.device_id}_setpoint/config"
        number_payload = {
            "name": f"{self.entity_prefix_name} Setpoint",
            "state_topic": self.topic_pub,
            "command_topic": self.setpoint_cmd_topic,
            "command_template": '{"Setpoint": {{ value | float }} }',
            "value_template": "{{ value_json.Setpoint }}",
            "unit_of_measurement": "°C",
            "min": 0,
            "max": 1200,
            "step": 1,
            "mode": "box",
            "unique_id": f"{self.device_id}_setpoint_number",
            "availability_topic": self.availability_topic,
            "payload_available": "online",
            "payload_not_available": "offline",
            "device": device_info,
        }
        self._publish_json(number_topic, number_payload, retain=True)
        self.publish_schema()

    def publish_schema(self):
        if not self.enable or not self.client:
            return

        fields = {}
        for field, conf in self.sensor_definitions.items():
            fields[field] = {
                "name": conf.get("name"),
                "metric": conf.get("metric"),
                "unit": conf.get("unit"),
                "device_class": conf.get("device_class"),
                "state_class": conf.get("state_class"),
            }
        payload = {
            "schema": self.SCHEMA_NAME,
            "schema_version": self.SCHEMA_VERSION,
            "device_type": self.device_type,
            "device_id": self.device_id,
            "device_name": self.device_name,
            "state_topic": self.topic_pub,
            "history_topic": self.history_topic or None,
            "command_topic": self.setpoint_cmd_topic,
            "availability_topic": self.availability_topic,
            "timestamp_field": "ts",
            "fields": fields,
        }
        self._publish_json(self.schema_topic, payload, retain=True)

    def _on_connect(self, client, userdata, flags, reason_code, properties=None):
        rc_value = getattr(reason_code, "value", reason_code)
        try:
            ok = int(rc_value) == 0
        except Exception:
            ok = str(reason_code).lower() in {"success", "0"}

        if ok:
            print("[MQTT] connected")
            client.publish(self.availability_topic, "online", qos=self.qos, retain=True)
            client.subscribe(self.setpoint_cmd_topic, qos=self.qos)
            client.subscribe(self.ha_status_topic, qos=self.qos)
            self.publish_discovery()
            if self._last_payload:
                self._publish_json(self.topic_pub, self._last_payload, retain=self.retain)
        else:
            print(f"[MQTT] connect rc={reason_code}")

    def _on_message(self, client, userdata, msg):
        raw = msg.payload.decode(errors="replace").strip()
        try:
            if msg.topic == self.ha_status_topic:
                if raw.lower() == "online":
                    self.publish_discovery()
                    if self._last_payload:
                        self._publish_json(self.topic_pub, self._last_payload, retain=self.retain)
                return

            if msg.topic != self.setpoint_cmd_topic:
                return

            new_sv = self._parse_setpoint(raw)
            if new_sv is None:
                print(f"[MQTT] ignored setpoint payload: {raw}")
                return

            print(f"[MQTT] received new setpoint {new_sv}")
            if self.gui_ref:
                self.gui_ref.apply_remote_setpoint(new_sv)
        except Exception as e:
            print(f"[MQTT] message error: {e}")

    def _parse_setpoint(self, raw):
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict):
                for key in ("Setpoint", "SetTemperature", "setpoint", "value"):
                    if key in obj:
                        return float(obj[key])
                return None
            return float(obj)
        except Exception:
            try:
                return float(raw)
            except Exception:
                return None

    def _utc_now(self):
        try:
            now = datetime.datetime.now(datetime.timezone.utc)
        except AttributeError:
            now = datetime.now(timezone.utc)
        iso = now.isoformat(timespec="milliseconds").replace("+00:00", "Z")
        return iso, now.timestamp()

    def _clean_json(self, value):
        import math
        if isinstance(value, dict):
            return {str(k): self._clean_json(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [self._clean_json(v) for v in value]
        if isinstance(value, float):
            if math.isnan(value) or math.isinf(value):
                return None
            return value
        try:
            if isinstance(value, (datetime.datetime, datetime.date)):
                return value.isoformat()
        except AttributeError:
            if isinstance(value, datetime):
                return value.isoformat()
        return value

    def _build_payload(self, payload_dict):
        base = self._clean_json(dict(payload_dict or {}))
        if "Setpoint" not in base and "SetTemperature" in base:
            base["Setpoint"] = base.get("SetTemperature")

        ts_iso, unix_ts = self._utc_now()
        self._sequence += 1
        out = {
            "schema": self.SCHEMA_NAME,
            "schema_version": self.SCHEMA_VERSION,
            "device_type": self.device_type,
            "device_id": self.device_id,
            "device_name": self.device_name,
            "sequence": self._sequence,
            "ts": ts_iso,
            "unix_time": round(unix_ts, 3),
        }
        out.update(base)

        metrics = {}
        for field, conf in self.sensor_definitions.items():
            metric_name = conf.get("metric")
            if not metric_name or field not in out:
                continue
            metrics[metric_name] = out.get(field)
        out["metrics"] = metrics
        return self._clean_json(out)

    def _publish_json(self, topic, payload, retain=False):
        if not self.enable or not self.client:
            return
        try:
            text = json.dumps(payload, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
            self.client.publish(topic, text, qos=self.qos, retain=bool(retain))
        except Exception as e:
            print(f"[MQTT] publish error on {topic}: {e}")

    def publish(self, payload_dict, *, force=False):
        if not self.enable or not self.client:
            return

        now = time.monotonic()
        if not force and self.publish_interval > 0:
            if now - self._last_publish_monotonic < self.publish_interval:
                return
        self._last_publish_monotonic = now

        payload = self._build_payload(payload_dict)
        self._last_payload = payload
        self._publish_json(self.topic_pub, payload, retain=self.retain)
        if self.history_topic:
            self._publish_json(self.history_topic, payload, retain=False)

    def close(self):
        if not self.client:
            return
        try:
            self.client.publish(self.availability_topic, "offline", qos=self.qos, retain=True)
            self.client.loop_stop()
            self.client.disconnect()
        except Exception as e:
            print(f"[MQTT] close error: {e}")


API_PORT = cfg.getint("General", "api_port", fallback=5000)   # ← NEW
ACCENT_COLOR = "#2E7D32"                    # fresh green accent
WHITE_BG   = "#FFFFFF"                 # uniform background
POP_FONT   = ("Arial", max(8, BASE_FONT_SIZE - 7))   # smaller than before
BTN_FONT = ("Arial", 14)     # ← change 14 to any size you like
class TemperatureControlApp:
    def __init__(self, master):
        self.cfg = cfg

        self.mqtt_mgr = None                # will hold MQTTManager instance
        self.master = master
        self.master.title("ONWAY TEMPERATURE CONTROLLER")
        self.master.geometry("900x370")
        if ICON_PATH and os.path.exists(ICON_PATH):
            try:
                self.master.iconbitmap(ICON_PATH)
            except Exception:
                pass           # ignore if .ico not valid for this platform

        # state & threads
        self.client = None
        self.running = True
        self.api_thread = None
        self.modbus_thread = None
        # shared data controls
        self.set_point_var = tk.StringVar()
        self.com_var       = tk.StringVar(value=SERIAL_OPTS["port"])
        self.api_on_var    = tk.BooleanVar(value=self.cfg.getboolean("General", "api_enabled", fallback=True))
        self.data_save_var = tk.BooleanVar(value=self.cfg.getboolean("General", "save_logs",  fallback=True))
        self.port_var      = tk.StringVar(value=str(API_PORT))   
        self.log_dir       = tk.StringVar(value=self.cfg.get("Logging", "directory",
                                                            fallback=os.getcwd()))

        # plotting buffers
        self.times = []
        self.temps = []
        self.last_plot_ts = datetime.min
        # build the UI
        self._build_layout()
        self._build_left_panel()
        self._build_config_panel()
        self._build_chart_panel()
        self._bind_keys()
        self._schedule_ui_refresh()

    # ───────────────────────────────────────────────
    #  Configuration dialog – white, compact, accent
    # ───────────────────────────────────────────────
    def show_config_dialog(self):
        dlg = tk.Toplevel(self.master, bg="#FFFFFF")
        dlg.title("Configuration")
        dlg.transient(self.master); dlg.grab_set()
        if ICON_PATH and os.path.exists(ICON_PATH):
            try: dlg.iconbitmap(ICON_PATH)
            except Exception: pass

        # vars
        com_var  = tk.StringVar(value=self.com_var.get())
        port_var = tk.StringVar(value=self.port_var.get())
        api_var  = tk.BooleanVar(value=self.api_on_var.get())
        save_var = tk.BooleanVar(value=self.data_save_var.get())
        dir_var  = tk.StringVar(value=self.log_dir.get())

        row = 0
        def label(txt):
            tk.Label(dlg, text=txt, font=POP_FONT, bg="#FFFFFF", fg="#000000")\
            .grid(row=row, column=0, sticky="e", padx=6, pady=3)

        # COM port
        label("COM Port:")
        tk.Entry(dlg, textvariable=com_var, font=POP_FONT, width=12)\
            .grid(row=row, column=1, sticky="w", padx=6, pady=3)

        # API port
        row += 1; label("API Port:")
        tk.Entry(dlg, textvariable=port_var, font=POP_FONT, width=10)\
            .grid(row=row, column=1, sticky="w", padx=6, pady=3)

        # check-boxes
        row += 1
        tk.Checkbutton(dlg, text="Enable API", variable=api_var,
                    font=POP_FONT, bg="#FFFFFF", fg="#000000",
                    selectcolor="#FFFFFF", activebackground="#FFFFFF")\
            .grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=3)
        row += 1
        tk.Checkbutton(dlg, text="Save Logs", variable=save_var,
                    font=POP_FONT, bg="#FFFFFF", fg="#000000",
                    selectcolor="#FFFFFF", activebackground="#FFFFFF")\
            .grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=3)

        # log directory
        row += 1; label("Log Directory:")
        tk.Entry(dlg, textvariable=dir_var, font=POP_FONT, width=28)\
            .grid(row=row, column=1, sticky="w", padx=6, pady=3)
        tk.Button(dlg, text="Browse", font=POP_FONT,
                command=lambda: self._browse_dir(dir_var))\
            .grid(row=row, column=2, sticky="w", padx=6, pady=3)

        # buttons
        row += 1
        btn_frm = tk.Frame(dlg, bg="#FFFFFF")
        btn_frm.grid(row=row, column=0, columnspan=3, pady=8)
        tk.Button(btn_frm, text="Apply", font=POP_FONT, width=6,
                bg=ACCENT_COLOR, fg="#FFFFFF",
                command=lambda: self._apply_cfg_changes(
                    dlg, com_var, port_var, api_var, save_var, dir_var))\
            .pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frm, text="Cancel", font=POP_FONT, width=6,
                command=dlg.destroy)\
            .pack(side=tk.LEFT, padx=8)

    def _browse_dir(self, var):
        folder = filedialog.askdirectory(initialdir=var.get() or os.getcwd())
        if folder: var.set(folder)

    def _apply_cfg_changes(self, dlg, com_v, port_v, api_v, save_v, dir_v):
        # live update
        self.com_var.set(com_v.get().strip())
        self.port_var.set(port_v.get().strip() or "5000")
        self.api_on_var.set(api_v.get())
        self.data_save_var.set(save_v.get())
        self.log_dir.set(dir_v.get())

        # persist to INI
        self.cfg["Serial"]["port"]         = self.com_var.get()
        self.cfg["General"]["api_port"]    = self.port_var.get()
        self.cfg["General"]["api_enabled"] = str(self.api_on_var.get()).lower()
        self.cfg["General"]["save_logs"]   = str(self.data_save_var.get()).lower()
        self.cfg["Logging"]["directory"]   = self.log_dir.get()
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            self.cfg.write(f)

        dlg.destroy()


    def _schedule_ui_refresh(self):
        self._refresh_readouts()
        if self.running:
            # 100ms later, schedule again on the main thread
            self.master.after(250, self._schedule_ui_refresh)


    def _build_layout(self):
        main = tk.Frame(self.master, bg=WHITE_BG, bd=1, relief="solid")
        main.pack(fill="both", expand=True)

        self.left_frame  = tk.Frame(main,  bg=WHITE_BG)
        self.right_frame = tk.Frame(main,  bg=WHITE_BG, bd=1, relief="solid")

        self.left_frame.pack(side=tk.LEFT,  fill="y",   padx=10, pady=10)
        self.right_frame.pack(side=tk.RIGHT, fill="both", expand=True,
                            padx=10, pady=10)

    def _build_left_panel(self):
        # --- Temperature Display ---
        temp_f = tk.Frame(self.left_frame, bg=WHITE_BG)
        temp_f.pack(pady=10)
        self.temp_label_number = tk.Label(temp_f, text="--", font=BIG_FONT,  bg=WHITE_BG)
        self.temp_label_number.pack(side=tk.LEFT)
        self.temp_label_unit = tk.Label(temp_f, text="°C", font=UNIT_FONT,  bg=WHITE_BG)
        self.temp_label_unit.pack(side=tk.LEFT, padx=(10, 0))
        # --- Power Bar ---
        pow_f = tk.Frame(self.left_frame,  bg=WHITE_BG)
        pow_f.pack(pady=10)
        tk.Label(pow_f, text="Power:", font=BASE_FONT,  bg=WHITE_BG).pack(side=tk.LEFT)
        self.power_bar = ttk.Progressbar(pow_f, orient="horizontal", length=100, mode="determinate")
        self.power_bar["maximum"] = 100
        self.power_bar.pack(side=tk.LEFT, padx=5)
        self.power_percent_label = tk.Label(pow_f, text="--%", font=ENTRY_FONT,  bg=WHITE_BG)
        self.power_percent_label.pack(side=tk.LEFT, padx=5)
        # --- SetPoint Entry ---
        sp_f = tk.Frame(self.left_frame,  bg=WHITE_BG)
        sp_f.pack(pady=10)
        tk.Label(sp_f, text="SetPoint:", font=BASE_FONT,  bg=WHITE_BG).pack(side=tk.LEFT, padx=5)
        self.set_point_entry = tk.Entry(
            sp_f, textvariable=self.set_point_var,
            font=ENTRY_FONT, width=8, justify="center"
        )
        self.set_point_entry.pack(side=tk.LEFT)
        self.set_point_entry.bind("<Return>", self.on_enter_set_point)
        self.entry_typing = False
        self.set_point_entry.bind("<Key>",    self._on_entry_keypress)
        self.set_point_entry.bind("<FocusOut>", self._on_entry_focus_out)
        tk.Label(sp_f, text="°C", font=BASE_FONT,  bg=WHITE_BG).pack(side=tk.LEFT, padx=5)

    def _on_entry_keypress(self, event):
        self.entry_typing = True

    def _on_entry_focus_out(self, event):
        self.entry_typing = False

    def _build_config_panel(self):
        cfg = tk.Frame(self.left_frame, bg=WHITE_BG)
        cfg.pack(pady=10, fill="x", anchor="center")
        cfg.pack(pady=10, fill="x")

        std_btn = dict(
            bg=WHITE_BG,
            fg=ACCENT_COLOR,
            font=BTN_FONT,
            bd=1, relief="solid",            # thin black outline
            highlightthickness=1,
            highlightbackground="#000000"
        )

        tk.Button(cfg, text="Connect",      **std_btn,
                command=self.on_init_controller)\
        .grid(row=0, column=0, padx=8, pady=4)

        tk.Button(cfg, text="Configuration", **std_btn,
                command=self.show_config_dialog)\
        .grid(row=0, column=1, padx=8, pady=4)


    def _build_chart_panel(self):
        self.fig, self.ax = plt.subplots(figsize=(4, 3))
        self.ax.set_xlabel("Time");  self.ax.set_ylabel("Temperature (°C)")
        (self.line,) = self.ax.plot([], [], marker="o", markersize=2, linestyle="-")

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.right_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.right_frame)
        self.toolbar.update()

        def _on_move(event):
            if event.inaxes == self.ax and event.xdata and event.ydata:
                self.master.title(f"T = {event.ydata:0.2f} °C   "
                                f"t = {mdates.num2date(event.xdata).strftime('%H:%M:%S')}")
            else:
                self.master.title("ONWAY TEMPERATURE CONTROLLER")
        self.fig.canvas.mpl_connect('motion_notify_event', _on_move)


    def _bind_keys(self):
        # keyboard ↑/↓ bump bindings
        self.master.bind("<KeyPress-plus>",   self.on_key_up)
        self.master.bind("<KeyPress-minus>", self.on_key_down)
    
    def _start_ui_thread(self):
        """Spawn GUI/MQTT refresh loop in its own daemon thread."""
        self.ui_thread = threading.Thread(
            target=self.ui_update_loop,
            daemon=True
        )
        self.ui_thread.start()

    def on_init_controller(self):
        """Connect to the controller, then launch Modbus, API, and MQTT threads."""
        # ── 1.  open serial link ──────────────────────────────────────────
        port = self.com_var.get().strip() or "COM4"
        if not self._connect_modbus(port):
            print(f"[Init] ❌  Failed to connect on {port}")
            return

        # ── 2.  read existing set-point once ─────────────────────────────
        initial_sp = read_register(client, 3000)
        if initial_sp is not None:
            data_store["SetTemperature"] = initial_sp
            self.set_point_var.set(f"{initial_sp:.1f}")

        # ── 3.  make sure log directory exists ───────────────────────────
        self._ensure_log_dir(self.log_dir.get())
        self._start_ui_thread()
        # ── 4.  start background threads ─────────────────────────────────
        self._start_modbus_thread()

        if self.api_on_var.get():
            try:
                port_num = int(self.port_var.get() or 5000)
            except ValueError:
                port_num = 5000
                self.port_var.set(str(port_num))
            self._start_api_thread(port_num)

        # ── 5.  bring up MQTT discovery & telemetry ──────────────────────
        self._init_mqtt()

    def _connect_modbus(self, port_name):
        """Open serial link using parameters from INI; return True on success."""
        new_client = ModbusClient(
            port     = port_name,
            baudrate = SERIAL_OPTS["baudrate"],
            parity   = SERIAL_OPTS["parity"],
            stopbits = SERIAL_OPTS["stopbits"],
            bytesize = SERIAL_OPTS["bytesize"],
            timeout  = 3,
        )
        if new_client.connect():
            global client, stop_threads
            client = new_client
            stop_threads = False
            return True
        return False

    def _ensure_log_dir(self, path):
        """Create log directory if it doesn’t exist."""
        if not os.path.isdir(path):
            os.makedirs(path)
    def _start_modbus_thread(self):
        """Spawn thread polling Modbus values."""
        self.modbus_thread = threading.Thread(
            target=update_modbus_values_loop,
            args=(self,),
            daemon=True
        )
        self.modbus_thread.start()
    def _start_api_thread(self, port_num):
        """Spawn Flask API in its own thread."""
        self.api_thread = threading.Thread(
            target=run_flask_app,
            args=(port_num,),
            daemon=True
        )
        self.api_thread.start()  

    def on_enter_set_point(self, event):
        self.entry_typing = False
        try:
            val = float(self.set_point_var.get().strip())
        except ValueError:
            return
        data_store["SetTemperature"] = val
        threading.Thread(
            target=write_register,
            args=(client, 3000, val),
            daemon=True
        ).start()

    def browse_log_dir(self):
        """Open a folder dialog and update the log directory."""
        folder = filedialog.askdirectory()
        if folder:
            self.log_dir.set(folder)
    # —— keyboard bump helpers ——
    def _bump_set_point(self, delta):
        cur = data_store.get("SetTemperature")
        if cur is None:
            return

        new_val = round(cur + delta, 1)
        data_store["SetTemperature"] = new_val
        # Send to machine, non-blocking
        threading.Thread(
            target=write_register,
            args=(client, 3000, new_val),
            daemon=True
        ).start()

        # (Optional) update your entry immediately so it "feels" instant,
        # but don’t store it in data_store.
        self.set_point_var.set(f"{new_val:.1f}")


    def on_key_up(self, event):
        self._bump_set_point(+1)
    def on_key_down(self, event):
        self._bump_set_point(-1)
    # ---------------------------------------------------------
    def ui_update_loop(self):
        try:
            while not stop_threads and self.running:
                if self.mqtt_mgr:
                    payload = {
                        "Temperature": data_store.get("Temperature"),
                        "Power": data_store.get("Power"),
                        "Setpoint": data_store.get("SetTemperature"),
                        "SetTemperature": data_store.get("SetTemperature"),
                        "PowerLimit": data_store.get("PowerLimit"),
                        "Segment": data_store.get("Segment"),
                        "SegmentLeft": data_store.get("SegmentLeft"),
                        "P": data_store.get("P"),
                        "I": data_store.get("I"),
                        "D": data_store.get("D"),
                        "Cycle": data_store.get("Cycle"),
                        "Correction": data_store.get("Correction"),
                        "Filter": data_store.get("Filter"),
                        "OvertempAlarm": data_store.get("OvertempAlarm"),
                        "ManualOverride": data_store.get("manual_override", False),
                    }
                    self.mqtt_mgr.publish(payload)
                time.sleep(0.1)
        except Exception as e:
            print("[MQTT] telemetry loop stopped:", e)

    def _refresh_readouts(self):
            temp  = data_store.get("Temperature")
            power = data_store.get("Power")
            spt   = data_store.get("SetTemperature")

            self._update_temp_label(temp)
            self._update_power_display(power)
            self._update_setpoint_entry(spt)

            if temp is not None:
                self.add_temp_to_plot(temp)

    def _update_temp_label(self, temp):
        text = f"{temp:.1f}" if temp is not None else "--"
        self.temp_label_number.config(text=text)

    def _update_power_display(self, power):
        if power is not None:
            p_val = max(0, min(100, power))
            self.power_bar["value"] = p_val
            self.power_percent_label.config(text=f"{power:.1f}%")
        else:
            self.power_bar["value"] = 0
            self.power_percent_label.config(text="--%")

    def _update_setpoint_entry(self, spt):
        # only overwrite if user isn’t actively typing
        if not self.entry_typing and spt is not None:
            self.set_point_var.set(f"{spt:.1f}")

    def add_temp_to_plot(self, temp):
        now = datetime.now()
        if (now - self.last_plot_ts).total_seconds() < 1.0:
            return                           # ← too soon, skip
        self.last_plot_ts = now

        self.times.append(now)
        self.temps.append(temp)

        # keep only last 30 min of data
        cutoff = now - timedelta(minutes=30)
        while self.times and self.times[0] < cutoff:
            self.times.pop(0); self.temps.pop(0)

        self.line.set_data(self.times, self.temps)
        self.ax.relim()                 # update data limits—but…
        if self.toolbar.mode == "":     # …only autoscale when user NOT panning/zooming
            self.ax.autoscale_view()    # preserves manual zoom until Home is pressed

        self.ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        self.ax.tick_params(axis='x', rotation=45)
        self.fig.tight_layout()
        self.canvas.draw_idle()



    # ───────── MQTT integration ─────────
    def _init_mqtt(self):
        if not self.cfg.getboolean("MQTT", "enabled", fallback=False):
            return
        topic_pub = self.cfg["MQTT"]["topic"]
        self.mqtt_mgr = MQTTManager(
            host      = self.cfg["MQTT"]["host"],
            port      = self.cfg.getint("MQTT", "port"),
            topic_pub = topic_pub,
            client_id = self.cfg["MQTT"].get("client_id", ""),
            username  = self.cfg["MQTT"].get("username", ""),
            password  = self.cfg["MQTT"].get("password", ""),
            setpoint_cmd_topic = self.cfg["MQTT"].get("setpoint_cmd_topic", f"{topic_pub}/set"),
            discovery_prefix   = self.cfg["MQTT"].get("discovery_prefix", "homeassistant"),
            qos     = self.cfg.getint("MQTT", "qos", fallback=0),
            retain  = self.cfg.getboolean("MQTT", "retain", fallback=False),
            device_type = "onway",
            device_id   = self.cfg.get("Device", "identifiers", fallback="onway_tempctl"),
            device_name = self.cfg.get("Device", "name", fallback="Onway Temperature Controller"),
            manufacturer = self.cfg.get("Device", "manufacturer", fallback="ONWAY"),
            model = self.cfg.get("Device", "model", fallback="OTC-9600"),
            entity_prefix_name = self.cfg.get("MQTT", "entity_prefix_name", fallback="Onway"),
            publish_interval = self.cfg.getfloat("MQTT", "publish_interval_sec", fallback=1.0),
            expire_after = self.cfg.getint("MQTT", "expire_after", fallback=300),
            history_topic = self.cfg["MQTT"].get("history_topic", ""),
            gui_ref = self
        )

    # called by MQTTManager when a remote user changes the set-point in HA
    def apply_remote_setpoint(self, new_sv):
        data_store["SetTemperature"] = new_sv
        self.set_point_var.set(f"{new_sv:.1f}")
        threading.Thread(
            target=write_register,
            args=(client, 3000, new_sv),
            daemon=True
        ).start()

    def close(self):
        """Cleanup threads, close client, then exit."""
        global stop_threads
        stop_threads = True
        self.running = False
        try: self.master.after_cancel(self._schedule_ui_refresh)  # no stray after
        except Exception: pass
        if client:
            client.close()
        if self.mqtt_mgr:
            self.mqtt_mgr.close()
        self.master.destroy()
        import sys
        sys.exit(0)

### CHANGED: No other modifications

if __name__ == "__main__":
    root = tk.Tk()
    gui_app = TemperatureControlApp(root)       # ← GOOD: keeps names distinct
    root.protocol("WM_DELETE_WINDOW", gui_app.close)
    root.mainloop()