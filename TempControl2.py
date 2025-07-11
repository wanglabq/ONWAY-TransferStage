import os
import time
import csv
from datetime import datetime, timedelta
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

BASE_FONT = ("Arial", 20)          # was 18 → now 20
BASE_FONT_BOLD = ("Arial", 22, "bold")  # was 20 bold (e.g. 18+2)
BIG_FONT = ("Arial", 82, "bold")   # was 80 → now 82
UNIT_FONT = ("Arial", 32)          # was 30 → now 32
ENTRY_FONT = ("Arial", 18)         # was 16 → now 18

# Configuration
SECRET_TOKEN = "Wanglabisestablishedon20240101inUIUCPHYSICS"
LOG_ENCODING = "utf-8-sig"
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
def read_register(client, address, slave=10, retries=5):
    for attempt in range(retries):
        response = client.read_holding_registers(address, count=1, slave=slave)
        if not response.isError():
            value = response.registers[0] / 10
            # these registers need the *10 rescale
            if address in [2036, 18523, 2092, 18506, 18507, 18508, 18509, 18501]:
                return value * 10
            return value
        time.sleep(0.1)
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

class TemperatureControlApp:
    def __init__(self, master):
        self.master = master
        self.master.title("ONWAY TEMPERATURE CONTROLLER")
        self.master.geometry("900x440")
        # state & threads
        self.client = None
        self.running = True
        self.api_thread = None
        self.modbus_thread = None
        # shared data controls
        self.set_point_var = tk.StringVar()
        self.com_var = tk.StringVar(value="COM4")
        self.api_on_var = tk.BooleanVar(value=True)
        self.port_var = tk.StringVar(value="5000")
        self.data_save_var = tk.BooleanVar(value=True)
        self.log_dir = tk.StringVar(value=r"C:\Users\WangLabAdmin\Desktop\TransferLog")
        # plotting buffers
        self.times = []
        self.temps = []
        # build the UI
        self._build_layout()
        self._build_left_panel()
        self._build_config_panel()
        self._build_chart_panel()
        self._bind_keys()
        self._schedule_ui_refresh()

    def _schedule_ui_refresh(self):
        self._refresh_readouts()
        if self.running:
            # 100ms later, schedule again on the main thread
            self.master.after(100, self._schedule_ui_refresh)

    def _build_layout(self):
        # main frames
        main = tk.Frame(self.master)
        main.pack(fill="both", expand=True)
        self.left_frame = tk.Frame(main)
        self.left_frame.pack(side=tk.LEFT, fill="y", padx=10, pady=10)
        self.right_frame = tk.Frame(main)
        self.right_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=10, pady=10)
    def _build_left_panel(self):
        # --- Temperature Display ---
        temp_f = tk.Frame(self.left_frame)
        temp_f.pack(pady=10)
        self.temp_label_number = tk.Label(temp_f, text="--", font=BIG_FONT)
        self.temp_label_number.pack(side=tk.LEFT)
        self.temp_label_unit = tk.Label(temp_f, text="°C", font=UNIT_FONT)
        self.temp_label_unit.pack(side=tk.LEFT, padx=(10, 0))
        # --- Power Bar ---
        pow_f = tk.Frame(self.left_frame)
        pow_f.pack(pady=10)
        tk.Label(pow_f, text="Power:", font=BASE_FONT).pack(side=tk.LEFT)
        self.power_bar = ttk.Progressbar(pow_f, orient="horizontal", length=100, mode="determinate")
        self.power_bar["maximum"] = 100
        self.power_bar.pack(side=tk.LEFT, padx=5)
        self.power_percent_label = tk.Label(pow_f, text="--%", font=ENTRY_FONT)
        self.power_percent_label.pack(side=tk.LEFT, padx=5)
        # --- SetPoint Entry ---
        sp_f = tk.Frame(self.left_frame)
        sp_f.pack(pady=10)
        tk.Label(sp_f, text="SetPoint:", font=BASE_FONT).pack(side=tk.LEFT, padx=5)
        self.set_point_entry = tk.Entry(
            sp_f, textvariable=self.set_point_var,
            font=ENTRY_FONT, width=8, justify="center"
        )
        self.set_point_entry.pack(side=tk.LEFT)
        self.set_point_entry.bind("<Return>", self.on_enter_set_point)
        self.entry_typing = False
        self.set_point_entry.bind("<Key>",    self._on_entry_keypress)
        self.set_point_entry.bind("<FocusOut>", self._on_entry_focus_out)
        tk.Label(sp_f, text="°C", font=BASE_FONT).pack(side=tk.LEFT, padx=5)

    def _on_entry_keypress(self, event):
        self.entry_typing = True

    def _on_entry_focus_out(self, event):
        self.entry_typing = False

    def _build_config_panel(self):
        cfg = tk.Frame(self.left_frame, bg="#F0F0F0", bd=2, relief=tk.GROOVE)
        cfg.pack(pady=10, fill="x")
        # row 0
        tk.Label(cfg, text="COM Port:", font=("Arial", 12), bg="#F0F0F0")\
            .grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(cfg, textvariable=self.com_var, font=("Arial", 12), width=10)\
            .grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Button(cfg, text="Connect", font=("Arial", 12),
                  command=self.on_init_controller)\
            .grid(row=0, column=2, padx=10, pady=5, sticky="w")
        # row 1
        tk.Checkbutton(cfg, text="API On", variable=self.api_on_var,
                       font=("Arial",12), bg="#F0F0F0")\
            .grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Label(cfg, text="Port:", font=("Arial",12), bg="#F0F0F0")\
            .grid(row=1, column=1, padx=5, pady=5, sticky="e")
        tk.Entry(cfg, textvariable=self.port_var,
                 font=("Arial",12), width=6)\
            .grid(row=1, column=2, padx=5, pady=5, sticky="w")
        # row 2
        tk.Checkbutton(cfg, text="Save Data", variable=self.data_save_var,
                       font=("Arial",12), bg="#F0F0F0")\
            .grid(row=2, column=0, padx=5, pady=5, sticky="w")
        tk.Label(cfg, text="Log Dir:", font=("Arial",12), bg="#F0F0F0")\
            .grid(row=2, column=1, padx=5, pady=5, sticky="e")
        tk.Entry(cfg, textvariable=self.log_dir,
                 font=("Arial",12), width=20)\
            .grid(row=2, column=2, padx=5, pady=5, sticky="w")
        tk.Button(cfg, text="Browse", command=self.browse_log_dir)\
            .grid(row=2, column=3, padx=5, pady=5, sticky="w")
    def _build_chart_panel(self):
        # Matplotlib chart
        self.fig, self.ax = plt.subplots(figsize=(4,3))
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Temperature (°C)")
        (self.line,) = self.ax.plot([], [], marker="o", markersize=2, linestyle="-")
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.right_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
    def _bind_keys(self):
        # keyboard ↑/↓ bump bindings
        self.master.bind("<KeyPress-plus>",   self.on_key_up)
        self.master.bind("<KeyPress-minus>", self.on_key_down)

    def on_init_controller(self):
        """Handler for the “Connect” button: set up Modbus, logging, and all background threads."""
        port = self.com_var.get().strip()
        if not self._connect_modbus(port):
            print(f"[Init] Failed to connect on {port}")
            return
        # --- Add this: one-time read of existing setpoint ---
        initial_sp = read_register(client, 3000)
        if initial_sp is not None:
            data_store["SetTemperature"] = initial_sp
            self.set_point_var.set(f"{initial_sp:.1f}")

        self._ensure_log_dir(self.log_dir.get())
        self._start_modbus_thread()
        self._ensure_log_dir(self.log_dir.get())
        self._start_modbus_thread()
        if self.api_on_var.get():
            self._start_api_thread(int(self.port_var.get().strip() or 5000))
    def _connect_modbus(self, port_name):
        """Attempt serial connection; return True on success."""
        new_client = ModbusClient(
            port=port_name, baudrate=9600,
            parity='N', stopbits=1,
            bytesize=8, timeout=3
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
        self.api_thread.start

    def _start_ui_thread(self):
        pass   

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
        """Background thread: update all readouts & chart every 0.5s."""
        while not stop_threads and self.running:
            self._refresh_readouts()
            time.sleep(0.1)

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
        # original logic untouched
        now_dt = datetime.now()
        self.times.append(now_dt)
        self.temps.append(temp)

        cutoff = now_dt - timedelta(minutes=30)
        while self.times and self.times[0] < cutoff:
            self.times.pop(0)
            self.temps.pop(0)

        self.ax.clear()
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Temperature (°C)")
        self.ax.plot(self.times, self.temps, marker="o", markersize=2, linestyle="-")

        self.ax.set_xlim([cutoff, now_dt])
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        self.ax.tick_params(axis='x', rotation=45)
        self.fig.tight_layout()
        self.canvas.draw()

    def close(self):
        """Cleanup threads, close client, then exit."""
        global stop_threads
        stop_threads = True
        self.running = False
        if client:
            client.close()
        self.master.destroy()
        import sys
        sys.exit(0)

### CHANGED: No other modifications

if __name__ == "__main__":
    root = tk.Tk()
    app = TemperatureControlApp(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()