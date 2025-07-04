import os
import time
import csv
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime, timedelta

from pymodbus.client import ModbusSerialClient as ModbusClient
from pymodbus.exceptions import ModbusIOException

# ------------------ Flask & API ------------------
from flask import Flask, request, jsonify
from flask_restful import Api, Resource

app = Flask(__name__)
api = Api(app)

SECRET_TOKEN = "Wanglabisestablishedon20240101inUIUCPHYSICS"

def check_token():
    token = request.headers.get("Authorization")
    if token is None:
        return True
    return token == SECRET_TOKEN

# ------------------ Modbus Read/Write ------------------
def read_register(client, address, slave=10, retries=5):
    for attempt in range(retries):
        response = client.read_holding_registers(address, count=1, slave=slave)
        if not response.isError():
            value = response.registers[0] / 10
            if address in [2036, 18523, 2092, 18506, 18507, 18508, 18509, 18501]:
                return value * 10
            return value
        time.sleep(1)
    return None

def write_register(client, address, value, slave=10):
    if address in [18506, 18507, 18508, 18509, 18523, 18501, 2036]:
        value = value / 10
    client.write_register(address, int(value * 10), slave=slave)

# ------------------ Shared Data Store ------------------
data_store = {
    "Temperature": None,
    "SetTemperature": None,
    "Power": None,
    "PowerLimit": None,
    "Segment": None,
    "SegmentLeft": None,
    "P": None,
    "I": None,
    "D": None,
    "Cycle": None,
    "Correction": None,
    "Filter": None,
    "OvertempAlarm": None,
    "manual_override": False,
}

last_param_values = {
    "P": None,
    "I": None,
    "D": None,
    "Cycle": None,
    "Correction": None,
    "Filter": None,
    "OvertempAlarm": None
}

client = None  # We'll assign after init
stop_threads = False

# ------------------ Logging Helpers ------------------
def get_current_time():
    now = datetime.now()
    return (
        now.strftime("%Y-%m-%d %H:%M:%S"),
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        now.strftime("%Y-%m")
    )

def log_temp_data(log_dir, temperature, power):
    ts, date_str, time_str, month_str = get_current_time()
    filename = os.path.join(log_dir, f"TEMP_LOG_{month_str}.csv")
    file_exists = os.path.isfile(filename)

    try:
        with open(filename, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "Date", "Time", "Temperature", "Power"])
            writer.writerow([ts, date_str, time_str, temperature, power])

    except PermissionError:
        print(f"[Warning] Cannot write to '{filename}' – file is open in Excel?")

def log_param_changes(log_dir, p, i, d, cyc, cor, fil, ovt):
    ts, date_str, time_str, month_str = get_current_time()
    filename = os.path.join(log_dir, f"PARAMETER_LOG_{month_str}.csv")
    file_exists = os.path.isfile(filename)

    try:
        with open(filename, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow([
                    "Timestamp","Date","Time",
                    "P","I","D","Cycle","Correction","Filter","Overtemp_alarm"
                ])
            writer.writerow([
                ts, date_str, time_str,
                p, i, d, cyc, cor, fil, ovt
            ])
    except PermissionError:
        print(f"[Warning] Cannot write to '{filename}' – file is open in Excel?")

def update_modbus_values_loop(app_ref):
    """Periodic read from 18504 => Temp, 3000 => SetTemp, 2036 => Power, etc."""
    global stop_threads, client, data_store
    while not stop_threads and app_ref.running:
        # Read temperature
        temp = read_register(client, 18504)

        # Read set temp if not override
        if data_store["manual_override"]:
            st = data_store["SetTemperature"]
        else:
            st = read_register(client, 3000)

        # Power
        pow_ = read_register(client, 2036)

        # Additional param
        p_ = read_register(client, 18506)
        i_ = read_register(client, 18507)
        d_ = read_register(client, 18508)
        cyc = read_register(client, 18509)
        cor = read_register(client, 18550)
        fil = read_register(client, 18501)
        ovt = read_register(client, 2490)

        data_store["Temperature"] = temp
        data_store["SetTemperature"] = st
        data_store["Power"] = pow_
        data_store["P"] = p_
        data_store["I"] = i_
        data_store["D"] = d_
        data_store["Cycle"] = cyc
        data_store["Correction"] = cor
        data_store["Filter"] = fil
        data_store["OvertempAlarm"] = ovt

        # If save data => log
        if app_ref.data_save_var.get():
            if (temp is not None) and (pow_ is not None):
                log_temp_data(app_ref.log_dir.get(), temp, pow_)

            # param changes
            changed = False
            new_params = {
                "P": p_,
                "I": i_,
                "D": d_,
                "Cycle": cyc,
                "Correction": cor,
                "Filter": fil,
                "OvertempAlarm": ovt
            }
            global last_param_values
            for name, val in new_params.items():
                old_val = last_param_values[name]
                if val is not None and old_val != val:
                    changed = True
            if changed:
                log_param_changes(
                    app_ref.log_dir.get(),
                    new_params["P"], new_params["I"], new_params["D"],
                    new_params["Cycle"], new_params["Correction"],
                    new_params["Filter"], new_params["OvertempAlarm"]
                )
                for n, v in new_params.items():
                    last_param_values[n] = v

        time.sleep(1)

# ------------------ API ------------------
class DataAPI(Resource):
    def get(self):
        return data_store

    def post(self):
        """SetTemperature => 3000"""
        new_value = request.json.get("SetTemperature")
        if new_value is not None:
            write_register(client, 3000, float(new_value))
            data_store["SetTemperature"] = float(new_value)
            return {"message": "Set Temperature updated", "SetTemperature": new_value}
        return {"error": "Invalid input"}, 400

api.add_resource(DataAPI, "/api/data")

class SetpointAPI(Resource):
    def post(self):
        """Writes to 3000 => manual_override=True"""
        if not check_token():
            return {"error": "Unauthorized"}, 401
        new_value = request.json.get("SetTemperature")
        if new_value is not None:
            write_register(client, 3000, float(new_value))
            data_store["SetTemperature"] = float(new_value)
            data_store["manual_override"] = True
            print(f"[✅] Setpoint Manually Updated to {new_value}°C - Manual Override Enabled FOREVER")
            return {"message": "Set Temperature updated", "SetTemperature": new_value}
        return {"error": "Invalid input"}, 400

api.add_resource(SetpointAPI, "/setpoint")

def run_flask_app(port=5000):
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)

# ------------------ Charting ------------------
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator

class TemperatureControlApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Transfer GUI")
        self.master.geometry("900x500")

        self.client = None
        self.running = True
        self.api_thread = None
        self.modbus_thread = None

        main_frame = tk.Frame(master)
        main_frame.pack(fill="both", expand=True)

        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill="y", padx=10, pady=10)

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=10, pady=10)

        # 1) Temperature Display
        temp_frame = tk.Frame(left_frame)
        temp_frame.pack(pady=10)

        self.temp_label_number = tk.Label(temp_frame, text="--", font=("Arial", 80, "bold"))
        self.temp_label_number.pack(side=tk.LEFT)

        # smaller font for the "°C" unit
        self.temp_label_unit = tk.Label(temp_frame, text="°C", font=("Arial", 30))
        self.temp_label_unit.pack(side=tk.LEFT, padx=(10,0))

        # 2) Power
        power_frame = tk.Frame(left_frame)
        power_frame.pack(pady=10)

        tk.Label(power_frame, text="Power:", font=("Arial", 18)).pack(side=tk.LEFT)
        self.power_bar = ttk.Progressbar(power_frame, orient="horizontal", length=100, mode="determinate")
        self.power_bar.pack(side=tk.LEFT, padx=5)
        self.power_bar["maximum"] = 100

        self.power_percent_label = tk.Label(power_frame, text="--%", font=("Arial", 16))
        self.power_percent_label.pack(side=tk.LEFT, padx=5)

        # 3) SetPoint
        sp_frame = tk.Frame(left_frame)
        sp_frame.pack(pady=10)

        tk.Label(sp_frame, text="SetPoint:", font=("Arial", 18)).pack(side=tk.LEFT, padx=5)
        self.set_point_var = tk.StringVar()
        self.set_point_entry = tk.Entry(sp_frame, textvariable=self.set_point_var,
                                        font=("Arial", 16), width=8, justify="center")
        self.set_point_entry.pack(side=tk.LEFT)
        self.set_point_entry.bind("<Return>", self.on_enter_set_point)
        tk.Label(sp_frame, text="°C", font=("Arial", 18)).pack(side=tk.LEFT, padx=5)

        # 4) Config Frame
        config_frame = tk.Frame(left_frame, bg="#F0F0F0", bd=2, relief=tk.GROOVE)
        config_frame.pack(pady=10, fill="x")

        tk.Label(config_frame, text="COM Port:", font=("Arial", 12), bg="#F0F0F0").grid(
            row=0, column=0, padx=5, pady=5, sticky="e"
        )
        self.com_var = tk.StringVar(value="COM4")
        self.com_entry = tk.Entry(config_frame, textvariable=self.com_var, font=("Arial", 12), width=10)
        self.com_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Button(config_frame, text="Connect", font=("Arial", 12),
                  command=self.on_init_controller).grid(row=0, column=2, padx=10, pady=5, sticky="w")

        self.api_on_var = tk.BooleanVar(value=True)
        chk_api = tk.Checkbutton(config_frame, text="API On", variable=self.api_on_var,
                                 font=("Arial", 12), bg="#F0F0F0")
        chk_api.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        tk.Label(config_frame, text="Port:", font=("Arial", 12), bg="#F0F0F0").grid(
            row=1, column=1, padx=5, pady=5, sticky="e"
        )
        self.port_var = tk.StringVar(value="5000")
        self.port_entry = tk.Entry(config_frame, textvariable=self.port_var,
                                   font=("Arial", 12), width=6)
        self.port_entry.grid(row=1, column=2, padx=5, pady=5, sticky="w")

        self.data_save_var = tk.BooleanVar(value=True)
        chk_save = tk.Checkbutton(config_frame, text="Save Data", variable=self.data_save_var,
                                  font=("Arial", 12), bg="#F0F0F0")
        chk_save.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        tk.Label(config_frame, text="Log Dir:", font=("Arial", 12), bg="#F0F0F0").grid(
            row=2, column=1, padx=5, pady=5, sticky="e"
        )

        self.log_dir = tk.StringVar(value=r"C:\Users\WangLabAdmin\Desktop\TransferLog")
        self.log_dir_entry = tk.Entry(config_frame, textvariable=self.log_dir,
                                      font=("Arial",12), width=20)
        self.log_dir_entry.grid(row=2, column=2, padx=5, pady=5, sticky="w")

        tk.Button(config_frame, text="Browse", command=self.browse_log_dir).grid(
            row=2, column=3, padx=5, pady=5, sticky="w"
        )

                # --- ADD HERE: keyboard ↑ / ↓ bump bindings --------------------
        self.master.bind("<KeyRelease-Up>",   self.on_key_up)
        self.master.bind("<KeyRelease-Down>", self.on_key_down)
        self.master.focus_set()
        # ----------------------------------------------------------------


        # 5) Chart
        self.fig, self.ax = plt.subplots(figsize=(4,3))
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Temperature (°C)")
        (self.line,) = self.ax.plot([], [], marker="o", markersize=2, linestyle="-")
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # We keep only the last 30 min, so memory won't grow unbounded
        self.times = []
        self.temps = []

    def on_init_controller(self):
        global client, stop_threads
        port_name = self.com_var.get().strip()
        new_client = ModbusClient(
            port=port_name,
            baudrate=9600,
            parity='N',
            stopbits=1,
            bytesize=8,
            timeout=3
        )
        if new_client.connect():
            client = new_client
            stop_threads = False

            log_path = self.log_dir.get()
            if not os.path.isdir(log_path):
                os.makedirs(log_path)

            self.modbus_thread = threading.Thread(
                target=update_modbus_values_loop,
                args=(self,),
                daemon=True
            )
            self.modbus_thread.start()

            if self.api_on_var.get():
                try:
                    port_num = int(self.port_var.get().strip())
                except ValueError:
                    port_num = 5000
                self.api_thread = threading.Thread(
                    target=run_flask_app,
                    args=(port_num,),
                    daemon=True
                )
                self.api_thread.start()

            self.ui_thread = threading.Thread(
                target=self.ui_update_loop,
                daemon=True
            )
            self.ui_thread.start()
        else:
            print(f"[Init] Failed to connect on {port_name}")

    def on_enter_set_point(self, event):
        if not client:
            print("[SetPoint] Not connected!")
            return

        val_str = self.set_point_var.get().strip()
        if not val_str:
            return
        try:
            val_f = float(val_str)
        except ValueError:
            print("[SetPoint] Invalid float:", val_str)
            return

        write_register(client, 3000, val_f)
        data_store["SetTemperature"] = val_f
        data_store["manual_override"] = True
        print(f"[UI] User set new setpoint => {val_f:.1f} °C (3000)")

    def browse_log_dir(self):
        folder = filedialog.askdirectory()
        if folder:
            self.log_dir.set(folder)



        # ---------- ADD HERE: keyboard-control helpers ----------
    def _bump_set_point(self, delta):
        """Internal: nudge set-temperature up/down by Δ°C."""
        cur = data_store.get("SetTemperature")
        if cur is None:
            return
        new_val = cur + delta

        if client:
            write_register(client, 3000, new_val)

        data_store["SetTemperature"] = new_val
        data_store["manual_override"] = True

        # Update entry field unless the user is typing
        if self.set_point_entry.focus_get() != self.set_point_entry:
            self.set_point_var.set(f"{new_val:.1f}")

        print(f"[KB] Set-point changed to {new_val:.1f} °C")

    def on_key_up(self, event):
        self._bump_set_point(+1)

    def on_key_down(self, event):
        self._bump_set_point(-1)
    # ---------------------------------------------------------




    def ui_update_loop(self):
        while True:
            if stop_threads or not self.running:
                break

            # 1) Temperature
            temp = data_store.get("Temperature", None)
            # 2) Power
            power = data_store.get("Power", None)
            # 3) SetTemperature
            spt  = data_store.get("SetTemperature", None)

            # Update the display
            if temp is not None:
                self.temp_label_number.config(text=f"{temp:.1f}")
            else:
                self.temp_label_number.config(text="--")

            if power is not None:
                p_val = max(0, min(100, power))
                self.power_bar["value"] = p_val
                self.power_percent_label.config(text=f"{power:.1f}%")
            else:
                self.power_bar["value"] = 0
                self.power_percent_label.config(text="--%")

            if spt is not None:
                if self.set_point_entry.focus_get() != self.set_point_entry:
                    self.set_point_var.set(f"{spt:.1f}")
            else:
                self.set_point_var.set("")

            # Plot temperature if we have it
            if temp is not None:
                self.add_temp_to_plot(temp)

            time.sleep(0.5)

    def add_temp_to_plot(self, temp):
        # We'll store only the last 30 minutes of data
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
        """Close event => cleanup threads, then fully exit Python process."""
        global stop_threads
        stop_threads = True
        self.running = False
        if client:
            client.close()
        self.master.destroy()

        ### CHANGED: Force entire program to stop
        import sys
        sys.exit(0)  # or os._exit(0)

### CHANGED: No other modifications

if __name__ == "__main__":
    root = tk.Tk()
    app = TemperatureControlApp(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()
