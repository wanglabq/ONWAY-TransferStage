import clr
import System
import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel
import time
import datetime
import threading
import os
import csv

from flask import Flask, jsonify

# ----------------------------------------------------------------------------
#  1) Create the main root window FIRST (important for tk.StringVar usage!)
# ----------------------------------------------------------------------------
root = tk.Tk()
root.title("R & Z Axis Control Panel")
root.geometry("1055x470")
root.configure(bg="#F0F0F0")

# ----------------------------------------------------------------------------
#  2) Flask-based read-only API
# ----------------------------------------------------------------------------
app = Flask(__name__)

api_data = {
    "r_position": 0.0,
    "z_position": 0.0,
    "r_velocity": 0.0,
    "z_velocity": 0.0,
    "r_acceleration": 0.0,
    "z_acceleration": 0.0,
}


@app.route("/api/status", methods=["GET"])
def get_status():
    return jsonify(api_data)

def run_api_server():
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)

# ----------------------------------------------------------------------------
#  3) DLL, Motion Controller, and Global Data
# ----------------------------------------------------------------------------
dll_path = r"C:\Users\WangLabAdmin\Desktop\DTS\MCC4DLL.dll"
clr.AddReference(dll_path)

from SerialPortLibrary import SPLibClass
serial_port = SPLibClass()

# Storing references to Tkinter Entry widgets
entry_widgets = {
    0: {"velocity": None, "acceleration": None},
    1: {"velocity": None, "acceleration": None},
}

position_display = {0: None, 1: None}
abs_input_widgets = {0: None, 1: None}
rel_input_widgets = {0: None, 1: None}

# For deciding whether to save logs
save_log_var = tk.BooleanVar(master=root, value=True)
# Default log path
log_save_path = tk.StringVar(master=root, value=r"C:\Users\WangLabAdmin\Desktop\DTS")
# Use API?
api_enabled_var = tk.BooleanVar(master=root, value=True)

written_headers = set()

# Track editing states so auto-refresh won't overwrite user typing
editing_flags = {
    (0, "velocity"): False,
    (0, "acceleration"): False,
    (1, "velocity"): False,
    (1, "acceleration"): False,
}

# ----------------------------------------------------------------------------
#  Default velocity/acceleration for each axis
# ----------------------------------------------------------------------------

# Track whether keyboard shortcuts are allowed
keyboard_enabled_var = tk.BooleanVar(master=root, value=True)


def on_focus_in_param(event, axis_id, param_name):
    editing_flags[(axis_id, param_name)] = True

def on_focus_out_param(event, axis_id, param_name):
    editing_flags[(axis_id, param_name)] = False

# ----------------------------------------------------------------------------
#  4) Functions
# ----------------------------------------------------------------------------
def log_message(info: str):
    now_str = time.strftime("%Y-%m-%d %H:%M:%S")
    full_line = f"{now_str}\t{info}"

    # A) Print in GUI
    log_textbox.config(state=tk.NORMAL)
    log_textbox.insert(tk.END, full_line + "\n")
    log_textbox.see(tk.END)
    log_textbox.config(state=tk.DISABLED)

    # B) Possibly save to CSV
    if save_log_var.get():
        folder = log_save_path.get().strip()
        if folder:
            append_csv_log(folder, now_str, info)

def append_csv_log(folder, timestamp_str, message):
    month_str = time.strftime("%Y-%m")
    filename = os.path.join(folder, f"log_{month_str}.csv")
    file_exists = os.path.isfile(filename)

    with open(filename, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists or filename not in written_headers:
            writer.writerow(["Timestamp", "Message"])
            written_headers.add(filename)
        writer.writerow([timestamp_str, message])

def pause_all_axes():
    ret = serial_port.MoCtrCard_PauseAxisMov(System.Byte(255))
    if ret != serial_port.FUNRES_OK:
        log_message("PauseAxisMov(255) returned an error.")

def resume_all_axes():
    ret = serial_port.MoCtrCard_ReStartAxisMov(System.Byte(255))
    if ret != serial_port.FUNRES_OK:
        log_message("ReStartAxisMov(255) returned an error.")

def stop_all_axes():
    """Stop both X and Y axes immediately."""
    ret_x = serial_port.MoCtrCard_StopAxisMov(System.Byte(0))
    ret_y = serial_port.MoCtrCard_StopAxisMov(System.Byte(1))
    if ret_x == serial_port.FUNRES_OK and ret_y == serial_port.FUNRES_OK:
        log_message("[STOP] Stopped both R and Z axes.")
        for ax in (0,1):
            entry_widgets[ax]["velocity"].config(state="normal")
            entry_widgets[ax]["acceleration"].config(state="normal")
    else:
        log_message("[STOP] Error stopping R and/or Z axis.")

def init_controller(com_port: str):
    """
    Initialize the controller, then set each axis velocity/accel
    to the current default values (both in the device & in the UI).
    """
    status_init = serial_port.MoCtrCard_Initial(com_port)
    if status_init != serial_port.FUNRES_OK:
        log_message(f"❌ Initialization failed on {com_port}!")
    else:
        log_message(f"✅ Motion Controller Initialized on {com_port}.")
    
    resume_all_axes()


# ----------------------------------------------------------------------------
#  5) Reading & Updating Axis Values
# ----------------------------------------------------------------------------
def read_axis_params(axis_id):
    ResPos = System.Array.CreateInstance(System.Single, 1)
    status_pos = serial_port.MoCtrCard_GetAxisPos(System.Byte(axis_id), ResPos)
    pos_val = ResPos[0] if status_pos == serial_port.FUNRES_OK else None

    ResSpd = System.Array.CreateInstance(System.Single, 1)
    status_spd = serial_port.MoCtrCard_ReadPara(System.Byte(axis_id), System.Byte(2), ResSpd)
    vel_val = ResSpd[0] if status_spd == serial_port.FUNRES_OK else None

    ResAcc = System.Array.CreateInstance(System.Single, 1)
    status_acc = serial_port.MoCtrCard_ReadPara(System.Byte(axis_id), System.Byte(3), ResAcc)
    acc_val = ResAcc[0] if status_acc == serial_port.FUNRES_OK else None

    # If you still want to read step size behind the scenes, you can keep it:
    ResStp = System.Array.CreateInstance(System.Single, 1)
    status_stp = serial_port.MoCtrCard_ReadPara(System.Byte(axis_id), System.Byte(0), ResStp)
    step_val = ResStp[0] if status_stp == serial_port.FUNRES_OK else None

    return (pos_val, vel_val, acc_val, step_val)

def update_axis_ui(axis_id):
    result = read_axis_params(axis_id)
    if not result:
        return
    pos_val, vel_val, acc_val, step_val = result
    if pos_val is None or vel_val is None or acc_val is None:
        return

    # -- Position (read-only)
    if position_display[axis_id] is not None:
        position_display[axis_id].config(state=tk.NORMAL)
        position_display[axis_id].delete(0, tk.END)
        # For X-axis => °, Y-axis => mm
        position_display[axis_id].insert(0, f"{pos_val:.3f}")
        position_display[axis_id].config(state="readonly")

    # -- Velocity
    if not editing_flags[(axis_id, "velocity")]:
        ent_vel = entry_widgets[axis_id]["velocity"]
        ent_vel.delete(0, tk.END)
        ent_vel.insert(0, f"{vel_val:.3f}")

    # -- Acceleration
    if not editing_flags[(axis_id, "acceleration")]:
        ent_acc = entry_widgets[axis_id]["acceleration"]
        ent_acc.delete(0, tk.END)
        ent_acc.insert(0, f"{acc_val:.3f}")

    # -- Update the read-only API
    if axis_id == 0:
        api_data["r_position"] = float(f"{pos_val:.3f}")
        api_data["r_velocity"] = float(f"{vel_val:.3f}")
        api_data["r_acceleration"] = float(f"{acc_val:.3f}")
    else:
        api_data["z_position"] = float(f"{pos_val:.3f}")
        api_data["z_velocity"] = float(f"{vel_val:.3f}")
        api_data["z_acceleration"] = float(f"{acc_val:.3f}")

def auto_refresh_loop():
    update_axis_ui(0)
    update_axis_ui(1)
    root.after(100, auto_refresh_loop)

# ----------------------------------------------------------------------------
#  6) Move / Param Commands
# ----------------------------------------------------------------------------


def move_abs(axis_id):
    label = "R" if axis_id==0 else "Z"
    unit = "°" if axis_id==0 else "mm"
    resume_all_axes()
    text_val = abs_input_widgets[axis_id].get().strip()
    try:
        val = float(text_val)
    except ValueError:
        log_message(f"[ABS] Invalid input for axis {label}: '{text_val}' {unit}")
        return
    if axis_id == 1 and val < 0.0:
        log_message("[ABS] Z abs target < 0, clamping to 0")
        val = 0.0


    status = serial_port.MoCtrCard_MCrlAxisAbsMove(System.Byte(axis_id), System.Single(val))
    if status == serial_port.FUNRES_OK:
        log_message(f"[ABS] Axis {label} => {val:.3f} {unit}")

    else:
        log_message(f"[ABS] Axis {label} => {val:.3f} {unit} FAILED")
        


def move_rel(axis_id, direction):
    label = "R" if axis_id==0 else "Z"
    unit = "°" if axis_id==0 else "mm"
    resume_all_axes()
    text_val = rel_input_widgets[axis_id].get().strip()
    try:
        val = float(text_val)
    except ValueError:
        log_message(f"[REL] Invalid input for axis {label}: '{text_val}' {unit}")
        return

    rel_val = val * direction
    # compute destination from current pos
    pos_current, _, _, _ = read_axis_params(axis_id)

    if axis_id == 1:
        pos_current, _, _, _ = read_axis_params(1)
        if pos_current is not None and pos_current + rel_val < 0.0:
            log_message("[REL] Z rel would go <0, clamping to 0")
            rel_val = -pos_current
    
    status = serial_port.MoCtrCard_MCrlAxisRelMove(System.Byte(axis_id), System.Single(rel_val))
    if status == serial_port.FUNRES_OK:
        log_message(f"[REL] Axis {label} => move by {rel_val:.3f} {unit}")

    else:
        log_message(f"[REL] Axis {label} => move by {rel_val:.3f} {unit} FAILED")
 

def on_velocity_enter(event, axis_id):
    label = "R" if axis_id==0 else "Z"
    unit = "°/s" if axis_id==0 else "mm/s"
    editing_flags[(axis_id, "velocity")] = False

    text_val = entry_widgets[axis_id]["velocity"].get().strip()
    try:
        val = float(text_val)
    except ValueError:
        log_message(f"[VEL] Invalid velocity for axis {label}: '{text_val}' {unit}")
        return

    stop_all_axes()   # <-- NEW – guarantee motor is idle
    ret = serial_port.MoCtrCard_SendPara(System.Byte(axis_id), System.Byte(2), System.Single(val))
    if ret == serial_port.FUNRES_OK:
        log_message(f"[VEL] Axis {label} => {val:.3f} {unit}")
    else:
        log_message(f"[VEL] Axis {label} => {val:.3f} {unit} FAILED")

def on_acceleration_enter(event, axis_id):
    label = "R" if axis_id==0 else "Z"
    unit = "°/s²" if axis_id==0 else "mm/s²"
    editing_flags[(axis_id, "acceleration")] = False

    text_val = entry_widgets[axis_id]["acceleration"].get().strip()
    try:
        val = float(text_val)
    except ValueError:
        log_message(f"[ACC] Invalid accel for axis {label}: '{text_val}' {unit}")
        return

    stop_all_axes()   # <-- NEW – guarantee motor is idle
    ret = serial_port.MoCtrCard_SendPara(System.Byte(axis_id), System.Byte(3), System.Single(val))
    if ret == serial_port.FUNRES_OK:
        log_message(f"[ACC] Axis {label} => {val:.3f} {unit}")
    else:
        log_message(f"[ACC] Axis {label} => {val:.3f} {unit} FAILED")

def set_default_params():
    """Set BOTH axes velocity = 2 and acceleration = 2."""
    stop_all_axes()   # <-- NEW – guarantee motor is idle
    # velocity
    ent_v = entry_widgets[0]["velocity"]
    ent_v.delete(0, tk.END)
    ent_v.insert(0, "0.5")
    on_velocity_enter(None, 0)          # writes to controller

    # velocity
    ent_v = entry_widgets[1]["velocity"]
    ent_v.delete(0, tk.END)
    ent_v.insert(0, "0.1")
    on_velocity_enter(None, 1)          # writes to controller

    # acceleration
    ent_a = entry_widgets[1]["acceleration"]
    ent_a.delete(0, tk.END)
    ent_a.insert(0, "0.5")
    on_acceleration_enter(None, 1)      # writes to controller

    # acceleration
    ent_a = entry_widgets[0]["acceleration"]
    ent_a.delete(0, tk.END)
    ent_a.insert(0, "0.5")
    on_acceleration_enter(None, 0)      # writes to controller

def stop_axis(axis_id):
    resume_all_axes()
    ret = serial_port.MoCtrCard_StopAxisMov(System.Byte(axis_id))
    if ret == serial_port.FUNRES_OK:
        log_message(f"[STOP] Stopped {['R','Z'][axis_id]} axis.")
        for ax in (0,1):
            entry_widgets[ax]["velocity"].config(state="normal")
            entry_widgets[ax]["acceleration"].config(state="normal")
    else:
        log_message(f"[STOP] Error stopping {['R','Z'][axis_id]} axis.")


# ----------------------------------------------------------------------------
#  7) Home only (no reset coords)
# ----------------------------------------------------------------------------
def home_axis(axis_id):
    """Send axis to absolute position 0.0."""
    resume_all_axes()
    ret = serial_port.MoCtrCard_MCrlAxisAbsMove(System.Byte(axis_id), System.Single(0.0))
    
    label = "R" if axis_id==0 else "Z"
    if ret == serial_port.FUNRES_OK:
        log_message(f"[HOME] Axis {label} => 0.0")
    else:
        log_message(f"[HOME] Axis {label} => 0.0 FAILED")

# ----------------------------------------------------------------------------
#  8) Keyboard Controls
# ----------------------------------------------------------------------------

# at module scope
# at module scope
repeat_job   = None
repeat_delta = None

def repeat_move():
    """Called every 2 s while the key is held."""
    global repeat_job
    y_rel_move(repeat_delta)
    repeat_job = root.after(500, repeat_move)

def on_key_press(event):
    if not keyboard_enabled_var.get():
        return None          #  <-- early-exit when checkbox is off
    global repeat_job, repeat_delta

    # ── ignore all OS auto-repeats once we've started ──
    if repeat_job is not None:
        return "break"

    key   = event.keysym.lower()
    state = event.state

    SHIFT = bool(state & 0x0001)
    CTRL  = bool(state & 0x0004)
    ALT   = bool((state & 0x0008) != 0 or (state & 0x20000) != 0)

    # ── compute delta just like before ──
    delta = None
    if key == "num_lock":
        if SHIFT and CTRL and ALT:
            delta = -0.001
        elif SHIFT and ALT:
            delta = -0.005
        elif CTRL  and ALT:
            delta = +0.001
        elif ALT:
            delta = +0.005

    elif key == "period" and CTRL:
        delta = +0.025
    elif key == "comma" and CTRL:
        delta = -0.025
    elif key == "space" and (CTRL and ALT and SHIFT):
        stop_all_axes()
        return "break"

    if delta is not None:
        # do one move now
        y_rel_move(delta)
        repeat_delta = delta
        # then start the 2 s loop
        repeat_job = root.after(500, repeat_move)

    return "break"

def on_key_release(event):
    if not keyboard_enabled_var.get():
        return None          #  <-- early-exit when checkbox is off
    global repeat_job, repeat_delta
    # cancel the 2 s loop as soon as any key is lifted
    if repeat_job is not None:
        root.after_cancel(repeat_job)
        repeat_job = None
        repeat_delta = None
    return None


def y_rel_move(delta):
    pos_current, _, _, _ = read_axis_params(1)

    axis_id = 1
    pos_current, _, _, _ = read_axis_params(1)
    if pos_current is not None and pos_current + delta < 0.0:
        log_message("[REL] Z key‐move would go <0, clamping to 0")
        delta = -pos_current
    resume_all_axes()
    status = serial_port.MoCtrCard_MCrlAxisRelMove(System.Byte(axis_id), System.Single(delta))
    if status == serial_port.FUNRES_OK:
        direction_str = "UP" if delta > 0 else "DOWN"
        log_message(f"[REL] Z => {direction_str} {abs(delta)} mm")

    else:
        log_message(f"[REL] Z => move {delta} mm FAILED")


# ----------------------------------------------------------------------------
#  9) Build the UI
# ----------------------------------------------------------------------------

#
#  Left side: X Axis frame, Y Axis frame
#
main_frame = tk.Frame(root, bg="#F0F0F0")
main_frame.pack(side=tk.LEFT, fill="both", padx=10, pady=10, expand=True)

def create_axis_frame(parent, axis_id, axis_label):
    frm = tk.LabelFrame(parent, text=f"{axis_label} Axis", font=("Arial", 14, "bold"), bg="#F0F0F0", bd=3)
    frm.pack(pady=10, fill="x")

    # Row: Position
    row_pos = tk.Frame(frm, bg="#F0F0F0")
    row_pos.pack(pady=5, fill="x")
    tk.Label(row_pos, text="Position:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)

    pos_display_ent = tk.Entry(row_pos, font=("Arial", 12), width=7, state="readonly")
    pos_display_ent.pack(side=tk.LEFT, padx=5)
    position_display[axis_id] = pos_display_ent

    # Show a unit label after the position display
    if axis_id == 0:
        tk.Label(row_pos, text="°", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    else:
        tk.Label(row_pos, text="mm", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)

    # Row: Absolute move
    tk.Label(row_pos, text="Abs:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=10)
    abs_var = tk.StringVar()
    abs_ent = tk.Entry(row_pos, font=("Arial", 12), width=7, textvariable=abs_var)
    abs_ent.pack(side=tk.LEFT, padx=2)
    unit = "°" if axis_id == 0 else "mm"
    tk.Label(row_pos, text=unit, font=("Arial",12), bg="#F0F0F0") \
    .pack(side=tk.LEFT, padx=(0,5))
    abs_input_widgets[axis_id] = abs_ent

    go_btn = tk.Button(row_pos, text="Go", font=("Arial", 10), command=lambda ax=axis_id: move_abs(ax))
    go_btn.pack(side=tk.LEFT, padx=5)

    # Row: Relative move
    tk.Label(row_pos, text="Rel:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    rel_var = tk.StringVar()
    rel_ent = tk.Entry(row_pos, font=("Arial", 12), width=7, textvariable=rel_var)
    rel_ent.pack(side=tk.LEFT, padx=2)
    unit = "°" if axis_id == 0 else "mm"
    tk.Label(row_pos, text=unit, font=("Arial",12), bg="#F0F0F0") \
    .pack(side=tk.LEFT, padx=(0,5))
    rel_input_widgets[axis_id] = rel_ent

    plus_btn = tk.Button(row_pos, text="+", font=("Arial", 10),
                         command=lambda ax=axis_id: move_rel(ax, +1))
    plus_btn.pack(side=tk.LEFT, padx=2)
    minus_btn = tk.Button(row_pos, text="-", font=("Arial", 10),
                          command=lambda ax=axis_id: move_rel(ax, -1))
    minus_btn.pack(side=tk.LEFT, padx=2)

    # Row: Velocity
    row_vel = tk.Frame(frm, bg="#F0F0F0")
    row_vel.pack(pady=5, fill="x")
    tk.Label(row_vel, text="Velocity:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    ent_vel = tk.Entry(row_vel, font=("Arial", 12), width=7)
    ent_vel.pack(side=tk.LEFT, padx=5)
    ent_vel.bind("<FocusIn>", lambda e, ax=axis_id: on_focus_in_param(e, ax, "velocity"))
    ent_vel.bind("<FocusOut>", lambda e, ax=axis_id: on_focus_out_param(e, ax, "velocity"))
    ent_vel.bind("<Return>", lambda e, ax=axis_id: on_velocity_enter(e, ax))
    entry_widgets[axis_id]["velocity"] = ent_vel

    if axis_id == 0:
        tk.Label(row_vel, text="°/s", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    else:
        tk.Label(row_vel, text="mm/s", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)

    # Row: Acceleration
    row_acc = tk.Frame(frm, bg="#F0F0F0")
    row_acc.pack(pady=5, fill="x")
    tk.Label(row_acc, text="Acceleration:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    ent_acc = tk.Entry(row_acc, font=("Arial", 12), width=7)
    ent_acc.pack(side=tk.LEFT, padx=5)
    ent_acc.bind("<FocusIn>", lambda e, ax=axis_id: on_focus_in_param(e, ax, "acceleration"))
    ent_acc.bind("<FocusOut>", lambda e, ax=axis_id: on_focus_out_param(e, ax, "acceleration"))
    ent_acc.bind("<Return>", lambda e, ax=axis_id: on_acceleration_enter(e, ax))
    entry_widgets[axis_id]["acceleration"] = ent_acc

    if axis_id == 0:
        tk.Label(row_acc, text="°/s²", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)
    else:
        tk.Label(row_acc, text="mm/s²", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=5)

    # Row: Home button
        # Row: Home + per-axis Stop
    stop_btn = tk.Button(row_acc,
                         text=f"Stop {axis_label}",
                         font=("Arial",12),
                         command=lambda ax=axis_id: stop_axis(ax))
    stop_btn.pack(side=tk.RIGHT, padx=5)

    home_btn = tk.Button(row_acc,
                         text=f"Home {axis_label}",
                         font=("Arial",12),
                         command=lambda ax=axis_id: home_axis(ax))
    home_btn.pack(side=tk.RIGHT, padx=5)


    return frm

frame_x = create_axis_frame(main_frame, 0, "R")
frame_y = create_axis_frame(main_frame, 1, "Z")

#
#  Settings Frame (below X & Y)
#
settings_frame = tk.LabelFrame(
    main_frame,
    text="Settings",
    font=("Arial", 14, "bold"),
    bg="#F0F0F0",
    bd=3
)
settings_frame.pack(fill="x")

top_line = tk.Frame(settings_frame, bg="#F0F0F0")
top_line.pack(fill="x", pady=5)

# Left chunk: Use API
chk_api = tk.Checkbutton(top_line, text="Use API", font=("Arial", 12),
                         variable=api_enabled_var, bg="#F0F0F0")
chk_api.pack(side=tk.LEFT, padx=(10, 10))

# COM port chunk
com_label = tk.Label(top_line, text="COM Port:", font=("Arial", 12), bg="#F0F0F0")
com_label.pack(side=tk.LEFT, padx=5)
com_var = tk.StringVar(value="COM5")
com_entry = tk.Entry(top_line, textvariable=com_var, font=("Arial", 12), width=7)
com_entry.pack(side=tk.LEFT, padx=5)

def on_init_controller():
    init_controller(com_var.get())

tk.Button(top_line, text="Connect", font=("Arial", 12), command=on_init_controller).pack(side=tk.LEFT, padx=10)
tk.Button(top_line,
          text="Default",
          font=("Arial", 12),
          command=set_default_params).pack(side=tk.LEFT, padx=10)

# (Removed the "Default" button entirely)

# Right chunk: Save Log
chk_savelog = tk.Checkbutton(top_line, text="Save Log", font=("Arial", 12),
                             variable=save_log_var, bg="#F0F0F0")
chk_savelog.pack(side=tk.RIGHT, padx=20)

chk_keyboard = tk.Checkbutton(
    top_line,
    text="Keyboard Ctrl",
    font=("Arial", 12),
    variable=keyboard_enabled_var,
    bg="#F0F0F0"
)
chk_keyboard.pack(side=tk.RIGHT, padx=10)

chk_savelog.pack_forget()           # remove, then re-pack so order is nice
chk_keyboard.pack(side=tk.RIGHT, padx=10)
chk_savelog.pack(side=tk.RIGHT, padx=10)

# Second line: Log path
bottom_line = tk.Frame(settings_frame, bg="#F0F0F0")
bottom_line.pack(fill="x", pady=5)

tk.Label(bottom_line, text="Log Save Path:", font=("Arial", 12), bg="#F0F0F0").pack(side=tk.LEFT, padx=(10,5))
log_path_entry = tk.Entry(bottom_line, textvariable=log_save_path, font=("Arial", 12), width=38)
log_path_entry.pack(side=tk.LEFT, padx=5)

def browse_log_folder():
    folder = filedialog.askdirectory()
    if folder:
        log_save_path.set(folder)

tk.Button(bottom_line, text="Browse", font=("Arial", 12), command=browse_log_folder).pack(side=tk.LEFT, padx=10)

#
# Right-side log frame
#
right_frame = tk.Frame(root, bg="#F0F0F0", bd=2, relief=tk.SUNKEN)
right_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=5, pady=5)

log_textbox = tk.Text(right_frame, width=50, state=tk.DISABLED, font=("Arial", 11))
log_textbox.pack(fill="both", expand=True)

# ----------------------------------------------------------------------------
# 10) Bind Keys, Start Loops, Possibly Launch API
# ----------------------------------------------------------------------------
root.bind("<KeyPress>",   on_key_press)
root.bind("<KeyRelease>", on_key_release)
root.after(100, auto_refresh_loop)

if api_enabled_var.get():
    api_thread = threading.Thread(target=run_api_server, daemon=True)
    api_thread.start()

def on_closing():
    # unload the controller first
    try:
        serial_port.MoCtrCard_Unload()
    except Exception:
        pass

    # stop the Tk event loop – this makes root.mainloop() return
    root.quit()                 # <─ change: quit() instead of destroy()
    # DO NOT call sys.exit() or destroy() here – let the code run on
    # after mainloop() has really finished.

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
root.destroy()                  # now it’s safe to destroy all widgets
import os, sys
os._exit(0)                     # guarantees the process dies, even if
                                # a non‑daemon thread (e.g. Flask) is
                                # still hanging around