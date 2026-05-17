"""
6-Axis Motion Controller  —  Redesigned UI
Dark industrial theme · reduced polling overhead · cleaner layout
"""

import os, csv, time, threading, queue
import clr, System
import tkinter as tk
from tkinter import filedialog, messagebox
import pygame

# ─────────────────────────────────────────────────────────────────────────────
#  Constants & axis map
# ─────────────────────────────────────────────────────────────────────────────
DLL_PATH         = r"C:\Users\WangLabAdmin\Desktop\MotionControl\MCC6DLL.dll"
DEFAULT_COM_PORT = "COM4"
NUM_AXES         = 6
JOYSTICK_SPEED   = 0.5

AXES = [
    dict(label="X", unit="mm", vel_unit="mm/s",  acc_unit="mm/s²"),
    dict(label="Y", unit="mm", vel_unit="mm/s",  acc_unit="mm/s²"),
    dict(label="Z", unit="°",  vel_unit="°/s",   acc_unit="°/s²"),
    dict(label="A", unit="mm", vel_unit="mm/s",  acc_unit="mm/s²"),
    dict(label="B", unit="mm", vel_unit="mm/s",  acc_unit="mm/s²"),
    dict(label="C", unit="mm", vel_unit="mm/s",  acc_unit="mm/s²"),
]

DEFAULT_VEL  = 1.0
DEFAULT_ACC  = 0.2
PARA_VELOCITY = 2
PARA_ACCEL    = 3

# ─────────────────────────────────────────────────────────────────────────────
#  Theme palette — clean white / minimal
# ─────────────────────────────────────────────────────────────────────────────
T = dict(
    bg          = "#F5F5F5",   # light page background
    panel       = "#FFFFFF",   # card / panel white
    border      = "#E0E0E0",   # subtle grey border
    accent      = "#1A73E8",   # blue — action / connect
    accent2     = "#D32F2F",   # red — STOP
    accent3     = "#2E7D32",   # green — Go / Home
    text        = "#212121",   # primary text
    text_dim    = "#757575",   # secondary / label text
    entry_bg    = "#FAFAFA",   # input background
    entry_act   = "#E8F0FE",   # focused input tint
    label_ax    = "#1A73E8",   # axis letter (blue)
)

FONT_MONO    = ("Segoe UI", 11)
FONT_MONO_SM = ("Segoe UI", 10)
FONT_BOLD    = ("Segoe UI", 12, "bold")
FONT_HEAD    = ("Segoe UI", 13, "bold")
FONT_SMALL   = ("Segoe UI", 9)

# ─────────────────────────────────────────────────────────────────────────────
#  DLL init
# ─────────────────────────────────────────────────────────────────────────────
try:
    clr.AddReferenceToFileAndPath(DLL_PATH)
except Exception:
    clr.AddReference(DLL_PATH)

from SerialPortLibrary import SPLibClass
serial_port = SPLibClass()
if not hasattr(serial_port, "FUNRES_OK"):
    serial_port.FUNRES_OK  = serial_port.FunResOk
    serial_port.FUNRES_ERR = serial_port.FunResErr

serial_lock = threading.RLock()
from contextlib import contextmanager

@contextmanager
def motor_lock():
    with serial_lock:
        yield

ALL_AXES = System.Byte(255)

def resume_all_axes():
    with motor_lock():
        serial_port.MoCtrCard_ReStartAxisMov(ALL_AXES)

# ─────────────────────────────────────────────────────────────────────────────
#  Root window
# ─────────────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("6-Axis Motion Controller")
root.geometry("1060x680")
root.configure(bg=T["bg"])
try:
    root.tk_setPalette(background=T["bg"])
except Exception:
    pass
root.resizable(True, True)

save_log_var  = tk.BooleanVar(value=True, master=root)
log_save_path = tk.StringVar(
    value=r"C:\Users\WangLabAdmin\Desktop\6motion", master=root)

# ─────────────────────────────────────────────────────────────────────────────
#  Per-axis state
# ─────────────────────────────────────────────────────────────────────────────
entry_widgets     = {ax: {"velocity": None, "acceleration": None}
                     for ax in range(NUM_AXES)}
position_display  = {ax: None for ax in range(NUM_AXES)}
abs_input_widgets = {ax: None for ax in range(NUM_AXES)}
rel_input_widgets = {ax: None for ax in range(NUM_AXES)}
editing_flags     = {(ax, p): False
                     for ax in range(NUM_AXES)
                     for p in ("velocity", "acceleration")}

# ─────────────────────────────────────────────────────────────────────────────
#  File logging
# ─────────────────────────────────────────────────────────────────────────────
file_log_queue: queue.Queue = queue.Queue()

def _log_writer():
    while True:
        ts, msg = file_log_queue.get()
        try:
            fn = os.path.join(
                log_save_path.get(), f"log_{time.strftime('%Y-%m')}.csv")
            os.makedirs(os.path.dirname(fn), exist_ok=True)
            new_file = not os.path.exists(fn) or os.path.getsize(fn) == 0
            with open(fn, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if new_file:
                    w.writerow(["Timestamp", "Message"])
                w.writerow([ts, msg])
        finally:
            file_log_queue.task_done()

threading.Thread(target=_log_writer, daemon=True).start()

# Status bar variable
status_var = tk.StringVar(value="Ready")

def log_message(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(ts, msg)
    root.after(0, lambda: status_var.set(f"[{time.strftime('%H:%M:%S')}]  {msg}"))
    if save_log_var.get():
        file_log_queue.put((ts, msg))

# ─────────────────────────────────────────────────────────────────────────────
#  Controller helpers
# ─────────────────────────────────────────────────────────────────────────────
def init_controller(com_port: str):
    with motor_lock():
        ok = serial_port.MoCtrCard_Initial(com_port) == serial_port.FUNRES_OK
    if ok:
        log_message(f"Controller initialized on {com_port}")
        resume_all_axes()
    else:
        log_message(f"Initialization FAILED on {com_port}")

def unload_controller():
    try:
        serial_port.MoCtrCard_Unload()
    except Exception:
        pass

def stop_all_axes():
    success = True
    with motor_lock():
        for ax in range(NUM_AXES):
            acc = read_single_param(ax, PARA_ACCEL) or DEFAULT_ACC
            if (serial_port.MoCtrCard_ParaStopAxisMov(System.Byte(ax), System.Single(acc))
                    != serial_port.FUNRES_OK):
                success = False
    log_message("All axes stopped" if success else "Error stopping some axes")

def read_single_param(ax: int, code: int):
    with motor_lock():
        try:
            arr_f = System.Array.CreateInstance(System.Single, 1)
            if (serial_port.MoCtrCard_ReadPara(
                    System.Byte(ax), System.Byte(code), arr_f)
                    == serial_port.FUNRES_OK):
                return float(arr_f[0])
        except Exception:
            pass
        arr_u = System.Array.CreateInstance(System.UInt32, 1)
        if (serial_port.MoCtrCard_ReadPara(
                System.Byte(ax), System.Byte(code), arr_u)
                == serial_port.FUNRES_OK):
            return float(arr_u[0])
    return None

def read_axis(ax: int):
    with motor_lock():
        arr = System.Array.CreateInstance(System.Single, 1)
        if (serial_port.MoCtrCard_GetAxisPos(System.Byte(ax), arr)
                != serial_port.FUNRES_OK):
            return None, None, None
        pos = float(arr[0])
    vel = read_single_param(ax, PARA_VELOCITY)
    acc = read_single_param(ax, PARA_ACCEL)
    return pos, vel, acc

# ─────────────────────────────────────────────────────────────────────────────
#  Refresh loop — 150 ms interval to reduce CPU hammering
# ─────────────────────────────────────────────────────────────────────────────
def refresh_ui():
    for ax in range(NUM_AXES):
        pos, vel, acc = read_axis(ax)
        if pos is None:
            continue
        pd = position_display[ax]
        pd.config(state="normal")
        pd.delete(0, tk.END)
        pd.insert(0, f"{pos:.3f}")
        pd.config(state="readonly")
        if not editing_flags[(ax, "velocity")]:
            e = entry_widgets[ax]["velocity"]
            e.delete(0, tk.END)
            e.insert(0, f"{(vel or 0):.3f}")
        if not editing_flags[(ax, "acceleration")]:
            e = entry_widgets[ax]["acceleration"]
            e.delete(0, tk.END)
            e.insert(0, f"{(acc or 0):.3f}")
    root.after(150, refresh_ui)

# ─────────────────────────────────────────────────────────────────────────────
#  Motion + param helpers
# ─────────────────────────────────────────────────────────────────────────────
def move_abs(ax):
    try:
        val = float(abs_input_widgets[ax].get().strip())
    except ValueError:
        log_message(f"Invalid abs input on {AXES[ax]['label']}")
        return
    resume_all_axes()
    with motor_lock():
        ok = (serial_port.MoCtrCard_MCrlAxisAbsMove(
            System.Byte(ax), System.Single(val)) == serial_port.FUNRES_OK)
    log_message(
        f"[ABS] {AXES[ax]['label']} → {val:.3f} {AXES[ax]['unit']}"
        if ok else f"[ABS] {AXES[ax]['label']} move FAILED")

def _move_rel_direct(ax, delta):
    """Send a relative move directly to the DLL — safe to call from any thread."""
    resume_all_axes()
    with motor_lock():
        ok = (serial_port.MoCtrCard_MCrlAxisRelMove(
            System.Byte(ax), System.Single(delta)) == serial_port.FUNRES_OK)
    log_message(
        f"[REL] {AXES[ax]['label']} Δ{delta:+.3f} {AXES[ax]['unit']}"
        if ok else f"[REL] {AXES[ax]['label']} move FAILED")

def move_rel(ax, sign):
    try:
        delta = float(rel_input_widgets[ax].get().strip()) * sign
    except ValueError:
        log_message(f"Invalid rel input on {AXES[ax]['label']}")
        return
    _move_rel_direct(ax, delta)

def send_param(ax, field):
    try:
        val = float(entry_widgets[ax][field].get().strip())
    except ValueError:
        log_message(f"Invalid {field} on {AXES[ax]['label']}")
        return
    unit_lbl = (AXES[ax]['vel_unit'] if field == 'velocity'
                else AXES[ax]['acc_unit'])
    if not messagebox.askyesno(
            "Confirm", f"Set {AXES[ax]['label']} {field} → {val:.3f} {unit_lbl}?"):
        return
    editing_flags[(ax, field)] = False
    stop_all_axes()
    code = PARA_VELOCITY if field == "velocity" else PARA_ACCEL
    with motor_lock():
        ok = (serial_port.MoCtrCard_SendPara(
            System.Byte(ax), System.Byte(code), System.Single(val))
            == serial_port.FUNRES_OK)
    log_message(
        f"[{field.upper()[:3]}] {AXES[ax]['label']} → {val:.3f}"
        if ok else f"[{field.upper()[:3]}] {AXES[ax]['label']} set FAILED")

def home_axis(ax):
    resume_all_axes()
    with motor_lock():
        ok = (serial_port.MoCtrCard_MCrlAxisAbsMove(
            System.Byte(ax), System.Single(0.0)) == serial_port.FUNRES_OK)
    log_message(f"[HOME] {AXES[ax]['label']} → 0"
                if ok else f"[HOME] {AXES[ax]['label']} FAILED")

def stop_axis(ax):
    acc = read_single_param(ax, PARA_ACCEL) or DEFAULT_ACC
    with motor_lock():
        ok = (serial_port.MoCtrCard_ParaStopAxisMov(System.Byte(ax), System.Single(acc))
              == serial_port.FUNRES_OK)
    log_message(f"[STOP] {AXES[ax]['label']}"
                if ok else f"[STOP] {AXES[ax]['label']} FAILED")

def reset_axis_coord(ax, reset_to: float = 0.0):
    with motor_lock():
        ok = (serial_port.MoCtrCard_ResetCoordinate(
            System.Byte(ax), System.Single(reset_to)) == serial_port.FUNRES_OK)
    log_message(f"[ZERO] {AXES[ax]['label']} ← {reset_to:.3f}"
                if ok else f"[ZERO] {AXES[ax]['label']} FAILED")

def reset_all_coords():
    with motor_lock():
        ok_all = all(
            serial_port.MoCtrCard_ResetCoordinate(
                System.Byte(ax), System.Single(0.0)) == serial_port.FUNRES_OK
            for ax in range(NUM_AXES))
    log_message("[ZERO] All axes ← 0"
                if ok_all else "[ZERO] Some axes failed")

def reset_controller():
    com = com_var.get()
    unload_controller()
    time.sleep(0.2)
    init_controller(com)

def set_defaults():
    stop_all_axes()
    for ax in range(NUM_AXES):
        for field, val in (("velocity", DEFAULT_VEL),
                           ("acceleration", DEFAULT_ACC)):
            e = entry_widgets[ax][field]
            e.delete(0, tk.END)
            e.insert(0, str(val))
        send_param(ax, "velocity")
        send_param(ax, "acceleration")

# ─────────────────────────────────────────────────────────────────────────────
#  Widget factory helpers
# ─────────────────────────────────────────────────────────────────────────────
def themed_entry(parent, width=10, readonly=False):
    e = tk.Entry(
        parent, width=width,
        font=FONT_MONO,
        bg=T["entry_bg"], fg=T["text"],
        insertbackground=T["accent"],
        relief="flat",
        highlightthickness=1,
        highlightbackground=T["border"],
        highlightcolor=T["accent"],
        disabledbackground=T["panel"],
        disabledforeground=T["text_dim"],
        readonlybackground=T["panel"],
    )
    if readonly:
        e.config(state="readonly")
    return e

def themed_button(parent, text, cmd, color=None, width=8):
    c = color or T["border"]
    return tk.Button(
        parent, text=text, command=cmd,
        font=FONT_MONO_SM,
        bg=T["panel"], fg=color or T["text"],
        activebackground=T["entry_act"], activeforeground=T["accent"],
        relief="flat",
        highlightthickness=1,
        highlightbackground=c,
        highlightcolor=c,
        cursor="hand2",
        width=width,
    )

def dim_label(parent, text):
    return tk.Label(parent, text=text, font=FONT_SMALL,
                    bg=T["panel"], fg=T["text_dim"])

def row_label(parent, text):
    return tk.Label(parent, text=text, font=FONT_MONO_SM,
                    bg=T["panel"], fg=T["text_dim"], width=4, anchor="e")

# ─────────────────────────────────────────────────────────────────────────────
#  Header bar
# ─────────────────────────────────────────────────────────────────────────────
hdr = tk.Frame(root, bg=T["panel"], height=48)
hdr.pack(fill="x", padx=0, pady=0)

tk.Label(hdr, text="6-Axis Motion Controller",
         font=("Segoe UI", 15, "bold"),
         bg=T["panel"], fg=T["text"]).pack(side="left", padx=20, pady=10)

tk.Label(hdr, text="WangLab  ·  MCC-6",
         font=("Segoe UI", 10),
         bg=T["panel"], fg=T["text_dim"]).pack(side="right", padx=20, pady=10)

# thin separator line
sep = tk.Frame(root, bg=T["border"], height=1)
sep.pack(fill="x")

# ─────────────────────────────────────────────────────────────────────────────
#  Main area  (6 axis cards  |  settings panel)
# ─────────────────────────────────────────────────────────────────────────────
body = tk.Frame(root, bg=T["bg"])
body.pack(fill="both", expand=True, padx=12, pady=10)

# Left: 2×3 grid of axis cards
axes_frame = tk.Frame(body, bg=T["bg"])
axes_frame.pack(side="left", fill="both", expand=True)

for col in range(3):
    axes_frame.grid_columnconfigure(col, weight=1, uniform="ax")
for row in range(2):
    axes_frame.grid_rowconfigure(row, weight=1, uniform="axr")

# Right: settings sidebar
sidebar = tk.Frame(body, bg=T["bg"], width=210)
sidebar.pack(side="right", fill="y", padx=(10, 0))
sidebar.pack_propagate(False)

# ─────────────────────────────────────────────────────────────────────────────
#  Axis card builder
# ─────────────────────────────────────────────────────────────────────────────
AXIS_COLORS = ["#1A73E8", "#2E7D32", "#E65100", "#6A1B9A", "#00838F", "#AD1457"]

def make_axis_card(parent, ax):
    cfg   = AXES[ax]
    color = AXIS_COLORS[ax]
    row_g = ax // 3
    col_g = ax % 3

    # Card frame
    card = tk.Frame(parent, bg=T["panel"],
                    highlightbackground=T["border"], highlightthickness=1)
    card.grid(row=row_g, column=col_g, padx=5, pady=5, sticky="nsew")
    card.grid_columnconfigure(1, weight=1)

    # Axis header strip
    hd = tk.Frame(card, bg=color, height=3)
    hd.pack(fill="x")

    # Title row
    title_row = tk.Frame(card, bg=T["panel"])
    title_row.pack(fill="x", padx=10, pady=(6, 2))
    tk.Label(title_row, text=cfg["label"], font=("Consolas", 22, "bold"),
             bg=T["panel"], fg=color).pack(side="left")
    tk.Label(title_row, text=f"AXIS  [{cfg['unit']}]",
             font=FONT_SMALL, bg=T["panel"], fg=T["text_dim"]).pack(
             side="left", padx=(6, 0), pady=(8, 0))

    def divider():
        tk.Frame(card, bg=T["border"], height=1).pack(fill="x", padx=8)

    # — Position (read-only, large) ————————————————————————
    pos_row = tk.Frame(card, bg=T["panel"])
    pos_row.pack(fill="x", padx=10, pady=(4, 2))
    dim_label(pos_row, "POS").pack(side="left")
    pos_e = themed_entry(pos_row, width=11, readonly=True)
    pos_e.config(font=("Consolas", 13, "bold"), fg=color,
                 readonlybackground=T["panel"])
    pos_e.pack(side="left", padx=(6, 2))
    dim_label(pos_row, cfg["unit"]).pack(side="left")
    position_display[ax] = pos_e

    divider()

    # — Abs move ————————————————————————————————————————————
    abs_row = tk.Frame(card, bg=T["panel"])
    abs_row.pack(fill="x", padx=10, pady=3)
    dim_label(abs_row, "ABS").pack(side="left")
    abs_e = themed_entry(abs_row, width=9)
    abs_e.pack(side="left", padx=(6, 2))
    dim_label(abs_row, cfg["unit"]).pack(side="left", padx=(0, 4))
    abs_input_widgets[ax] = abs_e
    themed_button(abs_row, "Go", lambda a=ax: move_abs(a),
                  color=T["accent3"], width=6).pack(side="left")

    # — Rel move ————————————————————————————————————————————
    rel_row = tk.Frame(card, bg=T["panel"])
    rel_row.pack(fill="x", padx=10, pady=3)
    dim_label(rel_row, "REL").pack(side="left")
    rel_e = themed_entry(rel_row, width=9)
    rel_e.pack(side="left", padx=(6, 2))
    dim_label(rel_row, cfg["unit"]).pack(side="left", padx=(0, 4))
    rel_input_widgets[ax] = rel_e
    themed_button(rel_row, "+", lambda a=ax: move_rel(a, +1),
                  color=T["accent3"], width=3).pack(side="left", padx=(0, 2))
    themed_button(rel_row, "−", lambda a=ax: move_rel(a, -1),
                  color=T["accent2"], width=3).pack(side="left")

    divider()

    # — Velocity ————————————————————————————————————————————
    vel_row = tk.Frame(card, bg=T["panel"])
    vel_row.pack(fill="x", padx=10, pady=3)
    dim_label(vel_row, "VEL").pack(side="left")
    vel_e = themed_entry(vel_row, width=9)
    vel_e.pack(side="left", padx=(6, 2))
    dim_label(vel_row, cfg["vel_unit"]).pack(side="left")
    vel_e.bind("<FocusIn>",
               lambda e, a=ax: editing_flags.__setitem__((a, "velocity"), True))
    vel_e.bind("<FocusOut>",
               lambda e, a=ax: editing_flags.__setitem__((a, "velocity"), False))
    vel_e.bind("<Return>", lambda e, a=ax: send_param(a, "velocity"))
    entry_widgets[ax]["velocity"] = vel_e

    # — Acceleration ————————————————————————————————————————
    acc_row = tk.Frame(card, bg=T["panel"])
    acc_row.pack(fill="x", padx=10, pady=3)
    dim_label(acc_row, "ACC").pack(side="left")
    acc_e = themed_entry(acc_row, width=9)
    acc_e.pack(side="left", padx=(6, 2))
    dim_label(acc_row, cfg["acc_unit"]).pack(side="left")
    acc_e.bind("<FocusIn>",
               lambda e, a=ax: editing_flags.__setitem__((a, "acceleration"), True))
    acc_e.bind("<FocusOut>",
               lambda e, a=ax: editing_flags.__setitem__((a, "acceleration"), False))
    acc_e.bind("<Return>", lambda e, a=ax: send_param(a, "acceleration"))
    entry_widgets[ax]["acceleration"] = acc_e

    divider()

    # — Per-axis buttons ————————————————————————————————————
    btn_row = tk.Frame(card, bg=T["panel"])
    btn_row.pack(fill="x", padx=10, pady=(4, 8))
    themed_button(btn_row, "Stop", lambda a=ax: stop_axis(a),
                  color=T["accent2"], width=7).pack(side="left", padx=(0, 4))
    themed_button(btn_row, "Zero", lambda a=ax: reset_axis_coord(a),
                  color="#E65100", width=7).pack(side="left", padx=(0, 4))
    themed_button(btn_row, "Home", lambda a=ax: home_axis(a),
                  color=T["accent3"], width=7).pack(side="left")

for ax in range(NUM_AXES):
    make_axis_card(axes_frame, ax)

# ─────────────────────────────────────────────────────────────────────────────
#  Settings sidebar
# ─────────────────────────────────────────────────────────────────────────────
def sidebar_section(title):
    tk.Label(sidebar, text=title, font=("Segoe UI", 9, "bold"),
             bg=T["bg"], fg=T["text_dim"]).pack(anchor="w", padx=4,
                                                pady=(12, 2))
    tk.Frame(sidebar, bg=T["border"], height=1).pack(fill="x", padx=4)

def sidebar_card():
    f = tk.Frame(sidebar, bg=T["panel"],
                 highlightbackground=T["border"], highlightthickness=1)
    f.pack(fill="x", pady=(4, 0))
    return f

# — Connection ——————————————————————————————————————————————
sidebar_section("CONNECTION")
conn_card = sidebar_card()

com_var = tk.StringVar(value=DEFAULT_COM_PORT, master=root)

com_row = tk.Frame(conn_card, bg=T["panel"])
com_row.pack(fill="x", padx=8, pady=(8, 4))
tk.Label(com_row, text="Port", font=FONT_SMALL,
         bg=T["panel"], fg=T["text_dim"]).pack(side="left")
com_entry = themed_entry(com_row, width=7)
com_entry.pack(side="right")
com_entry.insert(0, DEFAULT_COM_PORT)

def _sync_com(*a):
    com_var.set(com_entry.get())
com_entry.bind("<KeyRelease>", _sync_com)

themed_button(conn_card, "Connect",
              lambda: init_controller(com_var.get()),
              color=T["accent3"], width=18).pack(padx=8, pady=(0, 4), fill="x")
themed_button(conn_card, "Reset controller",
              reset_controller,
              color=T["accent"], width=18).pack(padx=8, pady=(0, 8), fill="x")

# — Global actions ——————————————————————————————————————————
sidebar_section("GLOBAL ACTIONS")
glob_card = sidebar_card()

themed_button(glob_card, "Stop all",
              stop_all_axes, color=T["accent2"], width=18).pack(
    padx=8, pady=(8, 4), fill="x")
themed_button(glob_card, "Zero all",
              reset_all_coords, color="#E65100", width=18).pack(
    padx=8, pady=(0, 4), fill="x")
themed_button(glob_card, "Set defaults",
              set_defaults, color=T["text_dim"], width=18).pack(
    padx=8, pady=(0, 8), fill="x")

# — Logging ——————————————————————————————————————————————
sidebar_section("LOGGING")
log_card = sidebar_card()

chk_row = tk.Frame(log_card, bg=T["panel"])
chk_row.pack(fill="x", padx=8, pady=(6, 4))
tk.Checkbutton(chk_row, text="Save to CSV", variable=save_log_var,
               font=FONT_SMALL,
               bg=T["panel"], fg=T["text"],
               selectcolor=T["entry_bg"],
               activebackground=T["panel"],
               activeforeground=T["accent"]).pack(side="left")

path_row = tk.Frame(log_card, bg=T["panel"])
path_row.pack(fill="x", padx=8, pady=(0, 4))
tk.Label(path_row, text="Path", font=FONT_SMALL,
         bg=T["panel"], fg=T["text_dim"]).pack(anchor="w")
path_e = themed_entry(log_card, width=22)
path_e.pack(fill="x", padx=8, pady=(0, 4))
path_e.insert(0, log_save_path.get())

def _sync_path(*a):
    log_save_path.set(path_e.get())
path_e.bind("<KeyRelease>", _sync_path)

def _browse():
    d = filedialog.askdirectory()
    if d:
        log_save_path.set(d)
        path_e.delete(0, tk.END)
        path_e.insert(0, d)

themed_button(log_card, "BROWSE …", _browse,
              color=T["text_dim"], width=18).pack(
    padx=8, pady=(0, 8), fill="x")

# ─────────────────────────────────────────────────────────────────────────────
#  Status bar
# ─────────────────────────────────────────────────────────────────────────────
tk.Frame(root, bg=T["border"], height=1).pack(fill="x")
status_bar = tk.Frame(root, bg=T["panel"], height=24)
status_bar.pack(fill="x")
tk.Label(status_bar, textvariable=status_var, font=FONT_SMALL,
         bg=T["panel"], fg=T["text_dim"], anchor="w").pack(
    side="left", padx=14, pady=3)
tk.Label(status_bar, text="Live",
         font=("Segoe UI", 9),
         bg=T["panel"], fg=T["accent3"]).pack(side="right", padx=14)

# ─────────────────────────────────────────────────────────────────────────────
#  Joystick loop
# ─────────────────────────────────────────────────────────────────────────────
DEADZONE = 0.2
# Timestamp (per axis) of last button-triggered rel move; _cont_axis will not
# stop that axis for 300 ms afterward so the move isn't immediately cancelled
# by spurious analog-stick release events.
_rel_move_ts = [0.0] * NUM_AXES
REL_MOVE_LOCKOUT = 0.30  # seconds

def _cont_axis(axis_id: int, direction: int):
    bid = System.Byte(axis_id)
    acc = read_single_param(axis_id, PARA_ACCEL) or DEFAULT_ACC
    with motor_lock():
        if direction == 0:
            # Skip stop if a rel-move was just issued for this axis
            if time.monotonic() - _rel_move_ts[axis_id] < REL_MOVE_LOCKOUT:
                return
            serial_port.MoCtrCard_ParaStopAxisMov(bid, System.Single(acc))
            return
        resume_all_axes()
        vel = JOYSTICK_SPEED * direction
        serial_port.MoCtrCard_MCrlAxisMoveAtSpd(
            bid, System.Single(vel), System.Single(acc))

def _joystick_loop():
    pygame.init()
    if pygame.joystick.get_count() < 1:
        log_message("[JOY] No joystick found")
        return

    js0 = pygame.joystick.Joystick(0); js0.init()
    log_message(f"[JOY] JS0: {js0.get_name()} → X/Y")
    js1 = None
    if pygame.joystick.get_count() > 1:
        js1 = pygame.joystick.Joystick(1); js1.init()
        log_message(f"[JOY] JS1: {js1.get_name()} → A/B/C")

    x_dir = y_dir = a_dir = b_dir = c_dir = 0

    # JS0 (Xbox 360) button map → (axis, step, sign)
    # Fine steps (±0.001): btn2=X+, btn3=X−, btn5=Y+, btn4=Y−
    # Coarse steps (±0.005): btn0=X+, btn1=X−, btn9=Y+, btn8=Y−
    JS0_BUTTON_MAP = {
        2: (0, 0.001, +1),   # X + 0.001
        3: (0, 0.001, -1),   # X - 0.001
        5: (1, 0.001, +1),   # Y + 0.001
        4: (1, 0.001, -1),   # Y - 0.001
        0: (0, 0.005, +1),   # X + 0.005
        1: (0, 0.005, -1),   # X - 0.005
        9: (1, 0.005, +1),   # Y + 0.005
        8: (1, 0.005, -1),   # Y - 0.005
    }

    # JS1 (T.16000M) button map → (axis, step, sign)
    JS1_BUTTON_MAP = {
        4:(3,0.025,+1), 9:(3,0.025,-1),
        5:(4,0.025,+1), 8:(4,0.025,-1),
        6:(5,0.025,+1), 7:(5,0.025,-1),
        12:(3,0.001,+1),13:(3,0.001,-1),
        11:(4,0.001,+1),14:(4,0.001,-1),
        10:(5,0.001,+1),15:(5,0.001,-1),
        0:(5,0.005,+1), 1:(5,0.005,-1),
    }
    clock = pygame.time.Clock()

    while root.winfo_exists():
        clock.tick(100)
        for evt in pygame.event.get():
            if evt.type == pygame.JOYHATMOTION and evt.joy == 0:
                # D-pad on JS0 → X/Y
                new_x = -evt.value[0]
                new_y = -evt.value[1]
                if new_x != x_dir: _cont_axis(0, new_x); x_dir = new_x
                if new_y != y_dir: _cont_axis(1, new_y); y_dir = new_y
            elif evt.type == pygame.JOYAXISMOTION and evt.joy == 0:
                # Analog stick on JS0 → X (axis 0) / Y (axis 1)
                if evt.axis == 0:
                    val = evt.value
                    new_x = int(val/abs(val)) if abs(val) > DEADZONE else 0
                    if new_x != x_dir: _cont_axis(0, new_x); x_dir = new_x
                elif evt.axis == 1:
                    val = evt.value
                    new_y = int(val/abs(val)) if abs(val) > DEADZONE else 0
                    if new_y != y_dir: _cont_axis(1, new_y); y_dir = new_y
            elif evt.type == pygame.JOYAXISMOTION and evt.joy == 1:
                if evt.axis == 0:
                    val = evt.value
                    new_a = int(val/abs(val)) if abs(val) > DEADZONE else 0
                    if new_a != a_dir: _cont_axis(3, new_a); a_dir = new_a
                elif evt.axis == 1:
                    val = evt.value
                    new_b = int(val/abs(val)) if abs(val) > DEADZONE else 0
                    if new_b != b_dir: _cont_axis(4, new_b); b_dir = new_b
                elif evt.axis == 2:
                    val = evt.value
                    new_c = int(val/abs(val)) if abs(val) > DEADZONE else 0
                    if new_c != c_dir: _cont_axis(5, new_c); c_dir = new_c
            elif evt.type == pygame.JOYBUTTONDOWN and evt.joy == 0:
                # Xbox controller buttons → X/Y relative moves
                if evt.button in JS0_BUTTON_MAP:
                    ax, step, sgn = JS0_BUTTON_MAP[evt.button]
                    delta = step * sgn
                    _rel_move_ts[ax] = time.monotonic()
                    root.after(0, lambda a=ax, s=step: (
                        rel_input_widgets[a].delete(0, tk.END),
                        rel_input_widgets[a].insert(0, f"{s}"),
                    ))
                    _move_rel_direct(ax, delta)
            elif evt.type == pygame.JOYBUTTONDOWN and evt.joy == 1:
                if evt.button in JS1_BUTTON_MAP:
                    ax, step, sgn = JS1_BUTTON_MAP[evt.button]
                    delta = step * sgn
                    _rel_move_ts[ax] = time.monotonic()
                    # Update the UI widget so the displayed value stays in sync,
                    # but fire the actual move directly — avoids root.after delay
                    # and eliminates the race where an analog-stick stop event
                    # arrives 10 ms later and cancels the move.
                    root.after(0, lambda a=ax, s=step: (
                        rel_input_widgets[a].delete(0, tk.END),
                        rel_input_widgets[a].insert(0, f"{s}"),
                    ))
                    _move_rel_direct(ax, delta)

    for ax in (0, 1, 3, 4, 5):
        acc = read_single_param(ax, PARA_ACCEL) or DEFAULT_ACC
        with motor_lock():
            serial_port.MoCtrCard_ParaStopAxisMov(System.Byte(ax), System.Single(acc))

threading.Thread(target=_joystick_loop, daemon=True).start()

# ─────────────────────────────────────────────────────────────────────────────
#  Start & cleanup
# ─────────────────────────────────────────────────────────────────────────────
root.after(150, refresh_ui)

def on_close():
    try:
        file_log_queue.join()
    except Exception:
        pass
    unload_controller()
    try:
        pygame.quit()
    except Exception:
        pass
    root.quit()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
