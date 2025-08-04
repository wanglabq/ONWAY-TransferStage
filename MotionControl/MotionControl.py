# ─────────────────────────────────────────────────────────────────────────────
#  Imports & constants
# ─────────────────────────────────────────────────────────────────────────────
import os, sys, time, csv, threading, datetime
import clr, System
import tkinter as tk
from tkinter import filedialog, messagebox
from flask import Flask, jsonify
import keyboard
import configparser
import paho.mqtt.client as mqtt
import json



# ─── load INI ───────────────────────────────────────────────────────────
INI_PATH = os.path.join(os.path.dirname(__file__), "MotionConfig.ini")
cfg = configparser.ConfigParser()
if not cfg.read(INI_PATH, encoding="utf-8"):
    raise FileNotFoundError(f"Cannot read {INI_PATH}")

gi = cfg.getint; gf = cfg.getfloat; gs = cfg.get    # short-hand

DLL_PATH = gs("Paths", "dll_path")
LOG_ROOT = gs("Paths", "log_root")
LOG_ENCODING = cfg.get("Logging", "encoding", fallback="utf-8")
Z_MIN = gf("Z_Limits", "min")
Z_MAX = gf("Z_Limits", "max")

AXES = {
    0: dict(lbl=gs("Axis_R","label"), unit=gs("Axis_R","unit"),
            vunit=gs("Axis_R","vunit"), aunit=gs("Axis_R","aunit"),
            v_def=gf("Axis_R","v_default"), a_def=gf("Axis_R","a_default")),
    1: dict(lbl=gs("Axis_Z","label"), unit=gs("Axis_Z","unit"),
            vunit=gs("Axis_Z","vunit"), aunit=gs("Axis_Z","aunit"),
            v_def=gf("Axis_Z","v_default"), a_def=gf("Axis_Z","a_default")),
}

JOG = {
    0: dict(fast_v=gf("Axis_R","jog_fast_v"), fast_a=gf("Axis_R","jog_fast_a"),
            slow_v=gf("Axis_R","jog_slow_v"), slow_a=gf("Axis_R","jog_slow_a")),
    1: dict(fast_v=gf("Axis_Z","jog_fast_v"), fast_a=gf("Axis_Z","jog_fast_a"),
            slow_v=gf("Axis_Z","jog_slow_v"), slow_a=gf("Axis_Z","jog_slow_a")),
}

MQTT_ENABLED = cfg.getboolean("MQTT","enabled",fallback=False)
ACCENT_COLOR = "#2E7D32"      # same fresh green
WHITE_BG     = "#FFFFFF"
BTN_FONT     = ("Arial", 12)  # tweak size here
POP_FONT     = ("Arial", 10)
ICON_PATH = cfg.get("UI", "icon_path", fallback="")


clr.AddReference(DLL_PATH)
from SerialPortLibrary import SPLibClass
sp = SPLibClass()

# ─────────────────────────────────────────────────────────────────────────────
#  Tk & REST API bootstrap
# ─────────────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.option_add("*Font", ("Arial", 13))
root.title(cfg["General"]["gui_title"])
root.geometry(cfg["General"]["geometry"])
root.configure(bg=WHITE_BG)
if ICON_PATH and os.path.exists(ICON_PATH):
    try: root.iconbitmap(ICON_PATH)
    except Exception: pass


app = Flask(__name__)
api_state = {k: 0.0 for k in ('r_position','z_position','r_velocity','z_velocity',
                              'r_acceleration','z_acceleration')}
@app.route("/api/status")
def _(): return jsonify(api_state)
def start_api(): app.run("0.0.0.0", gi("General","api_port"), debug=False, use_reloader=False)

# ─────────────────────────────────────────────────────────────────────────────
#  Widgets & globals
# ─────────────────────────────────────────────────────────────────────────────
log_box          = None          # filled later
pos_disp         = {}            # axis_id → Entry
entry            = {ax:{'velocity':None,'acceleration':None} for ax in AXES}
abs_inp, rel_inp = {}, {}
edit_flag        = {(ax,p):False for ax in AXES for p in ('velocity','acceleration')}

save_log  = tk.BooleanVar(value=cfg.getboolean("General","save_log"), master=root)
log_path  = tk.StringVar(value=LOG_ROOT, master=root)
use_api   = tk.BooleanVar(value=cfg.getboolean("General","use_api"),  master=root)
kb_enable = tk.BooleanVar(value=cfg.getboolean("General","keyboard_ctrl"), master=root)


written_files = set()
repeat_job = repeat_delta = None   # keyboard jog state

# ─────────────────────────────────────────────────────────────────────────────
#  Utility
# ─────────────────────────────────────────────────────────────────────────────
def log(msg: str):
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line  = f"{stamp}\t{msg}"

    # ── GUI console ────────────────────────────────────────────
    log_box.config(state=tk.NORMAL)
    log_box.insert(tk.END, line + "\n")
    log_box.see(tk.END)
    log_box.config(state=tk.DISABLED)

    # ── file logging ───────────────────────────────────────────
    if save_log.get():
        fn = os.path.join(log_path.get(), f"log_{stamp[:7]}.csv")

        # >>> ADD THESE TWO LINES <<<
        folder = os.path.dirname(fn)
        os.makedirs(folder, exist_ok=True)     # make sure path exists

        new = fn not in written_files
        with open(fn, "a", newline="", encoding=LOG_ENCODING) as f:
            writer = csv.writer(f)
            if new: writer.writerow(["Timestamp", "Message"])
            writer.writerow([stamp, msg])
        written_files.add(fn)

def _ok(ret): return ret == sp.FUNRES_OK

def _clamp_z(val: float) -> float:
    """Keep Z target within [Z_MIN, Z_MAX]."""
    return max(Z_MIN, min(Z_MAX, val))

# ─────────────────────────────────────────────────────────────────────────────
#  Low-level controller helpers
# ─────────────────────────────────────────────────────────────────────────────
def pause():  _ok(sp.MoCtrCard_PauseAxisMov( System.Byte(255)))
def resume(): _ok(sp.MoCtrCard_ReStartAxisMov(System.Byte(255)))
def stop_axis(ax):
    resume(); _ok(sp.MoCtrCard_StopAxisMov(System.Byte(ax)))
    log(f"[STOP] {'RZ'[ax]} axis stopped.")

def read_axis(ax):
    def _get(idx):
        buf = System.Array.CreateInstance(System.Single,1)
        _ok(sp.MoCtrCard_ReadPara(System.Byte(ax),System.Byte(idx),buf))
        return buf[0]
    pos  = _get(0 if ax else 0)      # pos is param 0, but we have separate call:
    buf  = System.Array.CreateInstance(System.Single,1)
    _ok(sp.MoCtrCard_GetAxisPos(System.Byte(ax),buf)); pos = buf[0]
    vel, acc = _get(2), _get(3)
    return pos, vel, acc
# ─────────────────────────────────────────────────────────────────────────────
#  Small modal popup to confirm a new velocity / acceleration
# ─────────────────────────────────────────────────────────────────────────────
def _confirm_param_change(ax: int, kind: str, new_val: float):
    """White dialog with black outline + green buttons."""
    _, cur_vel, cur_acc = read_axis(ax)
    old_val = cur_vel if kind=='velocity' else cur_acc

    win = tk.Toplevel(root, bg=WHITE_BG, bd=1, relief="solid")
    win.title("Confirm");  win.grab_set()
    if ICON_PATH and os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    msg = (f"Axis {AXES[ax]['lbl']} — {kind.capitalize()} \n\n"
           f"{old_val:.3f}  →  {new_val:.3f}\n")
    tk.Label(win, text=msg, font=POP_FONT, bg=WHITE_BG)\
        .pack(padx=30, pady=20)

    btn_fr = tk.Frame(win, bg=WHITE_BG); btn_fr.pack(pady=10)
    for txt, cmd in (("CONFIRM", lambda:_apply_param(ax,kind,new_val)),
                     ("CANCEL",  lambda:None)):
        tk.Button(btn_fr, text=txt, width=8, **{
            "bg": WHITE_BG, "fg": ACCENT_COLOR, "font": BTN_FONT,
            "bd":1, "relief":"solid", "highlightthickness":1,
            "highlightbackground":"#000000",
            "command": lambda c=cmd: (c(), win.destroy())
        }).pack(side=tk.LEFT, padx=8)


    def _do_confirm():
        _apply_param(ax, kind, new_val)
        win.destroy()

    def _do_cancel():
        # restore entry box to the old value
        e = entry[ax][kind]
        e.delete(0, tk.END)
        e.insert(0, f"{old_val:.3f}")
        win.destroy()

    tk.Button(btns, text="CONFIRM", width=10, command=_do_confirm).pack(side=tk.LEFT, padx=10)
    tk.Button(btns, text="CANCEL",  width=10, command=_do_cancel ).pack(side=tk.LEFT, padx=10)

# ─────────────────────────────────────────────────────────────────────────────
#  Axis parameter callbacks
# ─────────────────────────────────────────────────────────────────────────────
def _apply_param(ax, kind, val):
    idx = 2 if kind=='velocity' else 3
    stop_axis(ax)
    if _ok(sp.MoCtrCard_SendPara(System.Byte(ax),System.Byte(idx),System.Single(val))):
        log(f"[{kind[:3].upper()}] Axis {AXES[ax]['lbl']} => {val:.3f} {AXES[ax][kind[0]+'unit']}")
    else:
        log(f"[{kind[:3].upper()}] Axis {AXES[ax]['lbl']} set FAILED")

def on_enter(event, ax, kind):
    """User pressed <Return> in a velocity/acceleration box."""
    edit_flag[(ax, kind)] = False
    try:
        val = float(entry[ax][kind].get())
    except ValueError:
        log(f"[{kind[:3].upper()}] Invalid {kind} for axis {AXES[ax]['lbl']}")
        return
    _confirm_param_change(ax, kind, val)

def on_focus(evt, ax, kind, state): edit_flag[(ax,kind)] = state

# ─────────────────────────────────────────────────────────────────────────────
#  Motion commands
# ─────────────────────────────────────────────────────────────────────────────
def move_abs(ax):
    txt = abs_inp[ax].get().strip() or "0"                   
    try: val = float(abs_inp[ax].get())
    except ValueError: return log(f"[ABS] bad input for {AXES[ax]['lbl']}")
    if ax == 1:                         # Z-axis clamps
        if val != _clamp_z(val):
            log("[ABS] Z target capped to range 0–17 mm")
        val = _clamp_z(val)
    resume()
    if _ok(sp.MoCtrCard_MCrlAxisAbsMove(System.Byte(ax),System.Single(val))):
        log(f"[ABS] Axis {AXES[ax]['lbl']} => {val:.3f} {AXES[ax]['unit']}")
def move_rel(ax, sgn):
    txt = rel_inp[ax].get().strip() or "0"   
    try: step = float(rel_inp[ax].get())*sgn
    except ValueError: return log(f"[REL] bad input for {AXES[ax]['lbl']}")
    pos,_,_ = read_axis(ax)
    if ax == 1:                         
        target = _clamp_z(pos + step)
        if target != pos + step:
            log("[REL] Z move limited to 0–17 mm")
        step = target - pos
    resume()
    if _ok(sp.MoCtrCard_MCrlAxisRelMove(System.Byte(ax),System.Single(step))):
        log(f"[REL] Axis {AXES[ax]['lbl']} move {step:+.3f} {AXES[ax]['unit']}")

def home(ax: int):
    """Send axis R (0) or Z (1) straight to 0.0, regardless of GUI fields."""
    resume()  # be sure the controller isn’t paused
    if _ok(sp.MoCtrCard_MCrlAxisAbsMove(System.Byte(ax), System.Single(0.0))):
        log(f"[HOME] Axis {AXES[ax]['lbl']} => 0.000 {AXES[ax]['unit']}")
    else:
        log(f"[HOME] Axis {AXES[ax]['lbl']} HOME FAILED")

def set_defaults():
    for ax,d in AXES.items():
        for kind,val in (('velocity',d['v_def']),('acceleration',d['a_def'])):
            entry[ax][kind].delete(0,tk.END); entry[ax][kind].insert(0,str(val))
            _apply_param(ax,kind,val)

# ─────────────────────────────────────────────────────────────────────────────
#  GUI refresh / API sync
# ─────────────────────────────────────────────────────────────────────────────
def refresh():
    for ax in AXES:
        pos,vel,acc = read_axis(ax)
        if pos_disp[ax]:
            pos_disp[ax].config(state='normal'); pos_disp[ax].delete(0,tk.END)
            pos_disp[ax].insert(0,f"{pos:.3f}"); pos_disp[ax].config(state='readonly')
        if not edit_flag[(ax,'velocity')]:
            e=entry[ax]['velocity']; e.delete(0,tk.END); e.insert(0,f"{vel:.3f}")
        if not edit_flag[(ax,'acceleration')]:
            e=entry[ax]['acceleration']; e.delete(0,tk.END); e.insert(0,f"{acc:.3f}")
        prefix = 'r_' if ax==0 else 'z_'
        api_state[prefix+'position'], api_state[prefix+'velocity'], api_state[prefix+'acceleration'] = pos,vel,acc
    if mqtt_mgr:                          # send telemetry if MQTT is active
        mqtt_mgr.publish({k: round(v, 3) for k, v in api_state.items()})


    root.after(gi("General", "refresh_ms"), refresh)

# ─────────────────────────────────────────────────────────────────────────────
#  Keyboard jog
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
#  Keyboard jog  —  velocity-mode start/stop
# ─────────────────────────────────────────────────────────────────────────────

# Per-axis jogging parameters

_z_dir = 0       
_z_job = None    
_LIMIT_EPS = 0.01  

def _guard_z_limit():
    global _z_job, _z_dir
    if _z_dir == 0:        
        _z_job = None
        return

    pos, _, _ = read_axis(1) 
    if (_z_dir > 0 and pos >= Z_MAX - _LIMIT_EPS) or \
       (_z_dir < 0 and pos <= Z_MIN + _LIMIT_EPS):
        _stop_jog(1)       
        _z_dir = 0
        _z_job = None
        return

    _z_job = root.after(50, _guard_z_limit)   

keys_held = set()    
# ------------------------------------------------------------------
#  Tap‑vs‑Hold bookkeeping for the six combo keys
# ------------------------------------------------------------------
press_info = {}          # keysym → {'t0': float, 'step': float}
TAP_MS     = 500         # ≤ 150 ms counts as a quick tap → single step
# ------------------------------------------------------------------
#  GLOBAL-HOTKEY HOOK  (keyboard package → Tk-compatible event stub)
# ------------------------------------------------------------------
def _make_fake_event(kb_event):
    """
    Convert keyboard.Event ('name', .event_type) → object with .keysym & .state
    so we can reuse the existing key_down / key_up functions.
    """
    class _E:             # minimalist stand-in for Tk event
        keysym: str
        state: int
    e = _E()
    # --- map key names ---
    name = kb_event.name
    if name == 'num lock':
        e.keysym = 'Num_Lock'
    elif name in ('.', 'dot', 'decimal'):
        e.keysym = 'period'      # NEW → matches old Tk “period”
    elif name in (',', 'comma'):
        e.keysym = 'comma'       # NEW → matches old Tk “comma”
    elif len(name) == 1:
        e.keysym = name          # ',', '.', etc. (kept for other combos)
    else:
        e.keysym = name.capitalize()   # 'up' → 'Up', etc.
    # --- modifier bitmask (same bits your code already expects) ---
    s = 0
    if keyboard.is_pressed('shift'): s |= 0x0001
    if keyboard.is_pressed('ctrl'):  s |= 0x0004
    if keyboard.is_pressed('alt'):   s |= 0x0008   # left-alt OR right-alt
    e.state = s
    return e

def _global_press(kb_event):
    if kb_event.event_type != 'down':      # we only want real key-presses
        return
    fake = _make_fake_event(kb_event)
    #   Dispatch into Tk's main thread ASAP (avoids cross-thread GUI calls)
    root.after_idle(lambda: key_down(fake))

def _global_release(kb_event):
    if kb_event.event_type != 'up':
        return
    fake = _make_fake_event(kb_event)
    root.after_idle(lambda: key_up(fake))

# ─────────────────────────────────────────────────────────────────────────────
#  Deferred start for tap‑vs‑hold keys
# ─────────────────────────────────────────────────────────────────────────────
def _begin_hold_jog(key):
    """Called 150 ms after key‑down if the key is still held."""
    info = press_info.get(key)
    if not info or key not in keys_held:          # released early → tap
        return
    # mark as started so key_up knows to stop jog, not single‑step
    info['started'] = True
    step = info['step']
    speed = abs(step)
    direction = 1 if step > 0 else -1
    _start_jog(1, direction, False,
               v_custom=speed, a_custom=0.5)
# ─────────────────────────────────────────────────────────────────────────────
#  Velocity‑mode jog start  (now supports optional custom speed/accel)
# ─────────────────────────────────────────────────────────────────────────────
def _start_jog(ax: int,
               direction: int,
               slow: bool,
               *,                   # everything after * must be passed by name
               v_custom: float | None = None,
               a_custom: float | None = None):
    """
    • ax        : 0 = R, 1 = Z
    • direction : +1 or –1
    • slow      : True → use JOG[…]['slow_*']  (ignored when v_custom given)
    • v_custom  : override speed in axis units / s     (None → use table)
    • a_custom  : override accel in axis units / s²    (None → use table)
    """

    # ── Z‑axis soft‑limit guard ────────────────────────────────────────────
    if ax == 1:
        pos, _, _ = read_axis(1)
        if (direction > 0 and pos >= Z_MAX - _LIMIT_EPS) or \
           (direction < 0 and pos <= Z_MIN + _LIMIT_EPS):
            log("[VEL] Z jog blocked (limit reached)")
            return

    resume()                              # clear any global pause / pause()

    # ── choose speed / acceleration ───────────────────────────────────────
    if v_custom is not None:              # NEW override path
        v = v_custom * direction
        a = a_custom if a_custom is not None else 0.5   # pick your default
    else:                                 # ORIGINAL fast / slow selector
        v = JOG[ax]['slow_v' if slow else 'fast_v'] * direction
        a = JOG[ax]['slow_a' if slow else 'fast_a']

    # ── send command ───────────────────────────────────────────────────────
    ret = sp.MoCtrCard_MCrlAxisAtSpd(System.Byte(ax),
                                     System.Single(v),
                                     System.Single(a))
    lbl = AXES[ax]['lbl']
    if ret == sp.FUNRES_OK:
        log(f"[VEL] Jog {lbl} {'+' if direction>0 else '-'} "
            f"at {abs(v)} {AXES[ax]['vunit']}")
        # kick off guard loop for Z
        if ax == 1:
            global _z_dir, _z_job
            _z_dir = direction
            if _z_job is None:
                _z_job = root.after(50, _guard_z_limit)
    else:
        log(f"[VEL] Jog {lbl} command FAILED")


def _stop_jog(ax: int):
    sp.MoCtrCard_StopAxisMov(System.Byte(ax))
    log(f"[VEL] Stop jog {AXES[ax]['lbl']}")
    if ax == 1:                        
        global _z_dir, _z_job          
        _z_dir = 0                      
        if _z_job is not None:            
            root.after_cancel(_z_job)     
            _z_job = None                 


# --- incremental-step helper (for the old 6 shortcuts) -----------------------
def _z_step(delta_mm: float):
    """Single relative move on the Z axis, clamped to 0–17 mm."""
    if delta_mm == 0:
        return
    pos, _, _ = read_axis(1)
    target = _clamp_z(pos + delta_mm)
    delta  = target - pos
    if delta == 0:
        return
    resume()
    if _ok(sp.MoCtrCard_MCrlAxisRelMove(System.Byte(1), System.Single(delta))):
        log(f"[STEP] Z => {'UP' if delta > 0 else 'DOWN'} {abs(delta):.3f} mm")

def key_down(event):
    if not kb_enable.get():
        return None
    key = event.keysym
    if key in keys_held:          # ignore auto-repeat
        return "break"
    shift = bool(event.state & 0x0001)
    ctrl  = bool(event.state & 0x0004)
    alt   = bool((event.state & 0x0008) or (event.state & 0x20000))
    # --- legacy fine‑step shortcuts  ➜  now tap‑OR‑hold ------------------
    # --- combo keys: tap = step, hold = continuous --------------------------
    step = None
    if key.lower() == "num_lock":
        if shift and ctrl and alt:   step = -0.001
        elif shift and alt:          step = -0.010
        elif ctrl  and alt:          step = +0.001
        elif alt:                    step = +0.010
    elif key == "period" and ctrl:   step = +0.025
    elif key == "comma"  and ctrl:   step = -0.025

    if step is not None:
        keys_held.add(key)           # track press
        press_info[key] = {'step': step, 'started': False}
        # schedule the hold‑action; if key is released before it fires,
        # key_up() will cancel it and treat as tap
        press_info[key]['job'] = root.after(
            TAP_MS, lambda k=key: _begin_hold_jog(k))
        return "break"


    if key == "Up":
        _start_jog(1, +1, shift); keys_held.add("Up")
    elif key == "Down":
        _start_jog(1, -1, shift); keys_held.add("Down")
    
    #elif key == "Right":
    #    _start_jog(0, +1, shift); keys_held.add("Right")
    #elif key == "Left":
    #    _start_jog(0, -1, shift); keys_held.add("Left")
    #elif key == "space" and shift and ctrl and alt:
    #    stop_axis(0); stop_axis(1)        # emergency stop
    return "break"
def key_up(event):
    key = event.keysym
    if key not in keys_held:
        return None
    if key in ("Up", "Down"):
        _stop_jog(1)
    elif key in ("Left", "Right"):
        _stop_jog(0)
    elif key in press_info:
        info = press_info.pop(key)
        # if jog never started, it was a quick tap → single step
        if not info['started']:
            root.after_cancel(info['job'])
            _z_step(info['step'])
        else:
            _stop_jog(1)             # continuous jog: stop on release
    keys_held.discard(key)
    return "break"


# Bind the handlers
#root.bind("<KeyPress>",   key_down)
#root.bind("<KeyRelease>", key_up)
# ─────────────────────────────────────────────────────────────────────────
#  MQTT Manager  (publishes telemetry + HA discovery)
# ─────────────────────────────────────────────────────────────────────────
class MQTTManager:
    def __init__(self, gui):
        self.gui   = gui
        if not MQTT_ENABLED:
            self.client = None
            return
        self.topic = cfg["MQTT"]["topic"]
        self.qos   = cfg.getint("MQTT","qos",fallback=0)
        self.retain= cfg.getboolean("MQTT","retain",fallback=False)

        self.client = mqtt.Client(client_id=cfg["MQTT"].get("client_id") or
                                               f"Motion_{os.getpid()}")
        user = cfg["MQTT"].get("username","")
        if user:
            self.client.username_pw_set(user, cfg["MQTT"].get("password",""))
        self.client.on_connect = self._on_connect
        self.client.connect(cfg["MQTT"]["host"], cfg.getint("MQTT","port"))
        self.client.loop_start()
        self.publish_discovery()

    # ---------- discovery on first connect ----------
    def _on_connect(self, client, userdata, flags, rc, *_):
        print("[MQTT] connected") if rc==0 else print("[MQTT] rc",rc)
        if rc==0:
            self.publish_discovery()

    def publish_discovery(self):
        prefix = cfg["MQTT"].get("discovery_prefix","homeassistant").rstrip('/')
        device = {
            "identifiers":  ["onway_motion"],
            "name":         "Onway Motion Controller",
            "manufacturer": "ONWAY",
            "model":        "MCC-4",
        }
        for axis,label in ((0,'r'),(1,'z')):
            for key in ("position","velocity","acceleration"):
                uid   = f"onway_{label}_{key}"
                topic = f"{prefix}/sensor/{uid}/config"
                payload = {
                    "name": f"Onway {label.upper()} {key}",
                    "state_topic": self.topic,
                    "value_template": f"{{{{ value_json.{label}_{key} }}}}",
                    "unique_id": uid,
                    "unit_of_measurement": AXES[axis]['unit' if key=="position" else
                                                                ('vunit' if key=="velocity" else 'aunit')],
                    "device_class": None,
                    "device": device
                }
                self.client.publish(topic,
                    json.dumps(payload, ensure_ascii=False),
                    retain=True)


    # ---------- publish telemetry ----------
    def publish(self, js_obj):
        if self.client:
            self.client.publish(self.topic, json.dumps(js_obj),
                                qos=self.qos, retain=self.retain)

    def stop(self):
        if self.client:
            self.client.loop_stop(); self.client.disconnect()

mqtt_mgr = None

# ──────────────────────────────────────────────────────────────
#  Configuration pop-up
# ──────────────────────────────────────────────────────────────
def show_config_dialog():
    dlg = tk.Toplevel(root, bg=WHITE_BG)
    dlg.title("Configuration"); dlg.grab_set()
    dlg.configure(bd=1, relief="solid")
    if ICON_PATH and os.path.exists(ICON_PATH):
        try: dlg.iconbitmap(ICON_PATH)
        except Exception: pass

    
    # local vars seeded from current values
    api_var  = tk.BooleanVar(value=use_api.get(), master=dlg)
    kb_var   = tk.BooleanVar(value=kb_enable.get(), master=dlg)
    save_var = tk.BooleanVar(value=save_log.get(), master=dlg)
    com_var2 = tk.StringVar(value=com_var.get(), master=dlg)
    path_var = tk.StringVar(value=log_path.get(), master=dlg)

    row = 0
    def _lbl(txt):
        tk.Label(dlg, text=txt, font=POP_FONT, bg=WHITE_BG)\
          .grid(row=row, column=0, sticky="e", padx=6, pady=4)

    _lbl("COM Port:");                       # row 0
    tk.Entry(dlg, textvariable=com_var2, font=POP_FONT, width=10)\
        .grid(row=row, column=1, sticky="w", padx=6, pady=4)

    row += 1; _lbl("Log Path:")              # row 1
    tk.Entry(dlg, textvariable=path_var, font=POP_FONT, width=32)\
        .grid(row=row, column=1, sticky="w", padx=6, pady=4)
    tk.Button(dlg, text="Browse", font=POP_FONT,
              command=lambda: path_var.set(filedialog.askdirectory()
                                            or path_var.get()))\
        .grid(row=row, column=2, sticky="w", padx=6, pady=4)

    row += 1                                 # row 2 check-boxes
    tk.Checkbutton(dlg, text="Enable API", variable=api_var,
                   font=POP_FONT, bg=WHITE_BG)\
        .grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=2)
    row += 1
    tk.Checkbutton(dlg, text="Keyboard Ctrl", variable=kb_var,
                   font=POP_FONT, bg=WHITE_BG)\
        .grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=2)
    row += 1
    tk.Checkbutton(dlg, text="Save Log", variable=save_var,
                   font=POP_FONT, bg=WHITE_BG)\
        .grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=2)

    # buttons
    row += 1
    btn_fr = tk.Frame(dlg, bg=WHITE_BG); btn_fr.grid(row=row, column=0,
                                                     columnspan=3, pady=8)
    def _apply():
        # live update
        use_api.set(api_var.get()); kb_enable.set(kb_var.get())
        save_log.set(save_var.get())
        com_var.set(com_var2.get().strip()); log_path.set(path_var.get())

        # write back to INI and save
        cfg["General"]["use_api"]       = str(api_var.get()).lower()
        cfg["General"]["keyboard_ctrl"] = str(kb_var.get()).lower()
        cfg["General"]["save_log"]      = str(save_var.get()).lower()
        cfg["General"]["com_port"]      = com_var.get()
        cfg["Paths"]["log_root"]        = log_path.get()
        with open(INI_PATH, "w", encoding="utf-8") as f:
            cfg.write(f)
        dlg.destroy()

    tk.Button(btn_fr, text="Apply",   **std_btn, width=6,
              command=_apply).pack(side=tk.LEFT, padx=6)
    tk.Button(btn_fr, text="Cancel",  **std_btn, width=6,
              command=dlg.destroy).pack(side=tk.LEFT, padx=6)

# ─────────────────────────────────────────────────────────────────────────────
#  UI builder
# ─────────────────────────────────────────────────────────────────────────────
def axis_frame(parent, ax):
    d=AXES[ax]; frm=tk.LabelFrame(parent,text=f"{d['lbl']} Axis",font=("Arial",14,'bold'),
                                  bg=WHITE_BG ,bd=3); frm.pack(pady=10,fill='x')
    def row(): return tk.Frame(frm,bg=WHITE_BG);          #
    # Position + Abs/Rel
    r=row(); r.pack(fill='x',pady=5)
    tk.Label(r,text="Position:",bg=WHITE_BG).pack(side=tk.LEFT,padx=5)
    pos_disp[ax]=tk.Entry(r,width=7,state='readonly'); pos_disp[ax].pack(side=tk.LEFT,padx=5)
    tk.Label(r,text=d['unit'],bg=WHITE_BG).pack(side=tk.LEFT)
    # Abs
    tk.Label(r,text="Abs:",bg=WHITE_BG).pack(side=tk.LEFT,padx=10)
    abs_inp[ax]=tk.Entry(r,width=7); abs_inp[ax].pack(side=tk.LEFT); abs_inp[ax].insert(0, "0")
    tk.Label(r,text=d['unit'],bg=WHITE_BG).pack(side=tk.LEFT)
    tk.Button(r,text="Go",command=lambda a=ax:move_abs(a)).pack(side=tk.LEFT,padx=5)
    # Rel
    tk.Label(r,text="Rel:",bg=WHITE_BG).pack(side=tk.LEFT,padx=5)
    rel_inp[ax]=tk.Entry(r,width=7); rel_inp[ax].pack(side=tk.LEFT); rel_inp[ax].insert(0, "1") 
    tk.Label(r,text=d['unit'],bg=WHITE_BG).pack(side=tk.LEFT)
    tk.Button(r,text="+",width=2,command=lambda a=ax:move_rel(a,+1)).pack(side=tk.LEFT)
    tk.Button(r,text="-",width=2,command=lambda a=ax:move_rel(a,-1)).pack(side=tk.LEFT)
    # Velocity
    rv=row(); rv.pack(fill='x',pady=5)
    tk.Label(rv,text="Velocity:",bg=WHITE_BG).pack(side=tk.LEFT,padx=5)
    ev=tk.Entry(rv,width=7); ev.pack(side=tk.LEFT)
    ev.bind("<FocusIn>",  lambda e,a=ax:on_focus(e,a,'velocity',True))
    ev.bind("<FocusOut>", lambda e,a=ax:on_focus(e,a,'velocity',False))
    ev.bind("<Return>",   lambda e,a=ax:on_enter(e,a,'velocity'))
    tk.Label(rv,text=d['vunit'],bg=WHITE_BG).pack(side=tk.LEFT)
    entry[ax]['velocity']=ev
    # Acceleration
    ra=row(); ra.pack(fill='x',pady=5)
    tk.Label(ra,text="Acceleration:",bg=WHITE_BG).pack(side=tk.LEFT,padx=5)
    ea=tk.Entry(ra,width=7); ea.pack(side=tk.LEFT)
    ea.bind("<FocusIn>",  lambda e,a=ax:on_focus(e,a,'acceleration',True))
    ea.bind("<FocusOut>", lambda e,a=ax:on_focus(e,a,'acceleration',False))
    ea.bind("<Return>",   lambda e,a=ax:on_enter(e,a,'acceleration'))
    tk.Label(ra,text=d['aunit'],bg=WHITE_BG).pack(side=tk.LEFT)
    entry[ax]['acceleration']=ea
    # Home / Stop
    tk.Button(ra,text=f"Stop {d['lbl']}",command=lambda a=ax:stop_axis(a)).pack(side=tk.RIGHT,padx=5)
    tk.Button(ra, text=f"Home {d['lbl']}", command=lambda a=ax: home(a)).pack(side=tk.RIGHT, padx=5)

left=tk.Frame(root,bg=WHITE_BG); left.pack(side=tk.LEFT,fill='both',expand=True,padx=10,pady=10)
for ax in AXES: axis_frame(left,ax)

settings=tk.LabelFrame(left,text="Settings",font=("Arial",14,'bold'),bg=WHITE_BG,bd=3)
settings.pack(fill='x')
top=tk.Frame(settings,bg=WHITE_BG); top.pack(fill='x',pady=5)
# ------------------------------------------------------------------
#  Helper for the green “Connect” button
# ------------------------------------------------------------------
def _connect_clicked():
    if _ok(sp.MoCtrCard_Initial(com_var.get())):
        log("✅ Initialized")
        if MQTT_ENABLED:
            global mqtt_mgr
            if mqtt_mgr is None:
                mqtt_mgr = MQTTManager(root)
    else:
        log("❌ Init failed")

# ── Settings top row  (all white, green-outline buttons) ────────────
std_btn = dict(
    bg=WHITE_BG, fg=ACCENT_COLOR, font=BTN_FONT,
    bd=1, relief="solid", highlightthickness=1, highlightbackground="#000000"
)

com_var = tk.StringVar(value=cfg["General"]["com_port"])
# ── Settings top row  (white bg, thin black outline, green text) ──
std_btn = dict(bg=WHITE_BG, fg=ACCENT_COLOR, font=BTN_FONT,
               bd=1, relief="solid", highlightthickness=1,
               highlightbackground="#000000")

com_var = tk.StringVar(value=cfg["General"]["com_port"])

tk.Button(top, text="Connect", **std_btn, command=_connect_clicked)\
   .pack(side=tk.LEFT, padx=6)
tk.Button(top, text="Default", **std_btn, command=set_defaults)\
   .pack(side=tk.LEFT, padx=6)
tk.Button(top, text="Configuration", **std_btn,
          command=show_config_dialog)\
   .pack(side=tk.LEFT, padx=6)

# Right-side log view
log_frame = tk.Frame(root, bg=WHITE_BG, bd=1, relief="solid")
log_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=5, pady=5)

log_box = tk.Text(log_frame, state=tk.DISABLED, width=48,
                  font=("Consolas",10), bg=WHITE_BG, relief="flat")
log_box.pack(fill="both", expand=True)


# ─────────────────────────────────────────────────────────────────────────────
#  Main loop & shutdown
# ─────────────────────────────────────────────────────────────────────────────
#root.bind("<KeyPress>",  key_down)
#root.bind("<KeyRelease>",key_up)
root.after(gi("General","refresh_ms"), refresh)
# ---- start global keyboard hook ----------------------------------
keyboard.hook(_global_press)          # fires for both down & up
keyboard.hook(_global_release)
threading.Thread(target=keyboard.wait, daemon=True).start()

if use_api.get(): threading.Thread(target=start_api,daemon=True).start()

def on_close():
    try:
        sp.MoCtrCard_Unload()
    finally:
        if mqtt_mgr:                    # stop MQTT loop nicely
            mqtt_mgr.stop()
        root.quit()
        root.destroy()
        os._exit(0)

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()