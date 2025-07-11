# ─────────────────────────────────────────────────────────────────────────────
#  Imports & constants
# ─────────────────────────────────────────────────────────────────────────────
import os, sys, time, csv, threading, datetime
import clr, System
import tkinter as tk
from tkinter import filedialog, messagebox
from flask import Flask, jsonify
import keyboard

DLL_PATH = r"C:\Users\WangLabAdmin\Desktop\DTS\MCC4DLL.dll"
LOG_ROOT = r"C:\Users\WangLabAdmin\Desktop\DTS"
Z_MIN, Z_MAX = 0.0, 17.0          #  ← new: hard limits for the Z axis (mm)
AXES = {
    0: dict(lbl='R', unit='°', vunit='°/s',  aunit='°/s²', v_def=0.5, a_def=0.5),
    1: dict(lbl='Z', unit='mm', vunit='mm/s', aunit='mm/s²', v_def=0.1, a_def=0.5),
}

clr.AddReference(DLL_PATH)
from SerialPortLibrary import SPLibClass
sp = SPLibClass()

# ─────────────────────────────────────────────────────────────────────────────
#  Tk & REST API bootstrap
# ─────────────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.option_add("*Font", ("Arial", 13)) 
root.title("ONWAY MOTION CONTROLLER")
root.geometry("1080x470")
root.configure(bg="#F0F0F0")

app = Flask(__name__)
api_state = {k: 0.0 for k in ('r_position','z_position','r_velocity','z_velocity',
                              'r_acceleration','z_acceleration')}
@app.route("/api/status")
def _(): return jsonify(api_state)
def start_api(): app.run("0.0.0.0", 5000, debug=False, use_reloader=False)

# ─────────────────────────────────────────────────────────────────────────────
#  Widgets & globals
# ─────────────────────────────────────────────────────────────────────────────
log_box          = None          # filled later
pos_disp         = {}            # axis_id → Entry
entry            = {ax:{'velocity':None,'acceleration':None} for ax in AXES}
abs_inp, rel_inp = {}, {}
edit_flag        = {(ax,p):False for ax in AXES for p in ('velocity','acceleration')}

save_log   = tk.BooleanVar(value=True, master=root)
log_path   = tk.StringVar(value=LOG_ROOT, master=root)
use_api    = tk.BooleanVar(value=True,  master=root)
kb_enable  = tk.BooleanVar(value=True,  master=root)

written_files = set()
repeat_job = repeat_delta = None   # keyboard jog state

# ─────────────────────────────────────────────────────────────────────────────
#  Utility
# ─────────────────────────────────────────────────────────────────────────────
def log(msg: str):
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line  = f"{stamp}\t{msg}"
    log_box.config(state=tk.NORMAL)
    log_box.insert(tk.END, line+"\n");  log_box.see(tk.END)
    log_box.config(state=tk.DISABLED)
    if save_log.get():
        fn = os.path.join(log_path.get(), f"log_{stamp[:7]}.csv")
        new = fn not in written_files
        with open(fn, "a", newline="") as f:
            csv.writer(f).writerow(["Timestamp","Message"] if new else [stamp,msg])
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
    edit_flag[(ax,kind)] = False
    try:  val = float(entry[ax][kind].get())
    except ValueError:
        log(f"[{kind[:3].upper()}] Invalid {kind} for axis {AXES[ax]['lbl']}")
        return
    _apply_param(ax, kind, val)

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
    root.after(100,refresh)

# ─────────────────────────────────────────────────────────────────────────────
#  Keyboard jog
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
#  Keyboard jog  —  velocity-mode start/stop
# ─────────────────────────────────────────────────────────────────────────────

# Per-axis jogging parameters
JOG = {
    0: dict(fast_v=1.0, fast_a=1.0, slow_v=0.2, slow_a=0.5),  # R  (°/s)
    1: dict(fast_v=0.5, fast_a=0.5, slow_v=0.1, slow_a=0.2),  # Z  (mm/s)
}

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


def _start_jog(ax: int, direction: int, slow: bool):

    if ax == 1:
        pos, _, _ = read_axis(1)
        if (direction > 0 and pos >= Z_MAX - _LIMIT_EPS) or \
           (direction < 0 and pos <= Z_MIN + _LIMIT_EPS):
            log("[VEL] Z jog blocked (limit reached)")
            return

    resume()  # clear any global pause
    v = JOG[ax]['slow_v' if slow else 'fast_v'] * direction
    a = JOG[ax]['slow_a' if slow else 'fast_a']
    ret = sp.MoCtrCard_MCrlAxisAtSpd(System.Byte(ax),
                                     System.Single(v),
                                     System.Single(a))
    lbl = AXES[ax]['lbl']
    if ret == sp.FUNRES_OK:
        log(f"[VEL] Jog {lbl} {'+' if direction>0 else '-'} at {abs(v)} {AXES[ax]['vunit']}")
        # ── kick off the guard loop if this is the Z axis ──
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
        # --- legacy fine-step shortcuts -----------------------------------------
    step = None
    if key.lower() == "num_lock":
        if shift and ctrl and alt:   step = -0.001
        elif shift and alt:          step = -0.005
        elif ctrl  and alt:          step = +0.001
        elif alt:                    step = +0.005
    elif key == "period" and ctrl:   step = +0.025
    elif key == "comma"  and ctrl:   step = -0.025

    if step is not None:
        _z_step(step)
        return "break"          # consume the keystroke, no jog started


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
    keys_held.discard(key)
    return "break"

# Bind the handlers
#root.bind("<KeyPress>",   key_down)
#root.bind("<KeyRelease>", key_up)

# ─────────────────────────────────────────────────────────────────────────────
#  UI builder
# ─────────────────────────────────────────────────────────────────────────────
def axis_frame(parent, ax):
    d=AXES[ax]; frm=tk.LabelFrame(parent,text=f"{d['lbl']} Axis",font=("Arial",14,'bold'),
                                  bg="#F0F0F0",bd=3); frm.pack(pady=10,fill='x')
    def row(): return tk.Frame(frm,bg="#F0F0F0");          #
    # Position + Abs/Rel
    r=row(); r.pack(fill='x',pady=5)
    tk.Label(r,text="Position:",bg="#F0F0F0").pack(side=tk.LEFT,padx=5)
    pos_disp[ax]=tk.Entry(r,width=7,state='readonly'); pos_disp[ax].pack(side=tk.LEFT,padx=5)
    tk.Label(r,text=d['unit'],bg="#F0F0F0").pack(side=tk.LEFT)
    # Abs
    tk.Label(r,text="Abs:",bg="#F0F0F0").pack(side=tk.LEFT,padx=10)
    abs_inp[ax]=tk.Entry(r,width=7); abs_inp[ax].pack(side=tk.LEFT); abs_inp[ax].insert(0, "0")
    tk.Label(r,text=d['unit'],bg="#F0F0F0").pack(side=tk.LEFT)
    tk.Button(r,text="Go",command=lambda a=ax:move_abs(a)).pack(side=tk.LEFT,padx=5)
    # Rel
    tk.Label(r,text="Rel:",bg="#F0F0F0").pack(side=tk.LEFT,padx=5)
    rel_inp[ax]=tk.Entry(r,width=7); rel_inp[ax].pack(side=tk.LEFT); rel_inp[ax].insert(0, "1") 
    tk.Label(r,text=d['unit'],bg="#F0F0F0").pack(side=tk.LEFT)
    tk.Button(r,text="+",width=2,command=lambda a=ax:move_rel(a,+1)).pack(side=tk.LEFT)
    tk.Button(r,text="-",width=2,command=lambda a=ax:move_rel(a,-1)).pack(side=tk.LEFT)
    # Velocity
    rv=row(); rv.pack(fill='x',pady=5)
    tk.Label(rv,text="Velocity:",bg="#F0F0F0").pack(side=tk.LEFT,padx=5)
    ev=tk.Entry(rv,width=7); ev.pack(side=tk.LEFT)
    ev.bind("<FocusIn>",  lambda e,a=ax:on_focus(e,a,'velocity',True))
    ev.bind("<FocusOut>", lambda e,a=ax:on_focus(e,a,'velocity',False))
    ev.bind("<Return>",   lambda e,a=ax:on_enter(e,a,'velocity'))
    tk.Label(rv,text=d['vunit'],bg="#F0F0F0").pack(side=tk.LEFT)
    entry[ax]['velocity']=ev
    # Acceleration
    ra=row(); ra.pack(fill='x',pady=5)
    tk.Label(ra,text="Acceleration:",bg="#F0F0F0").pack(side=tk.LEFT,padx=5)
    ea=tk.Entry(ra,width=7); ea.pack(side=tk.LEFT)
    ea.bind("<FocusIn>",  lambda e,a=ax:on_focus(e,a,'acceleration',True))
    ea.bind("<FocusOut>", lambda e,a=ax:on_focus(e,a,'acceleration',False))
    ea.bind("<Return>",   lambda e,a=ax:on_enter(e,a,'acceleration'))
    tk.Label(ra,text=d['aunit'],bg="#F0F0F0").pack(side=tk.LEFT)
    entry[ax]['acceleration']=ea
    # Home / Stop
    tk.Button(ra,text=f"Stop {d['lbl']}",command=lambda a=ax:stop_axis(a)).pack(side=tk.RIGHT,padx=5)
    tk.Button(ra, text=f"Home {d['lbl']}", command=lambda a=ax: home(a)).pack(side=tk.RIGHT, padx=5)

left=tk.Frame(root,bg="#F0F0F0"); left.pack(side=tk.LEFT,fill='both',expand=True,padx=10,pady=10)
for ax in AXES: axis_frame(left,ax)

settings=tk.LabelFrame(left,text="Settings",font=("Arial",14,'bold'),bg="#F0F0F0",bd=3)
settings.pack(fill='x')
top=tk.Frame(settings,bg="#F0F0F0"); top.pack(fill='x',pady=5)
tk.Checkbutton(top,text="Use API",variable=use_api,bg="#F0F0F0").pack(side=tk.LEFT,padx=10)
com_var=tk.StringVar(value="COM5")
tk.Label(top,text="COM Port:",bg="#F0F0F0").pack(side=tk.LEFT)
tk.Entry(top,textvariable=com_var,width=7).pack(side=tk.LEFT,padx=5)
tk.Button(top,text="Connect",
          command=lambda:log("✅ Initialized") if _ok(sp.MoCtrCard_Initial(com_var.get()))
          else log("❌ Init failed")).pack(side=tk.LEFT,padx=10)
tk.Button(top,text="Default",command=set_defaults).pack(side=tk.LEFT,padx=10)
tk.Checkbutton(top,text="Keyboard Ctrl",variable=kb_enable,bg="#F0F0F0").pack(side=tk.RIGHT,padx=10)
tk.Checkbutton(top,text="Save Log",variable=save_log,bg="#F0F0F0").pack(side=tk.RIGHT,padx=10)

bot=tk.Frame(settings,bg="#F0F0F0"); bot.pack(fill='x',pady=5)
tk.Label(bot,text="Log Path:",bg="#F0F0F0").pack(side=tk.LEFT,padx=10)
tk.Entry(bot,textvariable=log_path,width=38).pack(side=tk.LEFT)
tk.Button(bot,text="Browse",command=lambda:log_path.set(filedialog.askdirectory() or log_path.get())
          ).pack(side=tk.LEFT,padx=10)

# Right-side log view
log_frame=tk.Frame(root,bg="#F0F0F0",bd=2,relief=tk.SUNKEN); log_frame.pack(side=tk.RIGHT,fill='both',expand=True,padx=5,pady=5)
log_box=tk.Text(log_frame,width=50,state=tk.DISABLED,font=("Consolas",10)); log_box.pack(fill='both',expand=True)

# ─────────────────────────────────────────────────────────────────────────────
#  Main loop & shutdown
# ─────────────────────────────────────────────────────────────────────────────
#root.bind("<KeyPress>",  key_down)
#root.bind("<KeyRelease>",key_up)
root.after(100,refresh)
# ---- start global keyboard hook ----------------------------------
keyboard.hook(_global_press)          # fires for both down & up
keyboard.hook(_global_release)
threading.Thread(target=keyboard.wait, daemon=True).start()

if use_api.get(): threading.Thread(target=start_api,daemon=True).start()

def on_close():
    try: sp.MoCtrCard_Unload()
    finally:
        root.quit(); root.destroy(); os._exit(0)
root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()