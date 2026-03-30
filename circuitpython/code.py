"""
SpacePad firmware  —  code.py
Raspberry Pi Pico / CircuitPython
"""
import time
import board
import usb_hid
import rotaryio
import keypad
import analogio
import busio
import json
import sys
import supervisor
import binascii
from adafruit_hid.keyboard import Keyboard
from adafruit_hid.keycode import Keycode
from adafruit_hid.consumer_control import ConsumerControl
from adafruit_hid.consumer_control_code import ConsumerControlCode
from adafruit_hid.mouse import Mouse

# Defined first so it's available throughout boot, including I²C init diagnostics
_stdout_write = sys.stdout.write   # local binding — avoids 3-deep global lookup

def send_json(obj):
    _stdout_write(json.dumps(obj) + "\n")

# ─────────────────────────────────────────────────────────────
#  1. KEYCODES + CONSTANTS
# ─────────────────────────────────────────────────────────────

KC = {
    "A": Keycode.A, "B": Keycode.B, "C": Keycode.C, "D": Keycode.D,
    "E": Keycode.E, "F": Keycode.F, "G": Keycode.G, "H": Keycode.H,
    "I": Keycode.I, "J": Keycode.J, "K": Keycode.K, "L": Keycode.L,
    "M": Keycode.M, "N": Keycode.N, "O": Keycode.O, "P": Keycode.P,
    "Q": Keycode.Q, "R": Keycode.R, "S": Keycode.S, "T": Keycode.T,
    "U": Keycode.U, "V": Keycode.V, "W": Keycode.W, "X": Keycode.X,
    "Y": Keycode.Y, "Z": Keycode.Z,
    "0": Keycode.ZERO,  "1": Keycode.ONE,   "2": Keycode.TWO,
    "3": Keycode.THREE, "4": Keycode.FOUR,  "5": Keycode.FIVE,
    "6": Keycode.SIX,   "7": Keycode.SEVEN, "8": Keycode.EIGHT,
    "9": Keycode.NINE,
    "F1":  Keycode.F1,  "F2":  Keycode.F2,  "F3":  Keycode.F3,
    "F4":  Keycode.F4,  "F5":  Keycode.F5,  "F6":  Keycode.F6,
    "F7":  Keycode.F7,  "F8":  Keycode.F8,  "F9":  Keycode.F9,
    "F10": Keycode.F10, "F11": Keycode.F11, "F12": Keycode.F12,
    "SHIFT": Keycode.SHIFT, "LEFT_SHIFT": Keycode.LEFT_SHIFT,
    "RIGHT_SHIFT": Keycode.RIGHT_SHIFT,
    "CTRL": Keycode.CONTROL, "LEFT_CTRL": Keycode.LEFT_CONTROL,
    "RIGHT_CTRL": Keycode.RIGHT_CONTROL, "RIGHT_CONTROL": Keycode.RIGHT_CONTROL,
    "ALT": Keycode.ALT, "LEFT_ALT": Keycode.LEFT_ALT,
    "RIGHT_ALT": Keycode.RIGHT_ALT,
    "GUI": Keycode.GUI, "CMD": Keycode.GUI, "WIN": Keycode.GUI,
    "ESCAPE": Keycode.ESCAPE, "SPACE": Keycode.SPACEBAR,
    "ENTER": Keycode.ENTER, "TAB": Keycode.TAB,
    "BACKSPACE": Keycode.BACKSPACE, "DELETE": Keycode.DELETE,
    "HOME": Keycode.HOME, "END": Keycode.END,
    "PAGE_UP": Keycode.PAGE_UP, "PAGE_DOWN": Keycode.PAGE_DOWN,
    "UP": Keycode.UP_ARROW, "DOWN": Keycode.DOWN_ARROW,
    "LEFT": Keycode.LEFT_ARROW, "RIGHT": Keycode.RIGHT_ARROW,
    "INSERT": Keycode.INSERT, "PRINT_SCREEN": Keycode.PRINT_SCREEN,
}

ENCODER_MODES = ["H_SCROLL","V_SCROLL","ZOOM","UNDO_REDO","NEXT_PREV_TAB","VOLUME"]
MOUSE_BUTTONS = {
    "LEFT":   Mouse.LEFT_BUTTON,
    "RIGHT":  Mouse.RIGHT_BUTTON,
    "MIDDLE": Mouse.MIDDLE_BUTTON,
}

# Key types (plain strings; compared by identity via 'is' after interning)
KT_NORMAL     = "normal"
KT_MO         = "mo"
KT_MOUSE_HOLD = "mouse_hold"
KT_ENC_MOD    = "enc_mod"

REPEAT_KEYS = frozenset((
    "UP","DOWN","LEFT","RIGHT",
    "PAGE_UP","PAGE_DOWN","HOME","END",
    "BACKSPACE","DELETE","SPACE",
))

ENC_SHORT = {
    "H_SCROLL":"H.SCR","V_SCROLL":"V.SCR",
    "ZOOM":"ZOOM","UNDO_REDO":"UNDO","NEXT_PREV_TAB":"TABS","VOLUME":"VOL",
}

SETTINGS_FILE = "settings.json"

# ─────────────────────────────────────────────────────────────
#  2. DEFAULT CONFIG
# ─────────────────────────────────────────────────────────────

def _blank_key():
    return {
        "tap":[], "hold":[], "tap_hold_enabled":False,
        "macro":None, "key_repeat":None, "key_type":KT_NORMAL,
        "mo_layer":None, "mouse_button":None, "enc_mod_factor":None,
    }

def _blank_layer(name="Layer"):
    return {
        "name":name,
        "keys":[_blank_key() for _ in range(25)],
        "enc1_mode":"H_SCROLL", "enc2_mode":"V_SCROLL",
        "enc1_sw":"RIGHT_CLICK", "enc2_sw":"HOME",
        "sm_active":False,
        "sm_orbit_mods":["SHIFT"],  # keys held during orbit (+ MMB always)
        "sm_pan_mods":[],           # keys held during pan (+ MMB always)
    }

def _default_layer():
    layer = _blank_layer("Default")
    # Windows standard shortcuts — Ctrl-based
    # Layout:  [0]  [1]  [2]   ·    ·
    #          [5]  [6]  [7]   ·    ·
    #          [10] [11] [12] [13] [14]
    #          [15] [16] [17] [18] [19]
    #           ·   [21] [22] [23] [24]
    defaults = [
        ["CTRL","Z"],     ["CTRL","Y"],   ["CTRL","A"],      [], [],
        ["CTRL","X"],     ["CTRL","C"],   ["CTRL","V"],      [], [],
        ["ESCAPE"],       ["CTRL","F"],   ["CTRL","S"],  ["DELETE"], ["BACKSPACE"],
        ["SHIFT"],        ["CTRL"],       ["UP"],        ["TAB"],   ["ENTER"],
        [],               ["LEFT"],       ["DOWN"],      ["RIGHT"], ["CTRL","SHIFT","S"],
    ]
    for i, tap in enumerate(defaults):
        layer["keys"][i]["tap"] = tap
    return layer

DEFAULT_CONFIG = {
    "layers":             [_default_layer()],
    "active_layer":       0,
    "sm_sensitivity":     15.0,
    "sm_deadzone":        100.0,
    "sm_z_threshold":     100.0,
    "sm_filter":          0.25,
    "sm_adapt":           0.003,
    "sm_accel":           True,
    "sm_accel_curve":     2.0,
    "sm_z_mode":          "ZOOM",
    "sm_orbit_enter_ms":  40,
    "sm_orbit_exit_ms":   80,
    "joy_deadzone":       2000,
    "joy_speed":          3.0,
    "joy_invert_x":       False,
    "joy_invert_y":       False,
    "joy_sw":             "LEFT_CLICK",
    "enc1_speed":         20,
    "enc1_invert":        False,
    "enc2_speed":         20,
    "enc2_invert":        False,
    "tap_hold_ms":        200,
    "btn_extra1":         "F6",
    "key_repeat_enabled": True,
    "key_repeat_delay_ms":400,
    "key_repeat_rate_ms": 50,
    "app_mappings":       [],
    "default_layer":      0,
}

# ─────────────────────────────────────────────────────────────
#  3. SETTINGS CACHE
#  All hot-path config reads go through these flat locals.
#  Call _sync_cache() once after any cfg mutation.
# ─────────────────────────────────────────────────────────────

class _C:
    """Flat cache of all settings accessed in the hot path."""
    __slots__ = (
        "sm_sensitivity","sm_deadzone","sm_z_threshold","sm_filter",
        "sm_adapt","sm_accel","sm_accel_curve","sm_z_mode","sm_inv_s10",
        "sm_orbit_enter","sm_orbit_exit","sm_active",
        "joy_deadzone","joy_speed","joy_invert_x","joy_invert_y",
        "enc1_speed","enc1_invert","enc2_speed","enc2_invert",
        "tap_hold_s","key_repeat_enabled",
        "key_repeat_delay_s","key_repeat_rate_s",
    )

SC = _C()   # single cache instance

def _sync_cache():
    SC.sm_sensitivity    = cfg["sm_sensitivity"]
    SC.sm_deadzone       = cfg["sm_deadzone"]
    SC.sm_z_threshold    = cfg["sm_z_threshold"]
    SC.sm_filter         = max(0.01, min(1.0, cfg["sm_filter"]))
    SC.sm_adapt          = max(0.0, min(0.02, cfg.get("sm_adapt", 0.003)))
    SC.sm_accel          = cfg["sm_accel"]
    SC.sm_accel_curve    = max(1.0, cfg["sm_accel_curve"])
    SC.sm_z_mode         = cfg["sm_z_mode"]
    SC.sm_inv_s10        = 1.0 / (max(0.1, cfg["sm_sensitivity"]) * 10)  # cached reciprocal
    SC.sm_orbit_enter    = cfg["sm_orbit_enter_ms"] / 1000.0   # stored as seconds
    SC.sm_orbit_exit     = cfg["sm_orbit_exit_ms"]  / 1000.0   # stored as seconds
    SC.sm_active         = False  # updated by _broadcast_layer on every layer change
    SC.joy_deadzone      = cfg["joy_deadzone"]
    SC.joy_speed         = cfg["joy_speed"]
    SC.joy_invert_x      = cfg["joy_invert_x"]
    SC.joy_invert_y      = cfg["joy_invert_y"]
    SC.enc1_speed        = cfg["enc1_speed"]
    SC.enc1_invert       = cfg["enc1_invert"]
    SC.enc2_speed        = cfg["enc2_speed"]
    SC.enc2_invert       = cfg["enc2_invert"]
    SC.tap_hold_s        = cfg["tap_hold_ms"] * 0.001
    SC.key_repeat_enabled= cfg["key_repeat_enabled"]
    SC.key_repeat_delay_s= cfg["key_repeat_delay_ms"] * 0.001
    SC.key_repeat_rate_s = cfg["key_repeat_rate_ms"] * 0.001

# ─────────────────────────────────────────────────────────────
#  4. PERSISTENCE  (CRC32 checksum, single file)
# ─────────────────────────────────────────────────────────────

def _crc(s):
    return binascii.crc32(s.encode()) & 0xFFFFFFFF

def _deep_copy_cfg(src):
    """Faster than json.loads(json.dumps(src)) — avoids double encode/decode."""
    if isinstance(src, dict):
        return {k: _deep_copy_cfg(v) for k, v in src.items()}
    if isinstance(src, list):
        return [_deep_copy_cfg(v) for v in src]
    return src

_settings_status = "defaults"

def load_config():
    global _settings_status
    base = _deep_copy_cfg(DEFAULT_CONFIG)
    try:
        with open(SETTINGS_FILE) as f:
            wrapper = json.load(f)
        data_str   = wrapper.get("data", "")
        stored_crc = wrapper.get("crc", -1)
        if stored_crc != _crc(data_str):
            _settings_status = "corrupt"
            return base
        saved = json.loads(data_str)
        if isinstance(saved, dict):
            base.update(saved)
            _settings_status = "loaded"
        else:
            _settings_status = "corrupt"
    except OSError:
        _settings_status = "new"
    except Exception:
        _settings_status = "error"
    return base

def save_config(cfg):
    try:
        data = json.dumps(cfg)
        with open(SETTINGS_FILE, "w") as f:
            json.dump({"v":1, "crc":_crc(data), "data":data}, f)
        return True
    except Exception as e:
        send_json({"event":"save_error","detail":str(e)})
        return False

cfg = load_config()

# ─────────────────────────────────────────────────────────────
#  5. HID + INPUT SETUP
# ─────────────────────────────────────────────────────────────

kbd   = Keyboard(usb_hid.devices)
cc    = ConsumerControl(usb_hid.devices)
mouse = Mouse(usb_hid.devices)

ROW_PINS = (board.GP0, board.GP1, board.GP2, board.GP3, board.GP4)
COL_PINS = (board.GP9, board.GP8, board.GP7, board.GP6, board.GP5)
ENC1_CLK, ENC1_DT, ENC1_SW = board.GP10, board.GP11, board.GP12
ENC2_CLK, ENC2_DT, ENC2_SW = board.GP13, board.GP14, board.GP15
JOY_X_PIN      = board.GP27
JOY_Y_PIN      = board.GP26
JOY_SW_PIN     = board.GP22
BTN_EXTRA1_PIN = board.GP18
BTN_EXTRA2_PIN = board.GP19
# MLX90393 — dedicated hardware I²C bus (GP16=SDA, GP17=SCL)
I2C_SDA, I2C_SCL = board.GP16, board.GP17
# SSD1306 OLED — separate bitbang I²C bus (GP20=SDA, GP21=SCL)
OLED_SDA, OLED_SCL = board.GP20, board.GP21

keys = keypad.KeyMatrix(ROW_PINS, COL_PINS, columns_to_anodes=False)
direct_keys = keypad.Keys(
    (BTN_EXTRA1_PIN, BTN_EXTRA2_PIN, JOY_SW_PIN),
    value_when_pressed=False, pull=True,
)
encoder1 = rotaryio.IncrementalEncoder(ENC1_CLK, ENC1_DT)
encoder2 = rotaryio.IncrementalEncoder(ENC2_CLK, ENC2_DT)
last_pos1 = last_pos2 = 0
encoder_switches = keypad.Keys(
    (ENC1_SW, ENC2_SW), value_when_pressed=False, pull=True,
)
joy_x = analogio.AnalogIn(JOY_X_PIN)
joy_y = analogio.AnalogIn(JOY_Y_PIN)

# ─────────────────────────────────────────────────────────────
#  6. I2C BUSES + MLX90393 + SSD1306 OLED
#  Bus 0 (hardware): GP16/GP17 — MLX90393 magnetometer @ 0x0C
#  Bus 1 (bitbang):  GP20/GP21 — SSD1306 OLED @ 0x3C
#  Separate buses = zero contention between MLX reads and OLED draws
# ─────────────────────────────────────────────────────────────

i2c      = None
i2c_oled = None
mlx      = None
oled_hw  = None
# Three flat variables instead of a list — avoids index lookups on every sensor tick
_mlx_ox = 0.0
_mlx_oy = 0.0
_mlx_oz = 0.0
mlx_ok = oled_ok = False

# Pre-allocated MLX I²C buffers — avoids bytearray allocation at 50 Hz
_MLX_CMD_MEAS  = bytearray([0x3E])   # single measurement request (X|Y|Z)
_MLX_CMD_READ  = bytearray([0x4E])   # read XYZ result
_MLX_STATUS_BUF = bytearray(1)       # reused for status polls
_MLX_DATA_BUF   = bytearray(7)       # reused for XYZ result reads

def _s16(hi, lo):
    """Signed 16-bit from two bytes — used for MLX XYZ decode."""
    v = (hi << 8) | lo
    return v - 65536 if v >= 32768 else v

try:
    i2c = busio.I2C(I2C_SCL, I2C_SDA, frequency=100_000)
    send_json({"event": "i2c_mlx_ok", "scl": "GP17", "sda": "GP16"})
except Exception as e:
    send_json({"event": "i2c_mlx_error", "detail": str(e)})

# OLED on separate bitbang I2C — doesn't block the MLX bus
try:
    import bitbangio
    i2c_oled = bitbangio.I2C(OLED_SCL, OLED_SDA)
    send_json({"event": "i2c_oled_ok", "scl": "GP21", "sda": "GP20"})
except Exception as e:
    send_json({"event": "i2c_oled_error", "detail": str(e)})

if i2c:
    # Scan MLX bus
    try:
        while not i2c.try_lock():
            pass
        _found = i2c.scan()
        i2c.unlock()
        send_json({"event": "i2c_scan", "bus": "mlx", "found": [hex(x) for x in _found]})
    except Exception as e:
        send_json({"event": "i2c_scan_error", "detail": str(e)})

    # MLX90393 — direct I²C driver
    # Uses separate write + read transactions (STOP between) — the MLX90393 does not
    # support repeated START so writeto_then_readfrom cannot be used.
    try:
        _MLX_ADDR = 0x0C

        def _mlx_write(bus, addr, write_buf):
            while not bus.try_lock(): pass
            try:
                bus.writeto(addr, write_buf)
            finally:
                bus.unlock()

        def _mlx_read(bus, addr, buf):
            """Read into pre-allocated buffer — no allocation."""
            while not bus.try_lock(): pass
            try:
                bus.readfrom_into(addr, buf)
            finally:
                bus.unlock()

        def _mlx_transceive(bus, addr, write_buf, read_buf):
            _mlx_write(bus, addr, write_buf)
            time.sleep(0.005)
            _mlx_read(bus, addr, read_buf)

        def _mlx_read_xyz(bus, addr):
            # Single measurement request using pre-allocated command buffer
            _mlx_write(bus, addr, _MLX_CMD_MEAS)
            time.sleep(0.01)
            for _ in range(20):
                try:
                    _mlx_read(bus, addr, _MLX_STATUS_BUF)
                    if _MLX_STATUS_BUF[0] & 0x01:
                        break
                except BaseException:
                    pass
                time.sleep(0.005)
            _mlx_transceive(bus, addr, _MLX_CMD_READ, _MLX_DATA_BUF)
            return (_s16(_MLX_DATA_BUF[1], _MLX_DATA_BUF[2]),
                    _s16(_MLX_DATA_BUF[3], _MLX_DATA_BUF[4]),
                    _s16(_MLX_DATA_BUF[5], _MLX_DATA_BUF[6]))

        # Test read
        x, y, z = _mlx_read_xyz(i2c, _MLX_ADDR)
        send_json({"event": "mlx_found", "address": hex(_MLX_ADDR), "test_xyz": [x, y, z]})

        # Calibration — average 20 samples at rest
        xs = ys = zs = 0.0
        for _ in range(20):
            x, y, z = _mlx_read_xyz(i2c, _MLX_ADDR)
            xs += x; ys += y; zs += z
            time.sleep(0.01)
        _mlx_ox = xs / 20
        _mlx_oy = ys / 20
        _mlx_oz = zs / 20

        # Non-blocking MLX reader — state machine so main loop never sleeps waiting
        # State 0 (IDLE):    send measurement request, record time, go to WAITING
        # State 1 (WAITING): poll DRDY each tick; when ready read data, go back to IDLE
        _i2c_ref  = i2c
        _addr_ref = _MLX_ADDR

        class _MLXReader:
            __slots__ = ("x","y","z","_state","_req_time","_ok_count","_err_count")
            def __init__(self):
                self.x = self.y = self.z = 0
                self._state    = 0       # 0=IDLE, 1=WAITING
                self._req_time = 0.0
                self._ok_count  = 0
                self._err_count = 0

            def tick(self, now):
                """Call every main loop iteration. Returns True when new XYZ is ready."""
                if self._state == 0:
                    try:
                        _mlx_write(_i2c_ref, _addr_ref, _MLX_CMD_MEAS)
                        self._req_time = now
                        self._state = 1
                    except BaseException as e:
                        self._err_count += 1
                        if self._err_count <= 5:
                            send_json({"event":"mlx_tick_err","state":0,"detail":str(e),"n":self._err_count})
                    return False
                else:
                    # MLX90393 needs ~10ms for conversion, then 5ms between polls
                    elapsed = now - self._req_time
                    if elapsed < 0.012:
                        return False   # too early — measurement not ready yet
                    try:
                        _mlx_read(_i2c_ref, _addr_ref, _MLX_STATUS_BUF)
                        if _MLX_STATUS_BUF[0] & 0x01:
                            _mlx_transceive(_i2c_ref, _addr_ref, _MLX_CMD_READ, _MLX_DATA_BUF)
                            self.x = _s16(_MLX_DATA_BUF[1], _MLX_DATA_BUF[2])
                            self.y = _s16(_MLX_DATA_BUF[3], _MLX_DATA_BUF[4])
                            self.z = _s16(_MLX_DATA_BUF[5], _MLX_DATA_BUF[6])
                            self._state = 0
                            self._ok_count += 1
                            return True
                        # DRDY not set — wait at least 5ms before next poll
                        self._req_time = now - 0.007   # so next check is ~5ms later
                    except BaseException as e:
                        self._err_count += 1
                        if self._err_count <= 5:
                            send_json({"event":"mlx_tick_err","state":1,"detail":str(e),"n":self._err_count})
                    if elapsed > 0.2:
                        self._state = 0
                    return False

        mlx = _MLXReader()
        mlx_ok = True

    except BaseException as e:
        send_json({"event": "mlx_error", "detail": str(e)})
        mlx = None

    # SSD1306 OLED — on separate bitbang I2C bus (GP20/GP21)
    if i2c_oled:
        try:
            import adafruit_ssd1306
            oled_hw = adafruit_ssd1306.SSD1306_I2C(128, 32, i2c_oled, addr=0x3C)
            oled_hw.fill(0)
            oled_hw.show()
            oled_ok = True
            send_json({"event": "oled_ok"})
        except Exception as e:
            send_json({"event": "oled_error", "detail": str(e)})

# ─────────────────────────────────────────────────────────────
#  7. SPACE MOUSE
# ─────────────────────────────────────────────────────────────

class SpaceMouse:
    __slots__ = ("fx","fy","fz","is_orbiting","is_panning",
                 "_above_since","_below_since","_zoom_accum",
                 "_orbit_kc","_pan_kc")
    def __init__(self):
        self.fx = self.fy = self.fz = 0.0
        self.is_orbiting = self.is_panning = False
        self._above_since = self._below_since = None
        self._zoom_accum  = 0.0
        self._orbit_kc = []   # resolved keycodes held during orbit
        self._pan_kc   = []   # resolved keycodes held during pan

    def recalibrate(self):
        global _mlx_ox, _mlx_oy, _mlx_oz
        if not mlx: return None
        xs = ys = zs = 0.0
        for _ in range(20):
            x, y, z = _mlx_read_xyz(_i2c_ref, _addr_ref)
            xs += x; ys += y; zs += z
            time.sleep(0.01)
        _mlx_ox = xs / 20
        _mlx_oy = ys / 20
        _mlx_oz = zs / 20
        self.fx = self.fy = self.fz = 0.0
        return [_mlx_ox, _mlx_oy, _mlx_oz]

    def safety_release(self):
        if self.is_orbiting:
            mouse.release(Mouse.MIDDLE_BUTTON)
            for kc in self._orbit_kc:
                kbd.release(kc)
            self._orbit_kc = []
            self.is_orbiting = False
        if self.is_panning:
            mouse.release(Mouse.MIDDLE_BUTTON)
            for kc in self._pan_kc:
                kbd.release(kc)
            self._pan_kc = []
            self.is_panning = False
        self.fx = self.fy = self.fz = 0.0
        self._above_since = self._below_since = None
        self._zoom_accum  = 0.0

    def update(self, raw_x, raw_y, raw_z, now):
        base  = SC.sm_filter   # minimum alpha — user's "smoothness" setting
        adapt = SC.sm_adapt    # how aggressively alpha scales with movement

        # X axis — adaptive alpha based on deviation from filtered value
        dx = (raw_x - _mlx_ox) - self.fx
        ax = base + abs(dx) * adapt
        if ax > 0.8: ax = 0.8
        self.fx += ax * dx

        # Y axis
        dy = (raw_y - _mlx_oy) - self.fy
        ay = base + abs(dy) * adapt
        if ay > 0.8: ay = 0.8
        self.fy += ay * dy

        # Z axis
        dz = (raw_z - _mlx_oz) - self.fz
        az = base + abs(dz) * adapt
        if az > 0.8: az = 0.8
        self.fz += az * dz

        self._process(now)

    def _accel(self, v, dz, inv_sens10):
        """Compute mouse delta.  inv_sens10 = 1.0 / (sensitivity * 10), pre-computed."""
        av = abs(v)
        if av <= dz: return 0
        norm = (av - dz) * inv_sens10
        if norm > 1.0: norm = 1.0
        if SC.sm_accel:
            norm = norm ** SC.sm_accel_curve
        delta = int(norm * 127)
        if delta > 127: delta = 127
        return delta if v > 0 else -delta

    def _process(self, now):
        lay   = _active_layer()
        dz    = SC.sm_deadzone
        zt    = SC.sm_z_threshold
        sens  = SC.sm_sensitivity
        enter = SC.sm_orbit_enter
        exit_ = SC.sm_orbit_exit
        fx    = self.fx
        fy    = self.fy

        xy_active = abs(fx) > dz or abs(fy) > dz
        inv_s10 = SC.sm_inv_s10

        if xy_active:
            self._below_since = None
            if self._above_since is None:
                self._above_since = now
            elif now - self._above_since >= enter:
                if not self.is_orbiting:
                    # Press orbit modifiers from layer config
                    mods = lay.get("sm_orbit_mods", ["SHIFT"]) if lay else ["SHIFT"]
                    self._orbit_kc = resolve(mods)
                    for kc in self._orbit_kc:
                        kbd.press(kc)
                    mouse.press(Mouse.MIDDLE_BUTTON)
                    self.is_orbiting = True
        else:
            self._above_since = None
            if self.is_orbiting:
                if self._below_since is None:
                    self._below_since = now
                elif now - self._below_since >= exit_:
                    mouse.release(Mouse.MIDDLE_BUTTON)
                    for kc in self._orbit_kc:
                        kbd.release(kc)
                    self._orbit_kc = []
                    self.is_orbiting = False
                    self._below_since = None

        if self.is_orbiting:
            mx = self._accel(fx, dz, inv_s10)
            my = self._accel(fy, dz, inv_s10) * -1
            if mx or my:
                mouse.move(x=mx, y=my)
            return

        z_mode = SC.sm_z_mode
        fz = self.fz
        if z_mode == "ZOOM":
            if abs(fz) > zt:
                self._zoom_accum += fz / (sens * 20)
                ticks = int(self._zoom_accum)
                if ticks:
                    if ticks > 5: ticks = 5
                    elif ticks < -5: ticks = -5
                    mouse.move(wheel=ticks)
                    self._zoom_accum -= ticks
            else:
                self._zoom_accum *= 0.8
        elif z_mode == "PAN":
            if abs(fz) > zt:
                if not self.is_panning:
                    # Press pan modifiers from layer config
                    mods = lay.get("sm_pan_mods", []) if lay else []
                    self._pan_kc = resolve(mods)
                    for kc in self._pan_kc:
                        kbd.press(kc)
                    mouse.press(Mouse.MIDDLE_BUTTON)
                    self.is_panning = True
                delta = int(fz / sens) * -1
                if delta > 127: delta = 127
                elif delta < -127: delta = -127
                mouse.move(y=delta)
            else:
                if self.is_panning:
                    mouse.release(Mouse.MIDDLE_BUTTON)
                    for kc in self._pan_kc:
                        kbd.release(kc)
                    self._pan_kc = []
                    self.is_panning = False

sm = SpaceMouse()

# ─────────────────────────────────────────────────────────────
#  8. OLED MANAGER
#  NoOpOLED avoids a branch on every loop tick when no OLED present.
# ─────────────────────────────────────────────────────────────

class NoOpOLED:
    """Drop-in replacement when no OLED is connected — all calls are free."""
    __slots__ = ()
    def flash(self, msg, ms=1200): pass
    def mark_dirty(self): pass
    def update(self): pass

class OLEDManager:
    __slots__ = ("hw","_flash_until","_flash_msg","_dirty","_last_draw")
    MIN_INTERVAL = 0.15   # ~7 Hz — own bus, no contention with MLX

    def __init__(self, hw):
        self.hw = hw
        self._flash_until = 0.0
        self._flash_msg   = ""
        self._dirty       = True
        self._last_draw   = 0.0

    def flash(self, msg, ms=1200):
        self._flash_msg   = msg[:21]
        self._flash_until = time.monotonic() + ms / 1000
        self._dirty = True

    def mark_dirty(self):
        self._dirty = True

    def update(self):
        if not self._dirty: return
        now = time.monotonic()
        if now - self._last_draw < self.MIN_INTERVAL: return
        hw = self.hw
        try:
            hw.fill(0)
            if now < self._flash_until:
                hw.text(self._flash_msg, 0, 12, 1)
            else:
                lay = _active_layer()
                name = lay.get("name","?")[:21] if lay else "\u2014"
                hw.text(name, 0, 0, 1)
                if lay:
                    m1 = ENC_SHORT.get(lay.get("enc1_mode",""), "?")
                    m2 = ENC_SHORT.get(lay.get("enc2_mode",""), "?")
                    hw.text("1:" + m1 + " 2:" + m2, 0, 10, 1)
                if enc2_zoom_override:
                    hw.text("ENC2:ZOOM", 0, 22, 1)
            hw.show()
        except Exception:
            pass   # missing font or I2C error — don't crash the firmware
        self._dirty     = False
        self._last_draw = now

oled = OLEDManager(oled_hw) if oled_ok else NoOpOLED()

# ─────────────────────────────────────────────────────────────
#  9. ACTION HELPERS
# ─────────────────────────────────────────────────────────────

def resolve(combo):
    return [KC[k] for k in combo if k in KC]

def press_combo(combo):
    c = resolve(combo)
    if c: kbd.press(*c)

def release_combo(combo):
    c = resolve(combo)
    if c: kbd.release(*c)

def send_combo(combo):
    c = resolve(combo)
    if c: kbd.send(*c)

# Pre-built action dispatch table — avoids elif chain on every button event
_CC = ConsumerControlCode
_ACTION_MAP = {
    "LEFT_CLICK":   lambda: mouse.click(Mouse.LEFT_BUTTON),
    "RIGHT_CLICK":  lambda: mouse.click(Mouse.RIGHT_BUTTON),
    "MIDDLE_CLICK": lambda: mouse.click(Mouse.MIDDLE_BUTTON),
    "MUTE":         lambda: cc.send(_CC.MUTE),
    "VOL_UP":       lambda: cc.send(_CC.VOLUME_INCREMENT),
    "VOL_DOWN":     lambda: cc.send(_CC.VOLUME_DECREMENT),
    "PLAY_PAUSE":   lambda: cc.send(_CC.PLAY_PAUSE),
}

def execute_action(name):
    if not name or name == "\u2014": return
    fn = _ACTION_MAP.get(name)
    if fn:
        fn()
    elif name in KC:
        kbd.send(KC[name])

def play_macro(steps):
    for step in steps:
        send_combo(step.get("combo", []))
        d = step.get("delay_ms", 0)
        if d > 0: time.sleep(d / 1000)

# ─────────────────────────────────────────────────────────────
#  10. LAYER MANAGEMENT  (base + MO stack)
# ─────────────────────────────────────────────────────────────

_mo_stack = []

# Pre-built strings — avoids runtime string construction on every event
_ENC_SW_KEYS = ("enc1_sw", "enc2_sw")   # indexed by ese.key_number (0 or 1)

# Pre-built key event JSON — avoids json.dumps on every keypress/release
# Keys 0-24 cover the full 5x5 matrix
_KP_JSON = ['{"event":"key_press","index":' + str(i) + '}\n' for i in range(25)]
_KR_JSON = ['{"event":"key_release","index":' + str(i) + '}\n' for i in range(25)]

def _active_layer_idx():
    layers = cfg["layers"]
    n = len(layers)
    if not n: return 0
    for i in range(len(_mo_stack) - 1, -1, -1):
        v = _mo_stack[i]
        if 0 <= v < n:
            return v
    al = cfg["active_layer"]
    return al if al < n else n - 1

def _active_layer():
    layers = cfg["layers"]
    if not layers: return None
    return layers[_active_layer_idx()]

def _broadcast_layer():
    idx = _active_layer_idx()
    layers = cfg["layers"]
    lay = layers[idx] if layers else None
    SC.sm_active = bool(lay.get("sm_active", False)) if lay else False
    send_json({
        "event": "layer_changed",
        "index": idx,
        "name":  lay["name"] if lay else "",
        "mo":    bool(_mo_stack),
        "sm_active": SC.sm_active,
    })
    oled.mark_dirty()

def cycle_layer():
    layers = cfg["layers"]
    n = len(layers)
    if n < 2: return
    cfg["active_layer"] = (cfg["active_layer"] + 1) % n
    _broadcast_layer()

# ─────────────────────────────────────────────────────────────
#  11. ENCODER MODIFIER TRACKING
# ─────────────────────────────────────────────────────────────

_enc_mod_keys = {}

def _enc_factor():
    return min(_enc_mod_keys.values()) if _enc_mod_keys else 1.0

# ─────────────────────────────────────────────────────────────
#  12. TAP / HOLD + KEY REPEAT
#  State stored as lists [value, value] instead of dicts to
#  reduce allocation overhead on every keypress.
#    tap_hold_state[idx] = [pressed_at, hold_fired]
#    repeat_state[idx]   = [started_at, last_repeat, combo]
# ─────────────────────────────────────────────────────────────

tap_hold_state = {}
repeat_state   = {}
press_layer    = {}   # idx -> layer at press time (for correct release)

def _wants_repeat(key_cfg, combo):
    per_key = key_cfg["key_repeat"]
    if per_key is not None: return per_key
    if not SC.key_repeat_enabled: return False
    return any(k in REPEAT_KEYS for k in combo)

def key_press(idx, now):
    layer = _active_layer()
    if not layer: return
    kc    = layer["keys"][idx]
    kt    = kc["key_type"]
    tap   = kc["tap"]
    hold  = kc["hold"]
    th_on = kc["tap_hold_enabled"]
    macro = kc["macro"]

    press_layer[idx] = layer

    if kt == KT_MO:
        mo_idx = kc["mo_layer"] or 0
        layers = cfg["layers"]
        if 0 <= mo_idx < len(layers):
            _mo_stack.append(mo_idx)
            _broadcast_layer()
        _stdout_write(_KP_JSON[idx])
        return

    if kt == KT_MOUSE_HOLD:
        if not passthrough_mode:
            mouse.press(MOUSE_BUTTONS.get(kc["mouse_button"] or "LEFT", Mouse.LEFT_BUTTON))
        _stdout_write(_KP_JSON[idx])
        return

    if kt == KT_ENC_MOD:
        _enc_mod_keys[idx] = kc["enc_mod_factor"] or 0.1
        _stdout_write(_KP_JSON[idx])
        return

    if macro:
        if not passthrough_mode:
            play_macro(macro)
            oled.flash("+".join(tap) if tap else "MACRO")
        _stdout_write(_KP_JSON[idx])
        return

    # Normal / tap-hold
    if th_on and hold:
        tap_hold_state[idx] = [now, False]
    else:
        if tap and not passthrough_mode:
            press_combo(tap)
        if _wants_repeat(kc, tap):
            repeat_state[idx] = [now, now, tap]
        if tap and not passthrough_mode:
            oled.flash("+".join(tap))
    _stdout_write(_KP_JSON[idx])

def key_release(idx, now):
    layer = press_layer.pop(idx, None) or _active_layer()
    if not layer: return
    kc = layer["keys"][idx]
    kt = kc["key_type"]

    if kt == KT_MO:
        if _mo_stack: _mo_stack.pop()
        _broadcast_layer()
        _stdout_write(_KR_JSON[idx])
        return

    if kt == KT_MOUSE_HOLD:
        if not passthrough_mode:
            mouse.release(MOUSE_BUTTONS.get(kc["mouse_button"] or "LEFT", Mouse.LEFT_BUTTON))
        _stdout_write(_KR_JSON[idx])
        return

    if kt == KT_ENC_MOD:
        _enc_mod_keys.pop(idx, None)
        _stdout_write(_KR_JSON[idx])
        return

    repeat_state.pop(idx, None)
    tap  = kc["tap"]
    hold = kc["hold"]
    th   = kc["tap_hold_enabled"]

    state = tap_hold_state.pop(idx, None)
    if state is not None:
        if not passthrough_mode:
            if state[1]:
                release_combo(hold)
            elif (now - state[0]) < SC.tap_hold_s:
                send_combo(tap)
                if tap: oled.flash("+".join(tap))
        _stdout_write(_KR_JSON[idx])
        return

    if tap and not th and not passthrough_mode:
        release_combo(tap)
    _stdout_write(_KR_JSON[idx])

def poll_tap_hold(now):
    if not tap_hold_state and not repeat_state:
        return   # fast path — nothing held, nothing repeating
    thresh = SC.tap_hold_s
    delay  = SC.key_repeat_delay_s
    rate   = SC.key_repeat_rate_s
    layer  = _active_layer()
    if not layer: return

    for idx, state in tap_hold_state.items():
        if not state[1] and now - state[0] >= thresh:
            hold = layer["keys"][idx]["hold"]
            if hold: press_combo(hold)
            state[1] = True

    for idx, rs in list(repeat_state.items()):
        if now - rs[0] < delay: continue
        if now - rs[1] >= rate:
            send_combo(rs[2])
            rs[1] = now

# ─────────────────────────────────────────────────────────────
#  13. ENCODER HANDLER
# ─────────────────────────────────────────────────────────────

def handle_encoder(enc_num, delta):
    if passthrough_mode: return   # suppress HID in passthrough/visualiser mode
    layer = _active_layer()
    if enc_num == 1:
        mode = layer["enc1_mode"] if layer else "V_SCROLL"
        spd  = int(SC.enc1_speed * _enc_factor())
        if SC.enc1_invert: delta = -delta
    else:
        mode = layer["enc2_mode"] if layer else "V_SCROLL"
        spd  = int(SC.enc2_speed * _enc_factor())
        if SC.enc2_invert: delta = -delta

    if mode == "H_SCROLL":
        kbd.press(Keycode.SHIFT)
        mouse.move(wheel=delta * spd)
        kbd.release(Keycode.SHIFT)
    elif mode == "V_SCROLL":
        mouse.move(wheel=-delta * spd)
    elif mode == "ZOOM":
        kbd.press(Keycode.CONTROL)
        mouse.move(wheel=delta * spd)
        kbd.release(Keycode.CONTROL)
    elif mode == "UNDO_REDO":
        if delta > 0:
            for _ in range(delta):  kbd.send(Keycode.GUI, Keycode.Z)
        else:
            for _ in range(-delta): kbd.send(Keycode.GUI, Keycode.SHIFT, Keycode.Z)
    elif mode == "NEXT_PREV_TAB":
        if delta > 0:
            for _ in range(delta):  kbd.send(Keycode.CONTROL, Keycode.TAB)
        else:
            for _ in range(-delta): kbd.send(Keycode.CONTROL, Keycode.SHIFT, Keycode.TAB)
    elif mode == "VOLUME":
        for _ in range(abs(delta)):
            if delta > 0: cc.send(ConsumerControlCode.VOLUME_INCREMENT)
            else:         cc.send(ConsumerControlCode.VOLUME_DECREMENT)

# ─────────────────────────────────────────────────────────────
#  14. SERIAL PROTOCOL
# ─────────────────────────────────────────────────────────────

serial_buf        = ""
telemetry_active  = False
passthrough_mode  = False  # when True: key events reported but HID suppressed (visualiser)

def handle_command(raw):
    global telemetry_active, passthrough_mode
    try:
        cmd = json.loads(raw)
    except Exception:
        send_json({"error":"invalid_json"}); return

    action = cmd.get("action","")

    if action == "get_config":
        send_json({"event":"config","data":cfg})

    elif action == "subscribe":
        telemetry_active = True
        send_json({"event":"subscribed"})

    elif action == "unsubscribe":
        telemetry_active = False
        send_json({"event":"unsubscribed"})

    elif action == "passthrough_on":
        passthrough_mode = True
        send_json({"event":"passthrough_on"})

    elif action == "passthrough_off":
        passthrough_mode = False
        # Release any held keys in case they were stuck
        kbd.release_all()
        mouse.release_all()
        send_json({"event":"passthrough_off"})

    elif action == "set":
        k, v = cmd.get("key"), cmd.get("value")
        if k and k in cfg and k not in ("layers","active_layer"):
            cfg[k] = v
            _sync_cache()
            send_json({"event":"ack","key":k,"value":v})
            oled.mark_dirty()
        else:
            send_json({"error":"unknown_key","key":k})

    elif action == "set_layers":
        layers = cmd.get("layers")
        if isinstance(layers, list) and len(layers) >= 1:
            cfg["layers"] = layers
            cfg["active_layer"] = min(cfg["active_layer"], len(layers)-1)
            send_json({"event":"ack_layers","count":len(layers)})
            oled.mark_dirty()
        else:
            send_json({"error":"invalid_layers"})

    elif action == "add_layer":
        name = cmd.get("name", "Layer " + str(len(cfg["layers"])+1))
        src  = cmd.get("copy_from")   # optional: index of layer to copy
        layers = cfg["layers"]
        if src is not None and 0 <= src < len(layers):
            new_layer = _deep_copy_cfg(layers[src])
            new_layer["name"] = name
        else:
            new_layer = _blank_layer(name)
        # Allow overriding specific fields (for templates)
        for k in ("keys","sm_orbit_mods","sm_pan_mods","sm_active",
                   "enc1_mode","enc2_mode","enc1_sw","enc2_sw"):
            if k in cmd:
                new_layer[k] = cmd[k]
        layers.append(new_layer)
        send_json({"event":"ack_add_layer","index":len(layers)-1,"name":name})

    elif action == "remove_layer":
        idx = cmd.get("index",-1)
        layers = cfg["layers"]
        if len(layers) > 1 and 0 <= idx < len(layers):
            layers.pop(idx)
            cfg["active_layer"] = min(cfg["active_layer"], len(layers)-1)
            send_json({"event":"ack_remove_layer"})
        else:
            send_json({"error":"cannot_remove_last_layer"})

    elif action == "rename_layer":
        idx = cmd.get("index",0); name = cmd.get("name","")
        layers = cfg["layers"]
        if 0 <= idx < len(layers) and name:
            layers[idx]["name"] = name
            send_json({"event":"ack_rename_layer"})
            oled.mark_dirty()

    elif action == "set_active_layer":
        idx = cmd.get("index",0)
        if 0 <= idx < len(cfg["layers"]):
            cfg["active_layer"] = idx
            _broadcast_layer()

    elif action == "set_key":
        li = cmd.get("layer", cfg["active_layer"])
        ki = cmd.get("index")
        layers = cfg["layers"]
        if ki is not None and 0 <= li < len(layers) and 0 <= ki < 25:
            kc = layers[li]["keys"][ki]
            for f in ("tap","hold","tap_hold_enabled","macro","key_repeat",
                      "key_type","mo_layer","mouse_button","enc_mod_factor"):
                if f in cmd: kc[f] = cmd[f]
            send_json({"event":"ack_key","layer":li,"index":ki})
        else:
            send_json({"error":"invalid_key_index"})

    elif action == "set_layer_prop":
        li   = cmd.get("layer", cfg["active_layer"])
        key  = cmd.get("key")
        val  = cmd.get("value")
        layers = cfg["layers"]
        if key and 0 <= li < len(layers):
            layers[li][key] = val
            # Refresh sm_active cache if current layer changed
            if key == "sm_active" and li == _active_layer_idx():
                SC.sm_active = bool(val)
            send_json({"event":"ack_layer_prop","layer":li,"key":key})
            oled.mark_dirty()
        else:
            send_json({"error":"invalid_layer_prop"})

    elif action == "set_encoder_mode":
        li = cmd.get("layer", cfg["active_layer"])
        enc = cmd.get("enc"); mode = cmd.get("mode")
        layers = cfg["layers"]
        if enc in (1,2) and mode in ENCODER_MODES and 0 <= li < len(layers):
            layers[li]["enc" + str(enc) + "_mode"] = mode
            send_json({"event":"ack_encoder_mode"})
            oled.mark_dirty()
        else:
            send_json({"error":"invalid_encoder_mode"})

    elif action == "set_enc_sw":
        li = cmd.get("layer", cfg["active_layer"])
        enc = cmd.get("enc"); val = cmd.get("value","")
        layers = cfg["layers"]
        if enc in (1,2) and 0 <= li < len(layers):
            layers[li]["enc" + str(enc) + "_sw"] = val
            send_json({"event":"ack_enc_sw"})
        else:
            send_json({"error":"invalid_enc_sw"})

    elif action == "zero":
        if mlx:
            try:
                offsets = sm.recalibrate()
                send_json({"event":"zeroed","offsets":offsets})
                oled.flash("Space Mouse Zeroed")
            except Exception as e:
                send_json({"error":"zero_failed","detail":str(e)})
        else:
            send_json({"error":"no_mlx"})

    elif action == "save":
        ok = save_config(cfg)
        send_json({"event":"saved" if ok else "save_failed"})
        if ok: oled.flash("Saved!")

    elif action == "ping":
        diag = {"event":"pong","sm_active":SC.sm_active,"mlx":mlx is not None}
        if mlx:
            diag["mlx_ok"] = mlx._ok_count
            diag["mlx_err"] = mlx._err_count
        send_json(diag)

    else:
        send_json({"error":"unknown_action","action":action})

# ─────────────────────────────────────────────────────────────
#  15. BOOT
# ─────────────────────────────────────────────────────────────

# Populate the settings cache from the loaded config
_sync_cache()
# Initialise sm_active from the boot layer
_boot_layer = _active_layer()
SC.sm_active = bool(_boot_layer.get("sm_active", False)) if _boot_layer else False

send_json({"event":"boot_complete","health":{
    "mlx":      mlx_ok,
    "oled":     oled_ok,
    "settings": _settings_status,
    "enc1_pos": encoder1.position,
    "enc2_pos": encoder2.position,
    "layers":   len(cfg["layers"]),
}})

if oled_ok:
    oled.flash("Ready  (" + _settings_status + ")", ms=2000)

SENSOR_INTV    = 0.02    # hardware read rate — 50 Hz
TELEMETRY_INTV = 0.05   # GUI update rate  — 20 Hz (reduces serial GC pressure)
last_sensor    = 0.0
last_telemetry = 0.0

enc2_zoom_override = False
btn1_held          = False
btn1_down_at       = 0.0
ENC2_HOLD_S  = 0.300   # 300ms in seconds — consistent with other timers

prev_jx = prev_jy = 0.0   # float to match read_joy_x/y return type
jx = jy = 0.0              # initialised here so telemetry block is safe on first tick

# ── Joystick calibration ──────────────────────────────────────
# Wait for ADC to settle, then sample the resting position.
# This handles joysticks whose centre is not exactly 32768,
# and prevents spurious mouse movement at boot.
time.sleep(0.1)   # ADC settling time

_JOY_SAMPLES = 16
_jx_sum = _jy_sum = 0
for _ in range(_JOY_SAMPLES):
    _jx_sum += joy_x.value
    _jy_sum += joy_y.value
    time.sleep(0.005)

_JOY_X_CENTRE  = _jx_sum // _JOY_SAMPLES
_JOY_Y_CENTRE  = _jy_sum // _JOY_SAMPLES
# Divide by 500 instead of 1000 — gives values in range ~0–65 at full deflection
# before speed multiplier, providing much smoother movement
_JOY_INV_SCALE = 1.0 / 500.0

send_json({"event":"joy_calibrated",
           "cx": _JOY_X_CENTRE, "cy": _JOY_Y_CENTRE})

# Fractional accumulators — carry sub-pixel movement between ticks to eliminate stutter
_joy_accum_x = 0.0
_joy_accum_y = 0.0

def read_joy_x():
    v = joy_x.value - _JOY_X_CENTRE
    return 0.0 if abs(v) < SC.joy_deadzone else v * _JOY_INV_SCALE

def read_joy_y():
    v = joy_y.value - _JOY_Y_CENTRE
    return 0.0 if abs(v) < SC.joy_deadzone else v * _JOY_INV_SCALE

# ─────────────────────────────────────────────────────────────
#  16. MAIN LOOP
#  Local bindings for frequently-accessed globals — in CircuitPython
#  LOAD_FAST is ~40% faster than LOAD_GLOBAL on every tick.
# ─────────────────────────────────────────────────────────────

# Bind to locals once outside the loop
_monotonic      = time.monotonic
_serial_avail   = supervisor.runtime
_stdin_read     = sys.stdin.read
_keys_get       = keys.events.get
_direct_get     = direct_keys.events.get
_encsw_get      = encoder_switches.events.get
_enc1           = encoder1
_enc2           = encoder2

while True:
    now = _monotonic()

    # ── Serial ──────────────────────────────────────────────
    if _serial_avail.serial_bytes_available:
        serial_buf += _stdin_read(_serial_avail.serial_bytes_available)
        while "\n" in serial_buf:
            line, serial_buf = serial_buf.split("\n", 1)
            line = line.strip()
            if line: handle_command(line)

    # ── Tap/hold + key repeat ────────────────────────────────
    poll_tap_hold(now)

    # ── Matrix ──────────────────────────────────────────────
    ev = _keys_get()
    if ev:
        if ev.pressed: key_press(ev.key_number, now)
        else:          key_release(ev.key_number, now)

    # ── Direct buttons ──────────────────────────────────────
    de = _direct_get()
    if de:
        dkn = de.key_number
        if dkn == 0:
            if de.pressed:
                btn1_held = True
                btn1_down_at = now
            else:
                btn1_held = False
                held = now - btn1_down_at
                if enc2_zoom_override:
                    enc2_zoom_override = False
                    send_json({"event":"enc2_zoom_override","active":False})
                    oled.mark_dirty()
                elif held < ENC2_HOLD_S:
                    if not passthrough_mode:
                        execute_action(cfg["btn_extra1"])
        elif dkn == 1:
            if de.pressed: cycle_layer()
        elif dkn == 2:
            if de.pressed and not passthrough_mode:
                execute_action(cfg["joy_sw"])

    # Activate enc2 zoom override when btn1 held past threshold
    if btn1_held and not enc2_zoom_override:
        if now - btn1_down_at >= ENC2_HOLD_S:
            enc2_zoom_override = True
            send_json({"event":"enc2_zoom_override","active":True})
            oled.flash("ENC2: ZOOM")

    # ── Encoder switches ────────────────────────────────────
    ese = _encsw_get()
    if ese and ese.pressed:
        lay = _active_layer()
        if lay and not passthrough_mode:
            execute_action(lay[_ENC_SW_KEYS[ese.key_number]])

    # ── Encoders ────────────────────────────────────────────
    p1 = _enc1.position
    if p1 != last_pos1:
        handle_encoder(1, p1 - last_pos1)
        last_pos1 = p1

    p2 = _enc2.position
    if p2 != last_pos2:
        d2 = p2 - last_pos2
        if enc2_zoom_override:
            spd = int(SC.enc2_speed * _enc_factor())
            if SC.enc2_invert: d2 = -d2
            kbd.press(Keycode.CONTROL)
            mouse.move(wheel=d2 * spd)
            kbd.release(Keycode.CONTROL)
        else:
            handle_encoder(2, d2)
        last_pos2 = p2

    # ── MLX space mouse (non-blocking — runs every tick) ────────
    # Only active when the current layer has sm_active = True.
    # Starts a measurement and returns immediately; reads result
    # on a later tick when DRDY is set — main loop never blocks.
    if mlx and SC.sm_active:
        if mlx.tick(now):
            sm.update(mlx.x, mlx.y, mlx.z, now)
    elif sm.is_orbiting or sm.is_panning:
        # Layer switched away from CAD — release any held buttons
        sm.safety_release()

    # ── Sensors (50 Hz — joystick read + mouse move) ─────────────
    if now - last_sensor > SENSOR_INTV:
        jx = read_joy_x()
        jy = read_joy_y()
        if jx or jy:
            if not passthrough_mode:
                spd = SC.joy_speed
                if SC.joy_invert_x: jx = -jx
                if SC.joy_invert_y: jy = -jy
                _joy_accum_x += jx * spd
                _joy_accum_y += jy * spd
                mx = int(_joy_accum_x)
                my = int(_joy_accum_y)
                if mx or my:
                    mouse.move(x=mx, y=my)
                    _joy_accum_x -= mx
                    _joy_accum_y -= my
        else:
            _joy_accum_x = 0.0
            _joy_accum_y = 0.0

        last_sensor = now

    # ── Telemetry (20 Hz — serial sends to GUI only) ────────────
    if telemetry_active and now - last_telemetry > TELEMETRY_INTV:
        if jx != prev_jx or jy != prev_jy:
            _stdout_write(
                '{"event":"joy_data","x":' + str(round(jx, 1)) +
                ',"y":' + str(round(jy, 1)) + '}\n'
            )
            prev_jx = jx
            prev_jy = jy
        if mlx and SC.sm_active:   # only send sm_data when space mouse is active
            _stdout_write(
                '{"event":"sm_data","x":' + str(round(sm.fx, 1)) +
                ',"y":' + str(round(sm.fy, 1)) +
                ',"z":' + str(round(sm.fz, 1)) +
                ',"ok":' + str(mlx._ok_count) +
                ',"err":' + str(mlx._err_count) + '}\n'
            )
        last_telemetry = now

    # ── OLED (own I2C bus — no contention with MLX) ────────
    oled.update()
