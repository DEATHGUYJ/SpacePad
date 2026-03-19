#!/usr/bin/env python3
"""
SpacePad Tray App  —  spacepad_tray.py
Watches the active window and switches Pico layers automatically.
pip install pystray pillow pyserial pywin32
"""
import json
import time
import threading
import sys
import os
import serial
import serial.tools.list_ports

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_OK = True
except ImportError:
    TRAY_OK = False

try:
    import win32gui
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

# Profiles stored next to the exe/script — not relative to cwd
_HERE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
PROFILES_FILE = os.path.join(_HERE, "profiles.json")
DEFAULT_PROFILES = {"port": "", "enabled": True, "mappings": []}


# ─────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────

def load_profiles():
    try:
        with open(PROFILES_FILE) as f:
            return json.load(f)
    except Exception:
        return dict(DEFAULT_PROFILES)

def save_profiles(p):
    try:
        with open(PROFILES_FILE, "w") as f:
            json.dump(p, f, indent=2)
    except Exception as e:
        print(f"Save error: {e}")

def get_foreground_title():
    if not WIN32_OK:
        return ""
    try:
        return win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
    except Exception:
        return ""

def list_ports():
    return [p.device for p in serial.tools.list_ports.comports()]

def make_icon_image(color=(100, 100, 255)):
    img = Image.new("RGB", (64, 64), color=(20, 20, 20))
    d   = ImageDraw.Draw(img)
    for r in range(3):
        for c in range(3):
            x, y = 8 + c * 18, 8 + r * 18
            d.rectangle([x, y, x + 12, y + 12], fill=color)
    return img


# ─────────────────────────────────────────────────────────────
#  SERIAL CONNECTION
# ─────────────────────────────────────────────────────────────

class PicoLink:
    RECONNECT_INTERVAL = 5.0   # seconds between reconnect attempts

    def __init__(self):
        self.port  = None
        self._lock = threading.Lock()
        self._last_connect_attempt = 0.0

    def connect(self, port_name):
        self.disconnect()
        try:
            p = serial.Serial(port_name, 115200, timeout=0.1)
            with self._lock:
                self.port = p
            return True
        except Exception:
            self.port = None
            return False

    def disconnect(self):
        with self._lock:
            p = self.port
            self.port = None
        if p and p.is_open:
            try: p.close()
            except Exception: pass

    def connected(self):
        with self._lock:
            return bool(self.port and self.port.is_open)

    def send_layer(self, idx):
        with self._lock:
            p = self.port
        if p and p.is_open:
            try:
                msg = json.dumps({"action": "set_active_layer", "index": idx}) + "\n"
                p.write(msg.encode())
            except Exception:
                with self._lock:
                    self.port = None

    def ensure_connected(self, port_name):
        """Try to reconnect with backoff — avoids hammering the port every 500ms."""
        if self.connected() or not port_name:
            return
        now = time.monotonic()
        if now - self._last_connect_attempt < self.RECONNECT_INTERVAL:
            return
        self._last_connect_attempt = now
        if port_name in list_ports():
            self.connect(port_name)


# ─────────────────────────────────────────────────────────────
#  PORT SELECTOR  (uses a small Tk window — no console needed)
# ─────────────────────────────────────────────────────────────

def _tk_port_selector(ports, current):
    """Show a simple Tk dialog for port selection. Returns selected port or None."""
    try:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        root.title("SpacePad — Select Port")
        root.resizable(False, False)
        root.attributes("-topmost", True)
        root.configure(bg="#0e0e1a")

        tk.Label(root, text="Select Pico serial port:",
                 bg="#0e0e1a", fg="#e2e8f0",
                 font=("Segoe UI", 11)).pack(padx=20, pady=(16, 6))

        var = tk.StringVar(value=current if current in ports else (ports[0] if ports else ""))
        cb = ttk.Combobox(root, textvariable=var, values=ports, state="readonly", width=28)
        cb.pack(padx=20, pady=6)

        result = [None]

        def ok():
            result[0] = var.get()
            root.destroy()

        def cancel():
            root.destroy()

        btn_frame = tk.Frame(root, bg="#0e0e1a")
        btn_frame.pack(pady=(6, 16))
        tk.Button(btn_frame, text="Connect", command=ok,
                  bg="#3b82f6", fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=5).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel", command=cancel,
                  bg="#1e1e30", fg="#94a3b8",
                  font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=5).pack(side="left", padx=6)

        root.update_idletasks()
        w, h = root.winfo_width(), root.winfo_height()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
        root.mainloop()
        return result[0]
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
#  TRAY APPLICATION
# ─────────────────────────────────────────────────────────────

class TrayApp:
    def __init__(self):
        self.profiles      = load_profiles()
        self.link          = PicoLink()
        self.current_layer = -1
        self._running      = True
        self._icon         = None
        self.link.connect(self.profiles.get("port", "")) if self.profiles.get("port") else None

    # ── Background watcher ────────────────────

    def _watch_loop(self):
        while self._running:
            try:
                self.link.ensure_connected(self.profiles.get("port", ""))
                if self.profiles.get("enabled", True):
                    title    = get_foreground_title()
                    mappings = self.profiles.get("mappings", [])
                    matched  = False
                    for m in mappings:
                        kw  = m.get("app", "").lower()
                        idx = m.get("layer", 0)
                        if kw and kw in title:
                            if idx != self.current_layer:
                                self.link.send_layer(idx)
                                self.current_layer = idx
                                self._update_icon(active=True)
                            matched = True
                            break
                    if not matched:
                        # Reset current_layer so re-entering a matched app
                        # always re-sends the layer switch
                        if self.current_layer != -1:
                            self.current_layer = -1
                            self._update_icon(active=False)
            except Exception:
                pass
            time.sleep(0.5)

    # ── Icon ──────────────────────────────────

    def _update_icon(self, active=False):
        if self._icon:
            color = (100, 220, 100) if active else (100, 100, 255)
            self._icon.icon = make_icon_image(color)

    # ── Menu actions ──────────────────────────

    def _show_port_selector(self, icon, item):
        ports = list_ports()
        if not ports:
            return
        current = self.profiles.get("port", "")
        selected = _tk_port_selector(ports, current)
        if selected:
            self.profiles["port"] = selected
            save_profiles(self.profiles)
            self.link.connect(selected)

    def _toggle_enabled(self, icon, item):
        self.profiles["enabled"] = not self.profiles.get("enabled", True)
        save_profiles(self.profiles)

    def _reload_profiles(self, icon, item):
        self.profiles = load_profiles()
        self.current_layer = -1

    def _quit(self, icon, item):
        self._running = False
        icon.stop()

    # ── Run ───────────────────────────────────

    def run(self):
        if not TRAY_OK:
            sys.exit(1)

        threading.Thread(target=self._watch_loop, daemon=True).start()

        menu = pystray.Menu(
            pystray.MenuItem("SpacePad Tray", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Auto-switching enabled",
                self._toggle_enabled,
                checked=lambda item: self.profiles.get("enabled", True),
            ),
            pystray.MenuItem("Select port",     self._show_port_selector),
            pystray.MenuItem("Reload profiles", self._reload_profiles),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit),
        )

        self._icon = pystray.Icon(
            "spacepad",
            make_icon_image(),
            "SpacePad",
            menu,
        )
        self._icon.run()


# ─────────────────────────────────────────────────────────────
#  ENTRY
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    TrayApp().run()
