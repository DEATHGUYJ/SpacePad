"""
SpacePad Configurator  —  spacepad_gui.py
PySide6 GUI with dark blue theme.
pip install PySide6 pyserial
"""
import sys
import json
import time
import copy

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog, QVBoxLayout, QHBoxLayout,
    QGridLayout, QStackedWidget, QTabWidget, QScrollArea, QFrame, QSplitter,
    QLabel, QPushButton, QSlider, QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox,
    QRadioButton, QButtonGroup, QCheckBox, QListWidget, QListWidgetItem,
    QTextEdit, QSizePolicy, QFileDialog, QMessageBox, QInputDialog,
    QStatusBar, QToolButton, QSystemTrayIcon, QMenu,
)
from PySide6.QtCore import (
    Qt, QThread, Signal, Slot, QTimer, QSize, QPointF, QRectF, QObject,
    QSettings,
)
from PySide6.QtGui import (
    QPainter, QPen, QBrush, QColor, QFont, QPalette, QPixmap, QIcon,
    QLinearGradient, QPainterPath, QFontDatabase, QAction, QImage,
    QShortcut, QKeySequence,
)

try:
    import psutil
    PSUTIL_OK = True
except ImportError:
    PSUTIL_OK = False

import serial
import serial.tools.list_ports

# ─────────────────────────────────────────────────────────────
#  1. THEME
# ─────────────────────────────────────────────────────────────

class Theme:
    # Backgrounds
    BG       = "#08080f"
    SURFACE  = "#0e0e1a"
    SURFACE2 = "#161625"
    SURFACE3 = "#1e1e30"
    BORDER   = "#252540"
    BORDER2  = "#2e2e50"

    # Blue accent family
    BLUE     = "#3b82f6"   # primary
    BLUE_LT  = "#60a5fa"   # hover / highlight
    BLUE_DK  = "#1d4ed8"   # pressed
    BLUE_DIM = "#1e3a5f"   # subtle backgrounds
    BLUE_MID = "#2563eb"   # secondary accent

    # Semantic colours
    TEXT     = "#e2e8f0"
    TEXT_DIM = "#64748b"
    TEXT_MID = "#94a3b8"
    GREEN    = "#22c55e"
    GREEN_DIM= "#14532d"
    RED      = "#ef4444"
    RED_DIM  = "#450a0a"
    YELLOW   = "#eab308"
    ORANGE   = "#f97316"
    CYAN     = "#06b6d4"
    PURPLE   = "#a855f7"

    # Key type colours
    KT_NORMAL    = SURFACE3
    KT_MO        = "#1a3a4a"
    KT_MOUSE     = "#2a1a3a"
    KT_ENC       = "#2a2a10"
    KT_MACRO     = "#2a1a10"

T = Theme  # shorthand

QSS = f"""
/* ── Global ─────────────────────────────────────────────── */
* {{
    font-family: "Segoe UI", "SF Pro Display", "Inter", sans-serif;
    font-size: 12px;
    color: {T.TEXT};
    outline: none;
}}

QMainWindow, QDialog {{
    background-color: {T.BG};
}}

QWidget {{
    background-color: transparent;
}}

/* ── Scrollbars ──────────────────────────────────────────── */
QScrollBar:vertical {{
    background: {T.SURFACE};
    width: 6px;
    border-radius: 3px;
}}
QScrollBar::handle:vertical {{
    background: {T.BORDER2};
    border-radius: 3px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
    background: {T.BLUE_DIM};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar:horizontal {{
    background: {T.SURFACE};
    height: 6px;
    border-radius: 3px;
}}
QScrollBar::handle:horizontal {{
    background: {T.BORDER2};
    border-radius: 3px;
    min-width: 20px;
}}

/* ── Tabs ────────────────────────────────────────────────── */
QTabWidget::pane {{
    background: {T.SURFACE};
    border: 1px solid {T.BORDER};
    border-top: none;
    border-radius: 0 0 8px 8px;
}}
QTabBar::tab {{
    background: {T.SURFACE2};
    color: {T.TEXT_DIM};
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.5px;
    padding: 8px 18px;
    border: 1px solid {T.BORDER};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    margin-right: 2px;
}}
QTabBar::tab:selected {{
    background: {T.SURFACE};
    color: {T.BLUE_LT};
    border-bottom: 2px solid {T.BLUE};
}}
QTabBar::tab:hover:!selected {{
    background: {T.SURFACE3};
    color: {T.TEXT_MID};
}}

/* ── Buttons ─────────────────────────────────────────────── */
QPushButton {{
    background: {T.SURFACE3};
    color: {T.TEXT};
    border: 1px solid {T.BORDER2};
    border-radius: 6px;
    padding: 6px 16px;
    font-weight: 600;
    font-size: 11px;
}}
QPushButton:hover {{
    background: {T.BORDER2};
    border-color: {T.BLUE_DIM};
}}
QPushButton:pressed {{
    background: {T.BLUE_DIM};
    border-color: {T.BLUE_DK};
}}
QPushButton[accent="true"] {{
    background: {T.BLUE};
    color: white;
    border: none;
}}
QPushButton[accent="true"]:hover {{
    background: {T.BLUE_LT};
}}
QPushButton[accent="true"]:pressed {{
    background: {T.BLUE_DK};
}}
QPushButton[danger="true"] {{
    background: {T.RED_DIM};
    color: {T.RED};
    border: 1px solid {T.RED_DIM};
}}
QPushButton[danger="true"]:hover {{
    background: {T.RED};
    color: white;
}}
QPushButton[success="true"] {{
    background: {T.GREEN_DIM};
    color: {T.GREEN};
    border: 1px solid {T.GREEN_DIM};
}}
QPushButton[success="true"]:hover {{
    background: {T.GREEN};
    color: white;
}}

/* ── ComboBox ────────────────────────────────────────────── */
QComboBox {{
    background: {T.SURFACE3};
    color: {T.TEXT};
    border: 1px solid {T.BORDER2};
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 11px;
}}
QComboBox:hover {{
    border-color: {T.BLUE_DIM};
}}
QComboBox:focus {{
    border-color: {T.BLUE};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox::down-arrow {{
    image: none;
    width: 0; height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid {T.TEXT_DIM};
    margin-right: 8px;
}}
QComboBox QAbstractItemView {{
    background: {T.SURFACE2};
    color: {T.TEXT};
    border: 1px solid {T.BORDER};
    border-radius: 6px;
    selection-background-color: {T.BLUE_DIM};
    selection-color: {T.BLUE_LT};
    padding: 4px;
}}

/* ── Sliders ─────────────────────────────────────────────── */
QSlider::groove:horizontal {{
    background: {T.SURFACE3};
    height: 4px;
    border-radius: 2px;
}}
QSlider::handle:horizontal {{
    background: {T.BLUE};
    width: 16px;
    height: 16px;
    border-radius: 8px;
    margin: -6px 0;
}}
QSlider::handle:horizontal:hover {{
    background: {T.BLUE_LT};
}}
QSlider::sub-page:horizontal {{
    background: {T.BLUE};
    height: 4px;
    border-radius: 2px;
}}

/* ── LineEdit / SpinBox ──────────────────────────────────── */
QLineEdit, QSpinBox, QDoubleSpinBox {{
    background: {T.SURFACE3};
    color: {T.TEXT};
    border: 1px solid {T.BORDER2};
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 11px;
    selection-background-color: {T.BLUE_DIM};
}}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
    border-color: {T.BLUE};
}}
QSpinBox::up-button, QDoubleSpinBox::up-button {{
    background: {T.BORDER};
    border-radius: 0 6px 0 0;
    width: 18px;
}}
QSpinBox::down-button, QDoubleSpinBox::down-button {{
    background: {T.BORDER};
    border-radius: 0 0 6px 0;
    width: 18px;
}}

/* ── List ────────────────────────────────────────────────── */
QListWidget {{
    background: {T.SURFACE2};
    color: {T.TEXT};
    border: 1px solid {T.BORDER};
    border-radius: 8px;
    padding: 4px;
    outline: none;
}}
QListWidget::item {{
    padding: 7px 12px;
    border-radius: 5px;
    color: {T.TEXT_MID};
    font-size: 11px;
    font-family: "Consolas", monospace;
}}
QListWidget::item:selected {{
    background: {T.BLUE_DIM};
    color: {T.BLUE_LT};
}}
QListWidget::item:hover:!selected {{
    background: {T.SURFACE3};
}}

/* ── RadioButton / CheckBox ──────────────────────────────── */
QRadioButton, QCheckBox {{
    color: {T.TEXT_MID};
    font-size: 11px;
    spacing: 8px;
}}
QRadioButton:checked, QCheckBox:checked {{
    color: {T.TEXT};
}}
QRadioButton::indicator, QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border-radius: 8px;
    border: 2px solid {T.BORDER2};
    background: {T.SURFACE3};
}}
QCheckBox::indicator {{
    border-radius: 4px;
}}
QRadioButton::indicator:checked, QCheckBox::indicator:checked {{
    background: {T.BLUE};
    border-color: {T.BLUE};
}}

/* ── Frame / GroupBox ────────────────────────────────────── */
QFrame[card="true"] {{
    background: {T.SURFACE};
    border: 1px solid {T.BORDER};
    border-radius: 10px;
}}

/* ── Labels ──────────────────────────────────────────────── */
QLabel[heading="true"] {{
    color: {T.TEXT};
    font-size: 13px;
    font-weight: 700;
}}
QLabel[section="true"] {{
    color: {T.BLUE_LT};
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
    padding: 4px 0;
}}
QLabel[dim="true"] {{
    color: {T.TEXT_DIM};
    font-size: 11px;
}}

/* ── TextEdit ────────────────────────────────────────────── */
QTextEdit {{
    background: {T.SURFACE2};
    color: {T.TEXT};
    border: 1px solid {T.BORDER};
    border-radius: 6px;
    padding: 8px;
    font-family: "Consolas", monospace;
    font-size: 11px;
    selection-background-color: {T.BLUE_DIM};
}}

/* ── Status bar ──────────────────────────────────────────── */
QStatusBar {{
    background: {T.SURFACE};
    color: {T.TEXT_DIM};
    border-top: 1px solid {T.BORDER};
    font-size: 11px;
    font-family: "Consolas", monospace;
    padding: 2px 8px;
}}
"""

# ─────────────────────────────────────────────────────────────
#  2. HELPERS — CARD, SECTION LABEL, SEPARATOR
# ─────────────────────────────────────────────────────────────

def card_widget(parent=None):
    """Returns a QFrame styled as a card."""
    f = QFrame(parent)
    f.setProperty("card", True)
    f.setStyle(f.style())  # force style refresh
    return f

def make_card(title=None, parent=None):
    """Card with optional section title. Returns (card, inner_layout)."""
    frame = QFrame(parent)
    frame.setProperty("card", True)
    frame.setStyleSheet(f"""
        QFrame[card="true"] {{
            background: {T.SURFACE};
            border: 1px solid {T.BORDER};
            border-radius: 10px;
        }}
    """)
    vbox = QVBoxLayout(frame)
    vbox.setContentsMargins(16, 14, 16, 14)
    vbox.setSpacing(8)
    if title:
        lbl = QLabel(title.upper())
        lbl.setStyleSheet(f"""
            color: {T.TEXT_DIM};
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 1.5px;
            padding-bottom: 6px;
            font-family: 'Consolas', monospace;
        """)
        vbox.addWidget(lbl)
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"background: {T.BORDER}; max-height: 1px; border: none;")
        vbox.addWidget(sep)
    return frame, vbox

def section_label(text):
    lbl = QLabel(text.upper())
    lbl.setStyleSheet(f"""
        color: {T.TEXT_DIM};
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1.5px;
        padding: 6px 0 4px 0;
        font-family: 'Consolas', monospace;
    """)
    return lbl

def hsep():
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setStyleSheet(f"background: {T.BORDER}; max-height: 1px; border: none;")
    return line

def accent_btn(text, parent=None):
    b = QPushButton(text, parent)
    b.setProperty("accent", True)
    b.setStyleSheet(b.styleSheet())
    return b

def danger_btn(text, parent=None):
    b = QPushButton(text, parent)
    b.setProperty("danger", True)
    b.setStyleSheet(b.styleSheet())
    return b

def success_btn(text, parent=None):
    b = QPushButton(text, parent)
    b.setProperty("success", True)
    b.setStyleSheet(b.styleSheet())
    return b

# ─────────────────────────────────────────────────────────────
#  3. CUSTOM WIDGETS
# ─────────────────────────────────────────────────────────────

class ToggleSwitch(QWidget):
    """Animated ON/OFF toggle."""
    toggled = Signal(bool)

    def __init__(self, initial=False, parent=None):
        super().__init__(parent)
        self._on = initial
        self.setFixedSize(44, 24)
        self.setCursor(Qt.PointingHandCursor)

    def isChecked(self):
        return self._on

    def setChecked(self, val):
        self._on = bool(val)
        self.update()

    def mousePressEvent(self, e):
        self._on = not self._on
        self.update()
        self.toggled.emit(self._on)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        track_col = QColor(T.BLUE) if self._on else QColor(T.BORDER2)
        p.setBrush(track_col)
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, 4, 44, 16, 8, 8)
        knob_x = 24 if self._on else 4
        p.setBrush(QColor("white"))
        p.drawEllipse(knob_x, 2, 20, 20)
        p.end()


class SliderRow(QWidget):
    """Label + slider + value label row."""
    valueChanged = Signal(float)

    def __init__(self, label, key, min_val, max_val, step, initial=None, parent=None):
        super().__init__(parent)
        self.key = key
        self._updating = False
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(10)

        lbl = QLabel(label)
        lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        lbl.setFixedWidth(200)
        layout.addWidget(lbl)

        self.slider = QSlider(Qt.Horizontal)
        is_float = isinstance(step, float) or isinstance(min_val, float)
        if is_float:
            self._scale = 100
            self.slider.setRange(int(min_val * self._scale), int(max_val * self._scale))
            self.slider.setSingleStep(max(1, int(step * self._scale)))
            init = int((initial or min_val) * self._scale)
        else:
            self._scale = 1
            self.slider.setRange(int(min_val), int(max_val))
            self.slider.setSingleStep(int(step))
            init = int(initial or min_val)
        self.slider.setValue(init)
        self._is_float = is_float
        layout.addWidget(self.slider, 1)

        self.val_lbl = QLabel()
        self.val_lbl.setFixedWidth(48)
        self.val_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.val_lbl.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 11px; font-family: 'Consolas', monospace;"
        )
        layout.addWidget(self.val_lbl)

        self._refresh_label()
        self.slider.valueChanged.connect(self._on_change)

    def _refresh_label(self):
        v = self.get_value()
        self.val_lbl.setText(f"{v:.1f}" if self._is_float else str(int(v)))

    def _on_change(self):
        self._refresh_label()
        if not self._updating:
            self.valueChanged.emit(self.get_value())

    def get_value(self):
        raw = self.slider.value()
        return raw / self._scale if self._is_float else float(raw)

    def set_value(self, val):
        self._updating = True
        self.slider.setValue(int(float(val) * self._scale))
        self._refresh_label()
        self._updating = False


class BarMeter(QWidget):
    """Vertical bar meter with centre-zero and deadzone markers."""

    def __init__(self, color=T.BLUE, parent=None):
        super().__init__(parent)
        self._value = 0.0
        self._max   = 500.0
        self._dz    = 100.0
        self._color = QColor(color)
        self.setFixedSize(28, 110)

    def set_value(self, v):
        self._value = v
        self.update()

    def set_deadzone(self, dz):
        self._dz = dz
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        mid = h // 2

        # Track
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(T.SURFACE3))
        p.drawRoundedRect(0, 0, w, h, 4, 4)

        # Bar
        norm = max(-1.0, min(1.0, self._value / self._max))
        bar_h = int(abs(norm) * mid)
        p.setBrush(self._color)
        if norm >= 0:
            p.drawRect(0, mid - bar_h, w, bar_h)
        else:
            p.drawRect(0, mid, w, bar_h)

        # Deadzone markers
        dz_frac = min(self._dz / self._max, 1.0)
        dz_px = int(dz_frac * mid)
        p.setPen(QPen(QColor(T.YELLOW), 1, Qt.DashLine))
        p.drawLine(0, mid - dz_px, w, mid - dz_px)
        p.drawLine(0, mid + dz_px, w, mid + dz_px)

        # Centre line
        p.setPen(QPen(QColor(T.BORDER2), 1))
        p.drawLine(0, mid, w, mid)
        p.end()


class JoystickView(QWidget):
    """Live joystick position display."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._x = 0.0
        self._y = 0.0
        self._dz = 2000
        self.setFixedSize(130, 130)

    def set_position(self, x, y):
        self._x = x
        self._y = y
        self.update()

    def set_deadzone(self, dz):
        self._dz = dz
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        cx, cy = w // 2, h // 2
        r = min(cx, cy) - 4

        # Background
        p.setBrush(QColor(T.SURFACE3))
        p.setPen(Qt.NoPen)
        p.drawEllipse(cx - r, cy - r, r * 2, r * 2)

        # Deadzone ring
        dz_r = int(r * min(self._dz / 32768, 1.0))
        p.setPen(QPen(QColor(T.YELLOW), 1, Qt.DashLine))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(cx - dz_r, cy - dz_r, dz_r * 2, dz_r * 2)

        # Cross hairs
        p.setPen(QPen(QColor(T.BORDER2), 1))
        p.drawLine(cx, cy - r, cx, cy + r)
        p.drawLine(cx - r, cy, cx + r, cy)

        # Dot
        MAX_J = 32.0
        dx = int((self._x / MAX_J) * r * 0.9)
        dy = int((self._y / MAX_J) * r * 0.9)
        px = max(cx - r + 8, min(cx + r - 8, cx + dx))
        py = max(cy - r + 8, min(cy + r - 8, cy + dy))
        moving = abs(self._x) > 1 or abs(self._y) > 1
        dot_col = QColor(T.BLUE) if moving else QColor(T.GREEN)
        p.setPen(Qt.NoPen)
        p.setBrush(dot_col)
        p.drawEllipse(px - 7, py - 7, 14, 14)
        p.end()


class KeyButton(QPushButton):
    """A single key cell in the 5×5 matrix grid."""
    doubleClicked = Signal(int)

    def __init__(self, index, label="", key_type="normal", parent=None):
        super().__init__(parent)
        self.index = index
        self._pressed_anim = False
        self.setFixedSize(90, 62)
        self.setCursor(Qt.PointingHandCursor)
        self.set_key_data(label, key_type)

    def set_key_data(self, label, key_type="normal"):
        self._label = label
        self._type  = key_type
        colors = {
            "mo":         (T.BLUE_DIM,  T.CYAN),
            "mouse_hold": ("#2a1235",   T.PURPLE),
            "enc_mod":    ("#2a2505",   T.YELLOW),
            "macro":      ("#2a1400",   T.ORANGE),
            "normal":     (T.SURFACE3,  T.TEXT_MID),
        }
        bg, fg = colors.get(key_type, colors["normal"])
        if not label:
            fg = T.BORDER2
        self.setStyleSheet(f"""
            QPushButton {{
                background: {bg};
                color: {fg};
                border: 1px solid {T.BORDER};
                border-radius: 8px;
                font-size: 9px;
                font-weight: 600;
                font-family: 'Consolas', monospace;
                padding: 4px;
            }}
            QPushButton:hover {{
                border-color: {T.BLUE_DIM};
                background: {"#1a2a3a" if bg == T.SURFACE3 else bg};
            }}
            QPushButton:pressed {{
                border-color: {T.BLUE};
                background: {T.BLUE_DIM};
                color: white;
            }}
        """)
        display = label if label else "—"
        self.setText(display)

    def flash(self):
        """Called on live key press — applies blue highlight directly."""
        self.setStyleSheet(f"""
            QPushButton {{
                background: {T.BLUE};
                color: white;
                border: 1px solid {T.BLUE};
                border-radius: 8px;
                font-size: 9px;
                font-weight: 600;
                font-family: 'Consolas', monospace;
                padding: 4px;
            }}
        """)
        QTimer.singleShot(120, self._unflash)

    def _unflash(self):
        self.set_key_data(self._label, self._type)

    def mouseDoubleClickEvent(self, e):
        self.doubleClicked.emit(self.index)


# ─────────────────────────────────────────────────────────────
#  4. SERIAL WORKER  (QThread)
# ─────────────────────────────────────────────────────────────

class SerialWorker(QThread):
    message_received = Signal(dict)
    disconnected     = Signal()

    def __init__(self, port_name, baud=115200):
        super().__init__()
        self.port_name = port_name
        self.baud      = baud
        self._port     = None
        self._running  = False
        self._send_queue = []

    def run(self):
        import serial as _serial
        try:
            self._port = _serial.Serial(self.port_name, self.baud, timeout=0.05)
        except Exception:
            self.disconnected.emit()
            return
        self._running = True
        buf = ""
        while self._running:
            # Send queued messages
            while self._send_queue:
                msg = self._send_queue.pop(0)
                try:
                    self._port.write((json.dumps(msg) + "\n").encode())
                except Exception:
                    self._running = False
                    break

            # Read incoming — timeout returns empty bytes, not an exception
            try:
                raw = self._port.read(512)
                if raw:
                    buf += raw.decode("utf-8", errors="ignore")
                    while "\n" in buf:
                        line, buf = buf.split("\n", 1)
                        line = line.strip()
                        if line:
                            try:
                                self.message_received.emit(json.loads(line))
                            except Exception:
                                pass
            except Exception as e:
                # Only disconnect on real errors, not timeouts
                err = str(e).lower()
                if "timeout" not in err and "timed out" not in err:
                    self._running = False

        if self._port and self._port.is_open:
            try:
                self._port.close()
            except Exception:
                pass
        self.disconnected.emit()

    def send(self, obj):
        self._send_queue.append(obj)

    def stop(self):
        self._running = False

    def is_connected(self):
        return self._running and self._port and self._port.is_open


# ─────────────────────────────────────────────────────────────
#  5. SERIAL MANAGER
# ─────────────────────────────────────────────────────────────

class SerialManager(QObject):
    message_received = Signal(dict)
    connected        = Signal(str)
    disconnected     = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker      = None
        self._last_port   = ""
        self._watchdog    = QTimer(self)
        self._watchdog.timeout.connect(self._check_reconnect)
        self._watchdog.start(2000)

    def list_ports(self):
        return [p.device for p in serial.tools.list_ports.comports()]

    def connect_port(self, port_name):
        self.disconnect_port()
        self._worker = SerialWorker(port_name)
        self._worker.message_received.connect(self.message_received)
        self._worker.disconnected.connect(self._on_worker_disconnected)
        self._worker.start()
        self._last_port = port_name
        self.connected.emit(port_name)

    def disconnect_port(self):
        if self._worker:
            self._worker.stop()
            self._worker.wait(1000)
            self._worker = None

    def send(self, obj):
        if self._worker and self._worker.is_connected():
            self._worker.send(obj)

    def is_connected(self):
        return bool(self._worker and self._worker.is_connected())

    def _on_worker_disconnected(self):
        self._worker = None
        self.disconnected.emit()

    def _check_reconnect(self):
        if not self.is_connected() and self._last_port:
            if self._last_port in self.list_ports():
                self.connect_port(self._last_port)


# ─────────────────────────────────────────────────────────────
#  5b. APP WATCHER — foreground process monitor
# ─────────────────────────────────────────────────────────────

def _get_foreground_exe():
    """Return the lowercase exe name of the foreground window, or ''."""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return ""
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value and PSUTIL_OK:
            try:
                return psutil.Process(pid.value).name().lower()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return ""
        return ""
    except Exception:
        return ""


def _list_running_apps():
    """Return sorted list of unique exe names currently running with visible windows."""
    if not PSUTIL_OK:
        return []
    seen = set()
    result = []
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32

        # Callback collects PIDs of all visible, non-zero-size windows
        pids = set()
        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _cb(hwnd, _):
            if user32.IsWindowVisible(hwnd):
                pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value:
                    pids.add(pid.value)
            return True
        user32.EnumWindows(WNDENUMPROC(_cb), 0)

        for pid in pids:
            try:
                name = psutil.Process(pid).name().lower()
                if name not in seen and not name.startswith("svchost"):
                    seen.add(name)
                    result.append(name)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception:
        # Fallback: just list all user processes
        try:
            for p in psutil.process_iter(["name"]):
                name = (p.info["name"] or "").lower()
                if name and name not in seen:
                    seen.add(name)
                    result.append(name)
        except Exception:
            pass
    result.sort()
    return result


class AppWatcher(QThread):
    """Monitors the foreground application and emits layer switch requests."""
    layer_requested = Signal(int, str)    # (layer_index, exe_name)
    foreground_changed = Signal(str)      # exe_name — for live display
    no_match = Signal()                   # no mapping matched

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = False
        self._enabled = True
        self._mappings = []          # list of {"exe": str, "layer": int}
        self._default_layer = 0
        self._current_exe = ""
        self._current_layer = -1

    def set_mappings(self, mappings, default_layer=0):
        self._mappings = list(mappings) if mappings else []
        self._default_layer = default_layer

    def set_enabled(self, val):
        self._enabled = bool(val)

    def run(self):
        self._running = True
        while self._running:
            if self._enabled:
                exe = _get_foreground_exe()
                if exe and exe != self._current_exe:
                    self._current_exe = exe
                    self.foreground_changed.emit(exe)
                    matched = False
                    for m in self._mappings:
                        if m.get("exe", "").lower() == exe:
                            target = m.get("layer", 0)
                            if target != self._current_layer:
                                self._current_layer = target
                                self.layer_requested.emit(target, exe)
                            matched = True
                            break
                    if not matched:
                        if self._current_layer != self._default_layer:
                            self._current_layer = self._default_layer
                            self.layer_requested.emit(self._default_layer, "")
                        self.no_match.emit()
            self.msleep(500)

    def stop(self):
        self._running = False


# ─────────────────────────────────────────────────────────────
#  6. VALIDATE CONFIG
# ─────────────────────────────────────────────────────────────

def validate_config(data):
    if not isinstance(data, dict):
        return False, "Root must be a JSON object"
    layers = data.get("layers")
    if not isinstance(layers, list) or len(layers) < 1:
        return False, "'layers' must be a non-empty list"
    for i, layer in enumerate(layers):
        if not isinstance(layer, dict):
            return False, f"Layer {i} is not an object"
        keys = layer.get("keys")
        if not isinstance(keys, list) or len(keys) != 25:
            return False, f"Layer {i} 'keys' must be a list of 25 entries"
    return True, ""


# ─────────────────────────────────────────────────────────────
#  7. MACRO EDITOR DIALOG
# ─────────────────────────────────────────────────────────────

KEYSYM_MAP = {
    "Return":"ENTER", "Tab":"TAB", "Escape":"ESCAPE", "BackSpace":"BACKSPACE",
    "Delete":"DELETE", "Home":"HOME", "End":"END", "Prior":"PAGE_UP",
    "Next":"PAGE_DOWN", "Up":"UP", "Down":"DOWN", "Left":"LEFT", "Right":"RIGHT",
    "space":"SPACE", **{f"F{i}":f"F{i}" for i in range(1,13)},
}

MOD_MAP = {
    Qt.Key_Control: "CTRL",
    Qt.Key_Shift:   "SHIFT",
    Qt.Key_Alt:     "ALT",
    Qt.Key_Meta:    "GUI",
}

class MacroEditorDialog(QDialog):
    def __init__(self, steps=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Macro Editor")
        self.setMinimumSize(520, 500)
        self.setStyleSheet(f"background: {T.BG};")
        self.steps   = copy.deepcopy(steps or [])
        self._held   = set()
        self._result = None
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("MACRO EDITOR")
        title.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 16px; font-weight: 700;"
        )
        layout.addWidget(title)

        sub = QLabel("Add steps manually or use Record to capture live timing.")
        sub.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        layout.addWidget(sub)

        # Step list
        self.step_list = QListWidget()
        self.step_list.setMinimumHeight(160)
        layout.addWidget(self.step_list)
        self._refresh_list()

        # Capture area
        cap_card, cap_inner = make_card("CAPTURE")
        cap_label = QLabel("Click here and press a shortcut")
        cap_label.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 13px; font-weight: 600;"
            f" padding: 10px; background: {T.SURFACE3}; border-radius: 6px;"
        )
        cap_label.setAlignment(Qt.AlignCenter)
        cap_label.setFocusPolicy(Qt.ClickFocus)
        cap_label.mousePressEvent = lambda e: cap_label.setFocus()
        self._cap_label = cap_label
        cap_inner.addWidget(cap_label)
        layout.addWidget(cap_card)

        # Delay
        delay_row = QHBoxLayout()
        delay_row.addWidget(QLabel("Delay after step (ms):"))
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(0, 5000)
        self.delay_spin.setValue(50)
        self.delay_spin.setSingleStep(10)
        self.delay_spin.setFixedWidth(80)
        delay_row.addWidget(self.delay_spin)
        delay_row.addStretch()
        layout.addLayout(delay_row)

        # Record row
        rec_row = QHBoxLayout()
        self.rec_btn = QPushButton("⏺  RECORD")
        self.rec_btn.setStyleSheet(
            f"background: {T.SURFACE3}; color: {T.ORANGE}; border: 1px solid {T.ORANGE};"
            f" border-radius: 6px; padding: 5px 14px; font-weight: 700;"
        )
        self.rec_btn.clicked.connect(self._toggle_record)
        rec_row.addWidget(self.rec_btn)
        self._rec_label = QLabel("")
        self._rec_label.setStyleSheet(f"color: {T.YELLOW}; font-size: 11px;")
        rec_row.addWidget(self._rec_label)
        rec_row.addStretch()
        layout.addLayout(rec_row)
        self._recording = False
        self._rec_last  = 0.0

        # Buttons
        btn_row = QHBoxLayout()
        add_btn = accent_btn("ADD STEP")
        add_btn.clicked.connect(self._add_step)
        del_btn = danger_btn("DELETE")
        del_btn.clicked.connect(self._del_step)
        up_btn  = QPushButton("↑")
        up_btn.clicked.connect(self._move_up)
        dn_btn  = QPushButton("↓")
        dn_btn.clicked.connect(self._move_down)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(del_btn)
        btn_row.addWidget(up_btn)
        btn_row.addWidget(dn_btn)
        btn_row.addStretch()
        save_btn = success_btn("SAVE")
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton("CANCEL")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)
        layout.addLayout(btn_row)

        self._cap_label.installEventFilter(self)

    def eventFilter(self, obj, event):
        from PySide6.QtCore import QEvent
        if obj == self._cap_label:
            if event.type() == QEvent.KeyPress:
                key = event.key()
                if key in MOD_MAP:
                    self._held.add(MOD_MAP[key])
                else:
                    name = event.text().upper()
                    if not name or not name.isprintable():
                        name = KEYSYM_MAP.get(event.key(), None)
                    if name:
                        self._main_key = name
                return True
            if event.type() == QEvent.KeyRelease:
                key = event.key()
                if key not in MOD_MAP and hasattr(self, "_main_key") and self._main_key:
                    mods = sorted(self._held)
                    combo = mods + [self._main_key]
                    label = "+".join(combo)
                    self._cap_label.setText(label)
                    if self._recording:
                        now = time.time()
                        delay = int((now - self._rec_last) * 1000) if self._rec_last else 0
                        self._rec_last = now
                        self.steps.append({"combo": combo, "delay_ms": delay})
                        self._refresh_list()
                        self._rec_label.setText(f"  +{label}  (+{delay}ms)")
                    self._held.clear()
                    self._main_key = None
                return True
        return super().eventFilter(obj, event)

    def _toggle_record(self):
        self._recording = not self._recording
        if self._recording:
            self._rec_last = time.time()
            self.rec_btn.setText("⏹  STOP")
            self.rec_btn.setStyleSheet(
                f"background: {T.SURFACE3}; color: {T.RED}; border: 1px solid {T.RED};"
                f" border-radius: 6px; padding: 5px 14px; font-weight: 700;"
            )
            self._rec_label.setText("Recording… type your sequence")
            self._cap_label.setFocus()
        else:
            self.rec_btn.setText("⏺  RECORD")
            self.rec_btn.setStyleSheet(
                f"background: {T.SURFACE3}; color: {T.ORANGE}; border: 1px solid {T.ORANGE};"
                f" border-radius: 6px; padding: 5px 14px; font-weight: 700;"
            )
            self._rec_label.setText(f"Recorded {len(self.steps)} steps")
            self._refresh_list()

    def _add_step(self):
        text = self._cap_label.text()
        if not text or "Click here" in text:
            return
        combo = text.split("+")
        delay = self.delay_spin.value()
        self.steps.append({"combo": combo, "delay_ms": delay})
        self._refresh_list()
        self._cap_label.setText("Click here and press a shortcut")

    def _del_step(self):
        row = self.step_list.currentRow()
        if row >= 0:
            self.steps.pop(row)
            self._refresh_list()

    def _move_up(self):
        row = self.step_list.currentRow()
        if row > 0:
            self.steps[row-1], self.steps[row] = self.steps[row], self.steps[row-1]
            self._refresh_list()
            self.step_list.setCurrentRow(row - 1)

    def _move_down(self):
        row = self.step_list.currentRow()
        if row < len(self.steps) - 1:
            self.steps[row+1], self.steps[row] = self.steps[row], self.steps[row+1]
            self._refresh_list()
            self.step_list.setCurrentRow(row + 1)

    def _refresh_list(self):
        self.step_list.clear()
        for i, s in enumerate(self.steps):
            combo = "+".join(s.get("combo", []))
            delay = s.get("delay_ms", 0)
            item = QListWidgetItem(f"  {i+1:2}.  {combo:<26}  +{delay}ms")
            self.step_list.addItem(item)

    def _save(self):
        self._result = self.steps
        self.accept()

    def get_steps(self):
        return self._result


# ─────────────────────────────────────────────────────────────
#  8. KEY EDITOR DIALOG
# ─────────────────────────────────────────────────────────────

KEY_TYPE_LABELS = {
    "normal":     "Normal key",
    "mo":         "Momentary layer (MO)",
    "mouse_hold": "Hold mouse button",
    "enc_mod":    "Encoder speed modifier",
}

class KeyEditorDialog(QDialog):
    def __init__(self, key_idx, layer_idx, key_cfg, serial_mgr, parent=None):
        super().__init__(parent)
        r, c = divmod(key_idx, 5)
        self.setWindowTitle(f"Key  [{r},{c}]  —  Layer {layer_idx}")
        self.setMinimumSize(480, 580)
        self.setStyleSheet(f"background: {T.BG};")
        self.key_idx   = key_idx
        self.layer_idx = layer_idx
        self.serial    = serial_mgr
        self.cfg       = copy.deepcopy(key_cfg)
        self._held     = set()
        self._main_key = None
        self._slot     = None
        self._result   = None
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        r, c = divmod(self.key_idx, 5)
        title = QLabel(f"KEY  [{r},{c}]")
        title.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 16px; font-weight: 700;"
        )
        layout.addWidget(title)

        # Key type selector
        type_card, type_inner = make_card("KEY TYPE")
        self._type_group = QButtonGroup(self)
        type_grid = QGridLayout()
        type_grid.setSpacing(8)
        self._type_radios = {}
        for i, (kt, lbl) in enumerate(KEY_TYPE_LABELS.items()):
            rb = QRadioButton(lbl)
            rb.setChecked(self.cfg.get("key_type", "normal") == kt)
            self._type_group.addButton(rb)
            self._type_radios[kt] = rb
            type_grid.addWidget(rb, i // 2, i % 2)
        type_inner.addLayout(type_grid)
        layout.addWidget(type_card)

        # Dynamic config area (stacked)
        self._stack = QStackedWidget()
        self._stack.addWidget(self._build_normal_page())   # 0
        self._stack.addWidget(self._build_mo_page())       # 1
        self._stack.addWidget(self._build_mouse_page())    # 2
        self._stack.addWidget(self._build_enc_page())      # 3
        layout.addWidget(self._stack)

        # Switch stack on type change
        for kt, rb in self._type_radios.items():
            rb.toggled.connect(lambda checked, k=kt: self._on_type_change(k) if checked else None)
        self._on_type_change(self.cfg.get("key_type","normal"))

        # Per-key repeat
        rep_card, rep_inner = make_card("KEY REPEAT OVERRIDE")
        self._rep_group = QButtonGroup(self)
        rep_row = QHBoxLayout()
        for val, lbl in (("inherit","Use global"),("on","Always repeat"),("off","Never repeat")):
            rb = QRadioButton(lbl)
            cur = self.cfg.get("key_repeat")
            rb.setChecked(
                (val == "inherit" and cur is None) or
                (val == "on"      and cur is True) or
                (val == "off"     and cur is False)
            )
            self._rep_group.addButton(rb)
            setattr(self, f"_rep_{val}", rb)
            rep_row.addWidget(rb)
        rep_inner.addLayout(rep_row)
        layout.addWidget(rep_card)

        # Buttons
        btn_row = QHBoxLayout()
        clear_btn = danger_btn("CLEAR ALL")
        clear_btn.clicked.connect(self._clear_all)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        cancel_btn = QPushButton("CANCEL")
        cancel_btn.clicked.connect(self.reject)
        save_btn = success_btn("SAVE & SEND")
        save_btn.clicked.connect(self._save)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(save_btn)
        layout.addLayout(btn_row)

        self.installEventFilter(self)

    def _build_normal_page(self):
        page = QWidget()
        page.setStyleSheet("background: transparent;")
        vbox = QVBoxLayout(page)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(8)

        slots_card, slots_inner = make_card("TAP & HOLD")

        for slot, key, lbl in [("tap","tap","TAP"), ("hold","hold","HOLD")]:
            combo = self.cfg.get(key, [])
            display = "+".join(combo) if combo else "—"
            slot_frame = QFrame()
            slot_frame.setStyleSheet(
                f"background: {T.SURFACE3}; border: 1px solid {T.BORDER};"
                f" border-radius: 8px;"
            )
            slot_frame.setCursor(Qt.PointingHandCursor)
            sf_layout = QVBoxLayout(slot_frame)
            sf_layout.setContentsMargins(12, 8, 12, 8)
            title_lbl = QLabel(lbl)
            title_lbl.setStyleSheet(
                f"color: {T.TEXT_DIM}; font-size: 9px; font-weight: 700;"
                f" letter-spacing: 1px; border: none; background: transparent;"
            )
            val_lbl = QLabel(display)
            val_lbl.setStyleSheet(
                f"color: {T.BLUE_LT}; font-size: 14px; font-weight: 700;"
                f" font-family: 'Consolas', monospace; border: none; background: transparent;"
            )
            setattr(self, f"_{slot}_lbl", val_lbl)
            sf_layout.addWidget(title_lbl)
            sf_layout.addWidget(val_lbl)
            slot_frame.mousePressEvent = lambda e, s=slot: self._select_slot(s)
            slots_inner.addWidget(slot_frame)

        # Tap/hold toggle
        th_row = QHBoxLayout()
        th_lbl = QLabel("Tap/Hold enabled")
        th_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        self._th_toggle = ToggleSwitch(self.cfg.get("tap_hold_enabled", False))
        th_row.addWidget(th_lbl)
        th_row.addStretch()
        th_row.addWidget(self._th_toggle)
        slots_inner.addLayout(th_row)

        slots_inner.addWidget(hsep())

        # Capture label
        self._cap_lbl = QLabel("— select TAP or HOLD above —")
        self._cap_lbl.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 11px; padding: 8px 12px;"
            f" background: {T.SURFACE2}; border-radius: 6px; border: none;"
        )
        self._cap_lbl.setAlignment(Qt.AlignCenter)
        slots_inner.addWidget(self._cap_lbl)

        vbox.addWidget(slots_card)

        # Macro
        mac_card, mac_inner = make_card("MACRO  (overrides tap/hold)")
        mac_row = QHBoxLayout()
        steps = self.cfg.get("macro") or []
        self._mac_count_lbl = QLabel(
            f"{len(steps)} steps" if steps else "none"
        )
        self._mac_count_lbl.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 11px;"
        )
        mac_row.addWidget(self._mac_count_lbl)
        mac_row.addStretch()
        edit_mac_btn = accent_btn("EDIT MACRO")
        edit_mac_btn.clicked.connect(self._open_macro)
        clear_mac_btn = QPushButton("CLEAR")
        clear_mac_btn.clicked.connect(self._clear_macro)
        mac_row.addWidget(clear_mac_btn)
        mac_row.addWidget(edit_mac_btn)
        mac_inner.addLayout(mac_row)
        vbox.addWidget(mac_card)
        return page

    def _build_mo_page(self):
        page = QWidget()
        page.setStyleSheet("background: transparent;")
        vbox = QVBoxLayout(page)
        vbox.setContentsMargins(0, 0, 0, 0)
        card, inner = make_card("MOMENTARY LAYER")
        desc = QLabel("Hold to activate a layer, release to return.")
        desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        inner.addWidget(desc)
        row = QHBoxLayout()
        row.addWidget(QLabel("Target layer index:"))
        self._mo_spin = QSpinBox()
        self._mo_spin.setRange(0, 15)
        self._mo_spin.setValue(self.cfg.get("mo_layer") or 0)
        self._mo_spin.setFixedWidth(70)
        row.addWidget(self._mo_spin)
        row.addStretch()
        inner.addLayout(row)
        vbox.addWidget(card)
        vbox.addStretch()
        return page

    def _build_mouse_page(self):
        page = QWidget()
        page.setStyleSheet("background: transparent;")
        vbox = QVBoxLayout(page)
        vbox.setContentsMargins(0, 0, 0, 0)
        card, inner = make_card("MOUSE BUTTON HOLD")
        desc = QLabel("Holds a mouse button for the duration of the keypress.")
        desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        inner.addWidget(desc)
        self._mb_group = QButtonGroup(self)
        row = QHBoxLayout()
        cur = self.cfg.get("mouse_button") or "LEFT"
        for btn_name in ("LEFT", "RIGHT", "MIDDLE"):
            rb = QRadioButton(btn_name)
            rb.setChecked(cur == btn_name)
            self._mb_group.addButton(rb)
            setattr(self, f"_mb_{btn_name.lower()}", rb)
            row.addWidget(rb)
        inner.addLayout(row)
        vbox.addWidget(card)
        vbox.addStretch()
        return page

    def _build_enc_page(self):
        page = QWidget()
        page.setStyleSheet("background: transparent;")
        vbox = QVBoxLayout(page)
        vbox.setContentsMargins(0, 0, 0, 0)
        card, inner = make_card("ENCODER SPEED MODIFIER")
        desc = QLabel(
            "Multiplies all encoder speeds while held.\n"
            "0.1 = 10% speed (precision)  ·  2.0 = double speed"
        )
        desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        inner.addWidget(desc)
        row = QHBoxLayout()
        row.addWidget(QLabel("Factor:"))
        self._em_spin = QDoubleSpinBox()
        self._em_spin.setRange(0.05, 5.0)
        self._em_spin.setSingleStep(0.05)
        self._em_spin.setDecimals(2)
        self._em_spin.setValue(self.cfg.get("enc_mod_factor") or 0.1)
        self._em_spin.setFixedWidth(80)
        row.addWidget(self._em_spin)
        row.addStretch()
        inner.addLayout(row)
        vbox.addWidget(card)
        vbox.addStretch()
        return page

    def _on_type_change(self, kt):
        pages = {"normal":0,"mo":1,"mouse_hold":2,"enc_mod":3}
        self._stack.setCurrentIndex(pages.get(kt, 0))
        self.cfg["key_type"] = kt

    def _select_slot(self, slot):
        self._slot = slot
        self._cap_lbl.setText(f"Listening for  {slot.upper()}  slot…")
        self._cap_lbl.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 11px; padding: 8px 12px;"
            f" background: {T.BLUE_DIM}; border-radius: 6px; border: none;"
        )
        self.setFocus()

    def eventFilter(self, obj, event):
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.KeyPress:
            key = event.key()
            if key in MOD_MAP:
                self._held.add(MOD_MAP[key])
                return True
            else:
                name = event.text().upper()
                if not name or not name.isprintable():
                    name = None
                    for k, v in KEYSYM_MAP.items():
                        if hasattr(Qt, f"Key_{k}") and getattr(Qt, f"Key_{k}") == key:
                            name = v
                            break
                if name:
                    self._main_key = name
                return True
        if event.type() == QEvent.KeyRelease:
            key = event.key()
            if key not in MOD_MAP and hasattr(self, "_main_key") and self._main_key and self._slot:
                mods  = sorted(self._held)
                combo = mods + [self._main_key]
                display = "+".join(combo)
                self.cfg[self._slot] = combo
                lbl = getattr(self, f"_{self._slot}_lbl", None)
                if lbl:
                    lbl.setText(display)
                self._cap_lbl.setText(f"✓  {self._slot.upper()} = {display}")
                self._cap_lbl.setStyleSheet(
                    f"color: {T.GREEN}; font-size: 11px; padding: 8px 12px;"
                    f" background: {T.GREEN_DIM}; border-radius: 6px; border: none;"
                )
                self._held.clear()
                self._main_key = None
            return True
        return super().eventFilter(obj, event)

    def _open_macro(self):
        dlg = MacroEditorDialog(self.cfg.get("macro") or [], self)
        if dlg.exec() == QDialog.Accepted:
            steps = dlg.get_steps()
            self.cfg["macro"] = steps or None
            self._mac_count_lbl.setText(
                f"{len(steps)} steps" if steps else "none"
            )

    def _clear_macro(self):
        self.cfg["macro"] = None
        self._mac_count_lbl.setText("none")

    def _clear_all(self):
        reply = QMessageBox.question(
            self, "Clear key", "Clear all settings for this key?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.cfg = {
                "tap":[], "hold":[], "tap_hold_enabled":False,
                "macro":None, "key_repeat":None, "key_type":"normal",
                "mo_layer":None, "mouse_button":None, "enc_mod_factor":None,
            }
            self._type_radios["normal"].setChecked(True)
            self._on_type_change("normal")
            if hasattr(self, "_tap_lbl"): self._tap_lbl.setText("—")
            if hasattr(self, "_hold_lbl"): self._hold_lbl.setText("—")

    def _save(self):
        kt = self.cfg.get("key_type", "normal")
        if kt == "mo":
            self.cfg["mo_layer"] = self._mo_spin.value()
        elif kt == "mouse_hold":
            for btn_name in ("LEFT","RIGHT","MIDDLE"):
                rb = getattr(self, f"_mb_{btn_name.lower()}", None)
                if rb and rb.isChecked():
                    self.cfg["mouse_button"] = btn_name
                    break
        elif kt == "enc_mod":
            self.cfg["enc_mod_factor"] = self._em_spin.value()

        if hasattr(self, "_th_toggle"):
            self.cfg["tap_hold_enabled"] = self._th_toggle.isChecked()

        if hasattr(self, "_rep_inherit"):
            if self._rep_inherit.isChecked():
                self.cfg["key_repeat"] = None
            elif self._rep_on.isChecked():
                self.cfg["key_repeat"] = True
            else:
                self.cfg["key_repeat"] = False

        self.serial.send({
            "action": "set_key",
            "layer": self.layer_idx,
            "index": self.key_idx,
            **{k: self.cfg.get(k) for k in
               ("tap","hold","tap_hold_enabled","macro","key_repeat",
                "key_type","mo_layer","mouse_button","enc_mod_factor")}
        })
        self._result = (self.key_idx, self.cfg)
        self.accept()

    def get_result(self):
        return self._result


# ─────────────────────────────────────────────────────────────
#  9. TABS
# ─────────────────────────────────────────────────────────────

BUTTON_ACTIONS = [
    "—","F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12",
    "ESCAPE","SPACE","ENTER","TAB","BACKSPACE","DELETE",
    "HOME","END","PAGE_UP","PAGE_DOWN","UP","DOWN","LEFT","RIGHT",
    "MUTE","PLAY_PAUSE",
    "LEFT_CLICK","RIGHT_CLICK","MIDDLE_CLICK",
]

ENCODER_MODES = {
    "H_SCROLL":     "Horizontal Scroll",
    "V_SCROLL":     "Vertical Scroll",
    "ZOOM":         "Zoom (Ctrl+Scroll)",
    "UNDO_REDO":    "Undo / Redo",
    "NEXT_PREV_TAB":"Next / Prev Tab",
    "VOLUME":       "Volume",
}

# Orbit/Pan modifier presets — label → list of key names held with MMB
SM_MOD_PRESETS = {
    "MMB only (no modifier)":      [],
    "Shift + MMB":                 ["SHIFT"],
    "Ctrl + MMB":                  ["CTRL"],
    "Alt + MMB":                   ["ALT"],
    "Ctrl + Shift + MMB":          ["CTRL", "SHIFT"],
}
SM_MOD_LABELS = list(SM_MOD_PRESETS.keys())

def _mods_to_label(mods):
    """Convert a list of mod names to the matching preset label."""
    mods_set = set(mods or [])
    for label, preset in SM_MOD_PRESETS.items():
        if set(preset) == mods_set:
            return label
    return SM_MOD_LABELS[0]

def _make_template_keys(taps):
    """Build a 25-element key config list from a simple list of tap combos.
    Ghost positions (3,4,8,9,20) are left blank automatically."""
    blank = lambda: {
        "tap":[],"hold":[],"tap_hold_enabled":False,
        "macro":None,"key_repeat":None,"key_type":"normal",
        "mo_layer":None,"mouse_button":None,"enc_mod_factor":None,
    }
    keys = [blank() for _ in range(25)]
    for i, tap in enumerate(taps):
        if i < 25:
            keys[i]["tap"] = tap
    return keys

# Key layouts per app — 25 entries matching the 5x5 matrix
# Ghost positions (3,4,8,9,20) are blank
# Layout:  [0]  [1]  [2]   ·    ·
#          [5]  [6]  [7]   ·    ·
#          [10] [11] [12] [13] [14]
#          [15] [16] [17] [18] [19]
#           ·   [21] [22] [23] [24]

_KEYS_FUSION = [
    ["E"],              ["S"],             ["L"],             [], [],
    ["F"],              ["M"],             ["R"],             [], [],
    ["ESCAPE"],         ["D"],             ["O"],         ["Q"],    ["DELETE"],
    ["CTRL","Z"],       ["CTRL","Y"],      ["T"],         ["C"],    ["ENTER"],
    [],                 ["H"],             ["J"],         ["I"],    ["CTRL","S"],
]

_KEYS_ONSHAPE = [
    ["SHIFT","E"],      ["S"],             ["L"],             [], [],
    ["SHIFT","F"],      ["SHIFT","M"],     ["R"],             [], [],
    ["ESCAPE"],         ["D"],             ["SHIFT","O"], ["SHIFT","S"], ["DELETE"],
    ["CTRL","Z"],       ["CTRL","Y"],      ["T"],         ["C"],    ["ENTER"],
    [],                 ["SHIFT","H"],     ["SHIFT","L"], ["SHIFT","I"], ["CTRL","S"],
]

_KEYS_SOLIDWORKS = [
    ["E"],              ["S"],             ["L"],             [], [],
    ["F"],              ["M"],             ["R"],             [], [],
    ["ESCAPE"],         ["D"],             ["O"],         ["CTRL","SHIFT","M"], ["DELETE"],
    ["CTRL","Z"],       ["CTRL","Y"],      ["T"],         ["CTRL","SHIFT","P"], ["ENTER"],
    [],                 ["CTRL","SHIFT","C"],["CTRL","SHIFT","H"],["CTRL","SHIFT","L"],["CTRL","S"],
]

_KEYS_BLENDER = [
    ["E"],              ["S"],             ["G"],             [], [],
    ["CTRL","1"],       ["R"],             ["CTRL","R"],      [], [],
    ["ESCAPE"],         ["K"],             ["M"],         ["CTRL","B"], ["X"],
    ["CTRL","Z"],       ["CTRL","SHIFT","Z"],["2"],       ["3"],    ["ENTER"],
    [],                 ["I"],             ["F"],         ["CTRL","E"], ["CTRL","S"],
]

_KEYS_MAYA = [
    ["CTRL","E"],       ["CTRL","SHIFT","S"],["SHIFT","LEFT"], [], [],
    ["CTRL","F"],       ["W"],             ["SHIFT","RIGHT"], [], [],
    ["ESCAPE"],         ["B"],             ["O"],         ["CTRL","D"], ["DELETE"],
    ["CTRL","Z"],       ["CTRL","Y"],      ["F8"],        ["F9"],   ["ENTER"],
    [],                 ["CTRL","SHIFT","C"],["CTRL","SHIFT","H"],["I"],["CTRL","S"],
]

_KEYS_FREECAD = [
    ["E"],              ["S"],             ["L"],             [], [],
    ["F"],              ["M"],             ["R"],             [], [],
    ["ESCAPE"],         ["D"],             ["O"],         ["CTRL","SHIFT","M"], ["DELETE"],
    ["CTRL","Z"],       ["CTRL","Y"],      ["T"],         ["C"],    ["ENTER"],
    [],                 ["CTRL","SHIFT","C"],["H"],        ["I"],    ["CTRL","S"],
]

# Layer templates — name → partial layer overrides
LAYER_TEMPLATES = {
    "Blank": {},
    "Fusion 360": {
        "sm_active": True,
        "sm_orbit_mods": ["SHIFT"], "sm_pan_mods": [],
        "keys": _make_template_keys(_KEYS_FUSION),
    },
    "Onshape": {
        "sm_active": True,
        "sm_orbit_mods": [], "sm_pan_mods": ["CTRL"],
        "keys": _make_template_keys(_KEYS_ONSHAPE),
    },
    "SolidWorks": {
        "sm_active": True,
        "sm_orbit_mods": [], "sm_pan_mods": ["CTRL"],
        "keys": _make_template_keys(_KEYS_SOLIDWORKS),
    },
    "Blender": {
        "sm_active": True,
        "sm_orbit_mods": [], "sm_pan_mods": ["SHIFT"],
        "keys": _make_template_keys(_KEYS_BLENDER),
    },
    "Maya": {
        "sm_active": True,
        "sm_orbit_mods": ["ALT"], "sm_pan_mods": ["ALT"],
        "keys": _make_template_keys(_KEYS_MAYA),
    },
    "FreeCAD": {
        "sm_active": True,
        "sm_orbit_mods": [], "sm_pan_mods": ["CTRL"],
        "keys": _make_template_keys(_KEYS_FREECAD),
    },
}

# Physical 20-key layout — indices that have no switch (never fire)
# Col0: rows 0-3, Col1-2: rows 0-4, Col3-4: rows 2-4
GHOST_KEYS = frozenset([3, 4, 8, 9, 20])

KT_COLORS_DISPLAY = {
    "mo":         T.CYAN,
    "mouse_hold": T.PURPLE,
    "enc_mod":    T.YELLOW,
    "macro":      T.ORANGE,
}

def _key_label_and_type(kc):
    kt = kc.get("key_type","normal")
    if kt == "mo":
        return f"MO({kc.get('mo_layer',0)})", "mo"
    if kt == "mouse_hold":
        return f"🖱 {kc.get('mouse_button','?')}", "mouse_hold"
    if kt == "enc_mod":
        return f"×{kc.get('enc_mod_factor',0.1):.1f}", "enc_mod"
    if kc.get("macro"):
        return f"MACRO\n{len(kc['macro'])}steps", "macro"
    tap = kc.get("tap",[])
    hold = kc.get("hold",[])
    if tap and hold and kc.get("tap_hold_enabled"):
        return f"T:{'+'.join(tap)}\nH:{'+'.join(hold)}", "normal"
    return "+".join(tap) if tap else "", "normal"

def _key_tooltip(kc):
    """Build a rich tooltip showing full key configuration."""
    parts = []
    kt = kc.get("key_type", "normal")
    if kt == "mo":
        parts.append(f"Type: Momentary Layer {kc.get('mo_layer', 0)}")
    elif kt == "mouse_hold":
        parts.append(f"Type: Mouse Hold ({kc.get('mouse_button', 'LEFT')})")
    elif kt == "enc_mod":
        parts.append(f"Type: Encoder Modifier (×{kc.get('enc_mod_factor', 0.1):.1f})")
    else:
        tap = kc.get("tap", [])
        hold = kc.get("hold", [])
        if tap:
            parts.append(f"Tap: {'+'.join(tap)}")
        if hold and kc.get("tap_hold_enabled"):
            parts.append(f"Hold: {'+'.join(hold)}")
    macro = kc.get("macro")
    if macro:
        parts.append(f"Macro: {len(macro)} steps")
    rep = kc.get("key_repeat")
    if rep is True:
        parts.append("Repeat: ON")
    elif rep is False:
        parts.append("Repeat: OFF")
    return "\n".join(parts) if parts else "Empty key"


class MatrixTab(QWidget):
    key_updated = Signal(int, int, dict)

    def __init__(self, serial_mgr, parent=None):
        super().__init__(parent)
        self.serial  = serial_mgr
        self._cfg    = {}
        self._layer  = 0
        self._btns   = []
        self._selected_key = -1
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Top bar
        top = QHBoxLayout()
        top.addWidget(QLabel("Editing layer:"))
        self._layer_combo = QComboBox()
        self._layer_combo.setFixedWidth(160)
        self._layer_combo.currentIndexChanged.connect(self._on_layer_change)
        top.addWidget(self._layer_combo)
        top.addStretch()
        hint = QLabel("Click = side panel  •  Double-click = full editor")
        hint.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        top.addWidget(hint)
        layout.addLayout(top)

        # Splitter: grid left, editor right
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet(f"QSplitter::handle {{ background: {T.BORDER}; width: 2px; }}")

        # ── Left: key grid ──
        grid_card, grid_inner = make_card()
        grid = QGridLayout()
        grid.setSpacing(6)
        for i in range(25):
            r, c = divmod(i, 5)
            if i in GHOST_KEYS:
                btn = KeyButton(i)
                btn.setEnabled(False)
                btn.setText("")
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background: {T.BG};
                        border: 1px dashed {T.BORDER};
                        border-radius: 8px;
                    }}
                """)
                btn.setToolTip("No switch at this position")
            else:
                btn = KeyButton(i)
                btn.clicked.connect(lambda checked=False, idx=i: self._select_key(idx))
                btn.doubleClicked.connect(lambda idx: self._open_full_editor(idx))
            grid.addWidget(btn, r, c)
            self._btns.append(btn)
        grid_inner.addLayout(grid)
        splitter.addWidget(grid_card)

        # ── Right: editor panel ──
        self._editor_scroll = QScrollArea()
        self._editor_scroll.setWidgetResizable(True)
        self._editor_scroll.setStyleSheet("border: none;")
        self._editor_scroll.setMinimumWidth(320)

        self._editor_placeholder = QLabel("← Select a key to edit")
        self._editor_placeholder.setAlignment(Qt.AlignCenter)
        self._editor_placeholder.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 13px; padding: 40px;"
        )
        self._editor_scroll.setWidget(self._editor_placeholder)
        splitter.addWidget(self._editor_scroll)

        splitter.setSizes([420, 340])
        layout.addWidget(splitter, 1)

    def _select_key(self, idx):
        if idx in GHOST_KEYS:
            return
        # Highlight selected key
        for i, btn in enumerate(self._btns):
            if i in GHOST_KEYS:
                continue
            if i == idx:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        border: 2px solid {T.BLUE_LT};
                        border-radius: 8px;
                    }}
                """)
            else:
                btn.set_key_data(btn._label, btn._type)  # reset style

        self._selected_key = idx
        layers = self._cfg.get("layers", [])
        li = self._layer_combo.currentIndex()
        if li < 0 or li >= len(layers):
            return
        kc = layers[li]["keys"][idx] if idx < 25 else {}
        # Build inline editor
        self._build_editor_panel(idx, li, kc)

    def _build_editor_panel(self, key_idx, layer_idx, kc):
        """Build the side panel editor for a specific key."""
        # Clean up previous panel to prevent widget accumulation
        old = self._editor_scroll.widget()
        if old and old is not self._editor_placeholder:
            old.setParent(None)
            old.deleteLater()

        cfg = copy.deepcopy(kc)

        panel = QWidget()
        panel.setStyleSheet(f"background: {T.SURFACE};")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Header
        r, c = divmod(key_idx, 5)
        header = QLabel(f"KEY [{r},{c}]  —  INDEX {key_idx}")
        header.setStyleSheet(
            f"color: {T.BLUE_LT}; font-size: 12px; font-weight: 700;"
            f" letter-spacing: 1px;"
        )
        layout.addWidget(header)
        layout.addWidget(hsep())

        # Key type selector
        layout.addWidget(section_label("Key Type"))
        type_group = QButtonGroup(panel)
        types = [("normal","Normal"),("mo","MO (Layer)"),
                 ("mouse_hold","Mouse Hold"),("enc_mod","Encoder Mod")]
        type_row = QHBoxLayout()
        type_radios = {}
        for val, label in types:
            rb = QRadioButton(label)
            rb.setStyleSheet(f"font-size: 11px;")
            type_group.addButton(rb)
            type_radios[val] = rb
            type_row.addWidget(rb)
        type_row.addStretch()
        layout.addLayout(type_row)
        current_type = cfg.get("key_type", "normal")
        if current_type in type_radios:
            type_radios[current_type].setChecked(True)

        # Type-specific options stacked widget
        type_stack = QStackedWidget()

        # Normal page
        normal_page = QWidget()
        np_layout = QVBoxLayout(normal_page)
        np_layout.setContentsMargins(0, 4, 0, 0)
        np_layout.setSpacing(6)

        tap_display = "+".join(cfg.get("tap",[])) or "—"
        hold_display = "+".join(cfg.get("hold",[])) or "—"
        tap_lbl = QLabel(f"Tap:  {tap_display}")
        tap_lbl.setStyleSheet(f"color: {T.TEXT}; font-size: 11px; font-family: Consolas;")
        hold_lbl = QLabel(f"Hold:  {hold_display}")
        hold_lbl.setStyleSheet(f"color: {T.TEXT}; font-size: 11px; font-family: Consolas;")
        np_layout.addWidget(tap_lbl)
        np_layout.addWidget(hold_lbl)

        th_row = QHBoxLayout()
        th_row.addWidget(QLabel("Tap/Hold enabled:"))
        th_toggle = ToggleSwitch(cfg.get("tap_hold_enabled", False))
        th_row.addStretch()
        th_row.addWidget(th_toggle)
        np_layout.addLayout(th_row)

        macro_steps = cfg.get("macro") or []
        mac_lbl = QLabel(f"Macro: {len(macro_steps)} steps" if macro_steps else "Macro: none")
        mac_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        np_layout.addWidget(mac_lbl)
        np_layout.addStretch()
        type_stack.addWidget(normal_page)

        # MO page
        mo_page = QWidget()
        mo_layout = QVBoxLayout(mo_page)
        mo_layout.setContentsMargins(0, 4, 0, 0)
        mo_lbl = QLabel("Target layer:")
        mo_spin = QSpinBox()
        mo_spin.setRange(0, 9)
        mo_spin.setValue(cfg.get("mo_layer", 0) or 0)
        mo_row = QHBoxLayout()
        mo_row.addWidget(mo_lbl)
        mo_row.addWidget(mo_spin)
        mo_row.addStretch()
        mo_layout.addLayout(mo_row)
        mo_layout.addStretch()
        type_stack.addWidget(mo_page)

        # Mouse hold page
        mh_page = QWidget()
        mh_layout = QVBoxLayout(mh_page)
        mh_layout.setContentsMargins(0, 4, 0, 0)
        mb_group = QButtonGroup(panel)
        cur_mb = cfg.get("mouse_button", "LEFT") or "LEFT"
        for bname in ("LEFT","RIGHT","MIDDLE"):
            rb = QRadioButton(f"{bname} button")
            mb_group.addButton(rb)
            if bname == cur_mb:
                rb.setChecked(True)
            setattr(panel, f"_mb_{bname.lower()}", rb)
            mh_layout.addWidget(rb)
        mh_layout.addStretch()
        type_stack.addWidget(mh_page)

        # Encoder mod page
        em_page = QWidget()
        em_layout = QVBoxLayout(em_page)
        em_layout.setContentsMargins(0, 4, 0, 0)
        em_row = QHBoxLayout()
        em_row.addWidget(QLabel("Speed factor:"))
        em_spin = QDoubleSpinBox()
        em_spin.setRange(0.01, 10.0)
        em_spin.setSingleStep(0.1)
        em_spin.setValue(cfg.get("enc_mod_factor", 0.1) or 0.1)
        em_row.addWidget(em_spin)
        em_row.addStretch()
        em_layout.addLayout(em_row)
        em_layout.addStretch()
        type_stack.addWidget(em_page)

        layout.addWidget(type_stack)

        type_map = {"normal": 0, "mo": 1, "mouse_hold": 2, "enc_mod": 3}
        type_stack.setCurrentIndex(type_map.get(current_type, 0))

        def _on_type(rb, checked):
            if not checked:
                return
            for val, radio in type_radios.items():
                if radio is rb:
                    type_stack.setCurrentIndex(type_map.get(val, 0))
                    cfg["key_type"] = val
                    break
        for val, rb in type_radios.items():
            rb.toggled.connect(lambda c, r=rb: _on_type(r, c))

        layout.addWidget(hsep())

        # Buttons: Edit Full / Apply / Open Macro
        btn_row = QHBoxLayout()
        edit_btn = QPushButton("EDIT FULL...")
        edit_btn.setToolTip("Open the full key editor dialog for tap/hold capture and macros")
        edit_btn.clicked.connect(lambda: self._open_full_editor(key_idx))
        edit_btn.setStyleSheet(
            f"background: {T.SURFACE3}; color: {T.TEXT}; border: 1px solid {T.BORDER2};"
            f" border-radius: 6px; padding: 6px 14px; font-weight: 600; font-size: 11px;"
        )
        btn_row.addWidget(edit_btn)

        apply_btn = accent_btn("APPLY")
        apply_btn.setFixedWidth(80)

        def _apply():
            kt = cfg.get("key_type", "normal")
            if kt == "mo":
                cfg["mo_layer"] = mo_spin.value()
            elif kt == "mouse_hold":
                for bname in ("LEFT","RIGHT","MIDDLE"):
                    rb = getattr(panel, f"_mb_{bname.lower()}", None)
                    if rb and rb.isChecked():
                        cfg["mouse_button"] = bname
                        break
            elif kt == "enc_mod":
                cfg["enc_mod_factor"] = em_spin.value()
            cfg["tap_hold_enabled"] = th_toggle.isChecked()
            self.serial.send({
                "action": "set_key", "layer": layer_idx, "index": key_idx,
                **{k: cfg.get(k) for k in
                   ("tap","hold","tap_hold_enabled","macro","key_repeat",
                    "key_type","mo_layer","mouse_button","enc_mod_factor")}
            })
            layers = self._cfg.get("layers", [])
            if layer_idx < len(layers):
                layers[layer_idx]["keys"][key_idx] = cfg
            self._refresh_grid()
            self._select_key(key_idx)   # refresh panel

        apply_btn.clicked.connect(_apply)
        btn_row.addWidget(apply_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._editor_scroll.setWidget(panel)

    def _open_full_editor(self, idx):
        """Open the full KeyEditorDialog for tap/hold capture and macro editing."""
        if idx in GHOST_KEYS:
            return
        layers = self._cfg.get("layers", [])
        li = self._layer_combo.currentIndex()
        if li < 0 or li >= len(layers):
            return
        kc = layers[li]["keys"][idx] if idx < 25 else {}
        dlg = KeyEditorDialog(idx, li, kc, self.serial, self)
        if dlg.exec() == QDialog.Accepted:
            result = dlg.get_result()
            if result:
                ki, kdata = result
                layers[li]["keys"][ki] = kdata
                self._refresh_grid(li)
                self.key_updated.emit(li, ki, kdata)
                self._select_key(ki)   # refresh side panel

    def apply_config(self, cfg):
        self._cfg = cfg
        layers = cfg.get("layers", [])
        names  = [l.get("name","") for l in layers]
        prev   = self._layer_combo.currentText()
        self._layer_combo.blockSignals(True)
        self._layer_combo.clear()
        self._layer_combo.addItems(names)
        if prev in names:
            self._layer_combo.setCurrentText(prev)
        self._layer_combo.blockSignals(False)
        self._refresh_grid()

    def set_active_layer(self, idx):
        layers = self._cfg.get("layers",[])
        if 0 <= idx < len(layers):
            self._layer_combo.setCurrentIndex(idx)

    def _on_layer_change(self):
        self._refresh_grid()
        if self._selected_key >= 0:
            self._select_key(self._selected_key)

    def _refresh_grid(self, layer_idx=None):
        layers = self._cfg.get("layers",[])
        if not layers: return
        if layer_idx is None:
            layer_idx = self._layer_combo.currentIndex()
            if layer_idx < 0: layer_idx = 0
        self._layer = layer_idx
        if layer_idx >= len(layers): return
        keys = layers[layer_idx].get("keys",[])
        for i, btn in enumerate(self._btns):
            if i in GHOST_KEYS: continue
            kc = keys[i] if i < len(keys) else {}
            label, kt = _key_label_and_type(kc)
            btn.set_key_data(label, kt)
            btn.setToolTip(_key_tooltip(kc))

    def flash_key(self, idx):
        if idx not in GHOST_KEYS and 0 <= idx < len(self._btns):
            self._btns[idx].flash()


class LayersTab(QWidget):
    def __init__(self, serial_mgr, parent=None):
        super().__init__(parent)
        self.serial = serial_mgr
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner = QWidget()
        inner.setStyleSheet(f"background: {T.SURFACE};")
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Layer list card
        lm_card, lm_inner = make_card("LAYER MANAGEMENT")
        list_row = QHBoxLayout()
        self._layer_list = QListWidget()
        self._layer_list.setMinimumHeight(140)
        list_row.addWidget(self._layer_list, 1)
        btn_col = QVBoxLayout()
        btn_col.setSpacing(6)
        add_btn    = accent_btn("ADD")
        remove_btn = danger_btn("REMOVE")
        rename_btn = QPushButton("RENAME")
        add_btn.clicked.connect(self._add)
        remove_btn.clicked.connect(self._remove)
        rename_btn.clicked.connect(self._rename)
        for b in (add_btn, remove_btn, rename_btn):
            b.setFixedWidth(90)
            btn_col.addWidget(b)
        btn_col.addStretch()
        list_row.addLayout(btn_col)
        lm_inner.addLayout(list_row)
        active_row = QHBoxLayout()
        active_row.addWidget(QLabel("Active layer:"))
        self._active_lbl = QLabel("—")
        self._active_lbl.setStyleSheet(
            f"color: {T.BLUE_LT}; font-weight: 700; font-size: 12px;"
        )
        active_row.addWidget(self._active_lbl)
        active_row.addStretch()
        lm_inner.addLayout(active_row)
        layout.addWidget(lm_card)

        # Space mouse per-layer toggle
        sm_card, sm_inner = make_card("SPACE MOUSE  (per layer)")
        sm_desc = QLabel(
            "Enable the space mouse and configure orbit/pan key combos per layer."
        )
        sm_desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        sm_desc.setWordWrap(True)
        sm_inner.addWidget(sm_desc)
        sm_row = QHBoxLayout()
        sm_lbl = QLabel("Space mouse active on selected layer:")
        sm_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        self._sm_toggle = ToggleSwitch(False)
        self._sm_toggle.toggled.connect(self._send_sm_active)
        sm_row.addWidget(sm_lbl)
        sm_row.addStretch()
        sm_row.addWidget(self._sm_toggle)
        sm_inner.addLayout(sm_row)
        sm_inner.addWidget(hsep())

        # Orbit combo
        orb_row = QHBoxLayout()
        orb_row.addWidget(QLabel("Orbit:"))
        self._orbit_combo = QComboBox()
        self._orbit_combo.addItems(SM_MOD_LABELS)
        self._orbit_combo.setFixedWidth(220)
        self._orbit_combo.currentTextChanged.connect(self._send_orbit_mods)
        orb_row.addWidget(self._orbit_combo)
        orb_row.addWidget(QLabel("+ MMB"))
        orb_row.addStretch()
        sm_inner.addLayout(orb_row)

        # Pan combo
        pan_row = QHBoxLayout()
        pan_row.addWidget(QLabel("Pan:"))
        self._pan_combo = QComboBox()
        self._pan_combo.addItems(SM_MOD_LABELS)
        self._pan_combo.setFixedWidth(220)
        self._pan_combo.currentTextChanged.connect(self._send_pan_mods)
        pan_row.addWidget(self._pan_combo)
        pan_row.addWidget(QLabel("+ MMB"))
        pan_row.addStretch()
        sm_inner.addLayout(pan_row)

        layout.addWidget(sm_card)

        # ── Encoder modes (per layer) ──
        enc_card, enc_inner = make_card("ENCODER MODES  (per layer — follows selection above)")
        modes_list = list(ENCODER_MODES.values())
        for enc in (1, 2):
            row = QHBoxLayout()
            row.addWidget(QLabel(f"Encoder {enc}:"))
            cb = QComboBox()
            cb.addItems(modes_list)
            cb.setFixedWidth(200)
            cb.currentIndexChanged.connect(
                lambda _, e=enc: self._send_enc_mode(e)
            )
            setattr(self, f"_enc{enc}_mode_cb", cb)
            row.addWidget(cb)
            row.addStretch()
            enc_inner.addLayout(row)
        enc_inner.addWidget(hsep())
        # Encoder switch actions
        for enc in (1, 2):
            row = QHBoxLayout()
            row.addWidget(QLabel(f"Enc {enc} switch:"))
            cb = QComboBox()
            cb.addItems(BUTTON_ACTIONS[1:])
            cb.setFixedWidth(200)
            cb.currentTextChanged.connect(
                lambda v, e=enc: self._send_enc_sw(e, v)
            )
            setattr(self, f"_enc{enc}_sw_cb", cb)
            row.addWidget(cb)
            row.addStretch()
            enc_inner.addLayout(row)
        enc_inner.addWidget(hsep())
        # Global encoder settings
        enc_inner.addWidget(section_label("Global Encoder Settings"))
        self._e1_speed  = SliderRow("Encoder 1 Speed", "enc1_speed", 1, 50, 1, 20)
        self._e2_speed  = SliderRow("Encoder 2 Speed", "enc2_speed", 1, 50, 1, 20)
        self._e1_speed.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"enc1_speed","value":v})
        )
        self._e2_speed.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"enc2_speed","value":v})
        )
        enc_inner.addWidget(self._e1_speed)
        e1_inv_row = QHBoxLayout()
        e1_inv_lbl = QLabel("Encoder 1 Invert")
        e1_inv_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        self._e1_invert = ToggleSwitch()
        self._e1_invert.toggled.connect(
            lambda v: self.serial.send({"action":"set","key":"enc1_invert","value":v})
        )
        e1_inv_row.addWidget(e1_inv_lbl)
        e1_inv_row.addStretch()
        e1_inv_row.addWidget(self._e1_invert)
        enc_inner.addLayout(e1_inv_row)
        enc_inner.addWidget(self._e2_speed)
        e2_inv_row = QHBoxLayout()
        e2_inv_lbl = QLabel("Encoder 2 Invert")
        e2_inv_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        self._e2_invert = ToggleSwitch()
        self._e2_invert.toggled.connect(
            lambda v: self.serial.send({"action":"set","key":"enc2_invert","value":v})
        )
        e2_inv_row.addWidget(e2_inv_lbl)
        e2_inv_row.addStretch()
        e2_inv_row.addWidget(self._e2_invert)
        enc_inner.addLayout(e2_inv_row)
        layout.addWidget(enc_card)

        # _cfg must be initialised before the signal connection below
        self._cfg = {}
        self._layer_list.currentRowChanged.connect(self._on_layer_sel_changed)

        # Buttons & settings card
        set_card, set_inner = make_card("BUTTONS, TAP/HOLD & KEY REPEAT")
        e1_row = QHBoxLayout()
        e1_row.addWidget(QLabel("Extra Button 1 action:"))
        self._btn1_combo = QComboBox()
        self._btn1_combo.addItems(BUTTON_ACTIONS[1:])
        self._btn1_combo.setCurrentText("F6")
        self._btn1_combo.currentTextChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"btn_extra1","value":v})
        )
        e1_row.addWidget(self._btn1_combo)
        e1_row.addStretch()
        set_inner.addLayout(e1_row)
        note = QLabel("Hold Extra Btn 1 for 300ms → enc2 becomes ZOOM until released")
        note.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 10px;")
        set_inner.addWidget(note)
        set_inner.addWidget(hsep())

        self._tap_hold_row = SliderRow(
            "Tap/Hold threshold (ms)", "tap_hold_ms", 50, 600, 10, 200
        )
        self._tap_hold_row.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"tap_hold_ms","value":v})
        )
        set_inner.addWidget(self._tap_hold_row)
        set_inner.addWidget(hsep())

        set_inner.addWidget(section_label("Key Repeat"))
        rep_top = QHBoxLayout()
        rep_top.addWidget(QLabel("Key repeat enabled:"))
        self._rep_toggle = ToggleSwitch(True)
        self._rep_toggle.toggled.connect(
            lambda v: self.serial.send({"action":"set","key":"key_repeat_enabled","value":v})
        )
        rep_top.addStretch()
        rep_top.addWidget(self._rep_toggle)
        set_inner.addLayout(rep_top)

        self._rep_delay = SliderRow(
            "Delay before repeat (ms)", "key_repeat_delay_ms", 100, 1000, 10, 400
        )
        self._rep_delay.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"key_repeat_delay_ms","value":v})
        )
        self._rep_rate = SliderRow(
            "Repeat interval (ms)", "key_repeat_rate_ms", 10, 200, 5, 50
        )
        self._rep_rate.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"key_repeat_rate_ms","value":v})
        )
        set_inner.addWidget(self._rep_delay)
        set_inner.addWidget(self._rep_rate)
        layout.addWidget(set_card)
        layout.addStretch()

        scroll.setWidget(inner)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    def apply_config(self, cfg):
        self._cfg = cfg
        layers = cfg.get("layers",[])
        al     = cfg.get("active_layer", 0)
        self._layer_list.blockSignals(True)
        self._layer_list.clear()
        for i, l in enumerate(layers):
            name = l.get("name","")
            sm   = " 🖱" if l.get("sm_active", False) else ""
            item = QListWidgetItem(f"  {i}:  {name}{sm}{'  ◀' if i==al else ''}")
            self._layer_list.addItem(item)
        self._layer_list.blockSignals(False)
        al_name = layers[al]["name"] if al < len(layers) else "—"
        self._active_lbl.setText(f"[{al}] {al_name}")
        row = self._layer_list.currentRow()
        if row < 0: row = al
        if row < len(layers):
            layer = layers[row]
            self._sm_toggle.setChecked(layer.get("sm_active", False))
            self._orbit_combo.blockSignals(True)
            self._orbit_combo.setCurrentText(_mods_to_label(layer.get("sm_orbit_mods", ["SHIFT"])))
            self._orbit_combo.blockSignals(False)
            self._pan_combo.blockSignals(True)
            self._pan_combo.setCurrentText(_mods_to_label(layer.get("sm_pan_mods", [])))
            self._pan_combo.blockSignals(False)
        self._refresh_enc_modes()
        if "btn_extra1" in cfg: self._btn1_combo.setCurrentText(cfg["btn_extra1"])
        if "tap_hold_ms" in cfg: self._tap_hold_row.set_value(cfg["tap_hold_ms"])
        if "key_repeat_enabled" in cfg: self._rep_toggle.setChecked(cfg["key_repeat_enabled"])
        if "key_repeat_delay_ms" in cfg: self._rep_delay.set_value(cfg["key_repeat_delay_ms"])
        if "key_repeat_rate_ms" in cfg: self._rep_rate.set_value(cfg["key_repeat_rate_ms"])
        if "enc1_speed" in cfg: self._e1_speed.set_value(cfg["enc1_speed"])
        if "enc2_speed" in cfg: self._e2_speed.set_value(cfg["enc2_speed"])
        if "enc1_invert" in cfg: self._e1_invert.setChecked(cfg["enc1_invert"])
        if "enc2_invert" in cfg: self._e2_invert.setChecked(cfg["enc2_invert"])

    def set_active(self, idx, name):
        self._active_lbl.setText(f"[{idx}] {name}")

    def set_sm_active(self, val):
        self._sm_toggle.blockSignals(True)
        self._sm_toggle.setChecked(val)
        self._sm_toggle.blockSignals(False)

    def _on_layer_sel_changed(self, row):
        layers = self._cfg.get("layers", [])
        if not layers or row < 0 or row >= len(layers):
            return
        layer = layers[row]
        self._sm_toggle.blockSignals(True)
        self._sm_toggle.setChecked(layer.get("sm_active", False))
        self._sm_toggle.blockSignals(False)
        # Orbit/pan combos
        self._orbit_combo.blockSignals(True)
        self._orbit_combo.setCurrentText(_mods_to_label(layer.get("sm_orbit_mods", ["SHIFT"])))
        self._orbit_combo.blockSignals(False)
        self._pan_combo.blockSignals(True)
        self._pan_combo.setCurrentText(_mods_to_label(layer.get("sm_pan_mods", [])))
        self._pan_combo.blockSignals(False)
        self._refresh_enc_modes()

    def _refresh_enc_modes(self):
        layers = self._cfg.get("layers", [])
        idx = self._layer_list.currentRow()
        if idx < 0 or idx >= len(layers):
            return
        layer = layers[idx]
        mode_vals = list(ENCODER_MODES.keys())
        for enc in (1, 2):
            mode = layer.get(f"enc{enc}_mode", "V_SCROLL")
            cb = getattr(self, f"_enc{enc}_mode_cb")
            cb.blockSignals(True)
            if mode in ENCODER_MODES:
                cb.setCurrentText(ENCODER_MODES[mode])
            cb.blockSignals(False)
            sw = layer.get(f"enc{enc}_sw", "")
            sw_cb = getattr(self, f"_enc{enc}_sw_cb")
            sw_cb.blockSignals(True)
            sw_cb.setCurrentText(sw)
            sw_cb.blockSignals(False)

    def _send_enc_mode(self, enc):
        layers = self._cfg.get("layers", [])
        li = self._layer_list.currentRow()
        if li < 0 or li >= len(layers):
            return
        lmap = {v: k for k, v in ENCODER_MODES.items()}
        cb = getattr(self, f"_enc{enc}_mode_cb")
        mode = lmap.get(cb.currentText(), "V_SCROLL")
        self.serial.send({"action": "set_encoder_mode", "layer": li, "enc": enc, "mode": mode})

    def _send_enc_sw(self, enc, val):
        layers = self._cfg.get("layers", [])
        li = self._layer_list.currentRow()
        if li < 0 or li >= len(layers):
            return
        self.serial.send({"action": "set_enc_sw", "layer": li, "enc": enc, "value": val})

    def _send_sm_active(self, val):
        row = self._layer_list.currentRow()
        if row < 0: return
        layers = self._cfg.get("layers", [])
        if row < len(layers):
            layers[row]["sm_active"] = val
            self.serial.send({"action":"set_layer_prop","layer":row,
                              "key":"sm_active","value":val})
            self.apply_config(self._cfg)

    def _send_orbit_mods(self, label):
        row = self._layer_list.currentRow()
        if row < 0: return
        mods = SM_MOD_PRESETS.get(label, ["SHIFT"])
        layers = self._cfg.get("layers", [])
        if row < len(layers):
            layers[row]["sm_orbit_mods"] = mods
            self.serial.send({"action":"set_layer_prop","layer":row,
                              "key":"sm_orbit_mods","value":mods})

    def _send_pan_mods(self, label):
        row = self._layer_list.currentRow()
        if row < 0: return
        mods = SM_MOD_PRESETS.get(label, [])
        layers = self._cfg.get("layers", [])
        if row < len(layers):
            layers[row]["sm_pan_mods"] = mods
            self.serial.send({"action":"set_layer_prop","layer":row,
                              "key":"sm_pan_mods","value":mods})

    def _add(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Add Layer")
        dlg.setMinimumWidth(360)
        lay = QVBoxLayout(dlg)

        # Name
        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Name:"))
        name_edit = QLineEdit()
        name_edit.setPlaceholderText("Layer name")
        name_row.addWidget(name_edit)
        lay.addLayout(name_row)

        lay.addWidget(hsep())

        # Source: template or copy
        src_group = QButtonGroup(dlg)
        rb_blank = QRadioButton("Blank layer")
        rb_template = QRadioButton("From template:")
        rb_copy = QRadioButton("Copy existing layer:")
        rb_blank.setChecked(True)
        src_group.addButton(rb_blank)
        src_group.addButton(rb_template)
        src_group.addButton(rb_copy)
        lay.addWidget(rb_blank)

        # Template row
        tmpl_row = QHBoxLayout()
        tmpl_row.addWidget(rb_template)
        tmpl_combo = QComboBox()
        tmpl_combo.addItems([k for k in LAYER_TEMPLATES.keys() if k != "Blank"])
        tmpl_combo.setFixedWidth(200)
        tmpl_row.addWidget(tmpl_combo)
        tmpl_row.addStretch()
        lay.addLayout(tmpl_row)

        # Copy row
        copy_row = QHBoxLayout()
        copy_row.addWidget(rb_copy)
        copy_combo = QComboBox()
        layers = self._cfg.get("layers", [])
        for i, l in enumerate(layers):
            copy_combo.addItem(f"{i}: {l.get('name','')}")
        copy_combo.setFixedWidth(200)
        copy_row.addWidget(copy_combo)
        copy_row.addStretch()
        lay.addLayout(copy_row)

        lay.addWidget(hsep())

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = accent_btn("ADD")
        cancel_btn = QPushButton("CANCEL")
        ok_btn.clicked.connect(dlg.accept)
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(ok_btn)
        lay.addLayout(btn_row)

        if dlg.exec() != QDialog.Accepted:
            return
        name = name_edit.text().strip() or f"Layer {len(layers)}"

        cmd = {"action": "add_layer", "name": name}
        if rb_template.isChecked():
            tmpl_name = tmpl_combo.currentText()
            tmpl = LAYER_TEMPLATES.get(tmpl_name, {})
            cmd.update(tmpl)
            if not name_edit.text().strip():
                cmd["name"] = tmpl_name
        elif rb_copy.isChecked():
            cmd["copy_from"] = copy_combo.currentIndex()
        self.serial.send(cmd)

    def _remove(self):
        row = self._layer_list.currentRow()
        if row < 0: return
        reply = QMessageBox.question(
            self, "Remove Layer", f"Remove layer {row}?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.serial.send({"action":"remove_layer","index":row})

    def _rename(self):
        row = self._layer_list.currentRow()
        if row < 0: return
        name, ok = QInputDialog.getText(self, "Rename Layer", "New name:")
        if ok and name:
            self.serial.send({"action":"rename_layer","index":row,"name":name})


class InputTab(QWidget):
    """Combined Joystick + Space Mouse settings tab."""
    def __init__(self, serial_mgr, parent=None):
        super().__init__(parent)
        self.serial = serial_mgr
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner = QWidget()
        inner.setStyleSheet(f"background: {T.SURFACE};")
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # ── Joystick section ──
        joy_card, joy_inner = make_card("ANALOG JOYSTICK")
        self._speed = SliderRow("Speed Multiplier", "joy_speed", 0.1, 5.0, 0.1, 1.0)
        self._dz    = SliderRow("Deadzone (raw units)", "joy_deadzone", 200, 8000, 100, 2000)
        self._speed.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"joy_speed","value":v})
        )
        self._dz.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"joy_deadzone","value":v})
        )
        joy_inner.addWidget(self._speed)
        joy_inner.addWidget(self._dz)
        joy_inner.addWidget(hsep())

        for label, key in (("Invert X axis","joy_invert_x"),("Invert Y axis","joy_invert_y")):
            row = QHBoxLayout()
            lbl = QLabel(label)
            lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
            tog = ToggleSwitch()
            tog.toggled.connect(
                lambda v, k=key: self.serial.send({"action":"set","key":k,"value":v})
            )
            setattr(self, f"_{'ix' if 'X' in label else 'iy'}_toggle", tog)
            row.addWidget(lbl)
            row.addStretch()
            row.addWidget(tog)
            joy_inner.addLayout(row)
        joy_inner.addWidget(hsep())

        act_row = QHBoxLayout()
        act_row.addWidget(QLabel("Click action:"))
        self._joy_sw_combo = QComboBox()
        self._joy_sw_combo.addItems(BUTTON_ACTIONS[1:])
        self._joy_sw_combo.setCurrentText("LEFT_CLICK")
        self._joy_sw_combo.currentTextChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"joy_sw","value":v})
        )
        act_row.addWidget(self._joy_sw_combo)
        act_row.addStretch()
        joy_inner.addLayout(act_row)
        layout.addWidget(joy_card)

        # ── Space Mouse section ──
        mv_card, mv_inner = make_card("SPACE MOUSE — MOVEMENT")
        self._sens = SliderRow("Sensitivity",        "sm_sensitivity",  1, 50,  0.5, 15.0)
        self._sm_dz = SliderRow("XY Deadzone",       "sm_deadzone",     10, 400, 5,   100.0)
        self._zt   = SliderRow("Z Threshold",        "sm_z_threshold",  10, 400, 5,   100.0)
        for s in (self._sens, self._sm_dz, self._zt):
            s.valueChanged.connect(lambda v, k=s.key: self.serial.send({"action":"set","key":k,"value":v}))
            mv_inner.addWidget(s)
        layout.addWidget(mv_card)

        # Filter & accel
        fa_card, fa_inner = make_card("SPACE MOUSE — FILTER & ACCELERATION")
        self._filt  = SliderRow("Filter (lower = smoother)", "sm_filter",       0.01, 1.0, 0.01, 0.25)
        self._adapt = SliderRow("Adaptive rate (0 = off)",   "sm_adapt",        0.0,  0.02, 0.001, 0.003)
        self._curve = SliderRow("Accel curve (1 = linear)",  "sm_accel_curve",  1.0,  4.0, 0.1,  2.0)
        self._filt.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_filter","value":v})
        )
        self._adapt.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_adapt","value":v})
        )
        self._curve.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_accel_curve","value":v})
        )
        fa_inner.addWidget(self._filt)
        fa_inner.addWidget(self._adapt)
        adapt_note = QLabel(
            "Adaptive filter: when still, uses Filter value (smooth). "
            "When moving, alpha increases automatically (responsive). "
            "Set to 0 for fixed filter."
        )
        adapt_note.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 10px;")
        adapt_note.setWordWrap(True)
        fa_inner.addWidget(adapt_note)
        fa_inner.addWidget(self._curve)
        fa_inner.addWidget(hsep())
        accel_row = QHBoxLayout()
        accel_lbl = QLabel("Acceleration enabled:")
        accel_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        self._accel_toggle = ToggleSwitch(True)
        self._accel_toggle.toggled.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_accel","value":v})
        )
        accel_row.addWidget(accel_lbl)
        accel_row.addStretch()
        accel_row.addWidget(self._accel_toggle)
        fa_inner.addLayout(accel_row)
        layout.addWidget(fa_card)

        # Hysteresis
        hy_card, hy_inner = make_card("SPACE MOUSE — ORBIT HYSTERESIS")
        self._enter = SliderRow("Enter orbit (ms)", "sm_orbit_enter_ms", 0, 200, 5, 40)
        self._exit  = SliderRow("Exit orbit (ms)",  "sm_orbit_exit_ms",  0, 300, 5, 80)
        self._enter.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_orbit_enter_ms","value":v})
        )
        self._exit.valueChanged.connect(
            lambda v: self.serial.send({"action":"set","key":"sm_orbit_exit_ms","value":v})
        )
        hy_inner.addWidget(self._enter)
        hy_inner.addWidget(self._exit)
        note = QLabel("Higher values = more stable, less twitchy")
        note.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 10px;")
        hy_inner.addWidget(note)
        layout.addWidget(hy_card)

        # Z-axis mode
        zm_card, zm_inner = make_card("SPACE MOUSE — Z-AXIS MODE")
        zm_row = QHBoxLayout()
        self._zm_group = QButtonGroup(self)
        self._zm_zoom = QRadioButton("Zoom (scroll wheel)")
        self._zm_pan  = QRadioButton("Pan (middle drag)")
        self._zm_zoom.setChecked(True)
        self._zm_group.addButton(self._zm_zoom)
        self._zm_group.addButton(self._zm_pan)
        zm_row.addWidget(self._zm_zoom)
        zm_row.addWidget(self._zm_pan)
        zm_row.addStretch()
        for rb in (self._zm_zoom, self._zm_pan):
            rb.toggled.connect(self._send_z_mode)
        zm_inner.addLayout(zm_row)
        layout.addWidget(zm_card)

        # Zero button
        zero_row = QHBoxLayout()
        zero_btn = QPushButton("⊕  ZERO SPACE MOUSE")
        zero_btn.setStyleSheet(
            f"background: {T.YELLOW}; color: black; font-weight: 700;"
            f" border: none; border-radius: 6px; padding: 7px 18px;"
        )
        zero_btn.clicked.connect(self._do_zero)
        self._zero_lbl = QLabel("")
        self._zero_lbl.setStyleSheet(f"color: {T.GREEN}; font-size: 11px;")
        reset_btn = QPushButton("↺  RESET DEFAULTS")
        reset_btn.setToolTip("Reset all space mouse settings to factory defaults")
        reset_btn.setStyleSheet(
            f"background: {T.SURFACE3}; color: {T.TEXT_DIM}; font-weight: 600;"
            f" border: 1px solid {T.BORDER2}; border-radius: 6px; padding: 7px 14px;"
        )
        reset_btn.clicked.connect(self._reset_sm_defaults)
        zero_row.addWidget(zero_btn)
        zero_row.addWidget(reset_btn)
        zero_row.addWidget(self._zero_lbl)
        zero_row.addStretch()
        layout.addLayout(zero_row)
        layout.addStretch()

        scroll.setWidget(inner)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    def _send_z_mode(self):
        mode = "ZOOM" if self._zm_zoom.isChecked() else "PAN"
        self.serial.send({"action":"set","key":"sm_z_mode","value":mode})

    def _do_zero(self):
        self.serial.send({"action":"zero"})
        self._zero_lbl.setText("Zeroing…")
        self._zero_lbl.setStyleSheet(f"color: {T.YELLOW}; font-size: 11px;")

    def set_zeroed(self, offsets):
        self._zero_lbl.setText(
            f"Zeroed ✓  ({offsets[0]:.1f}, {offsets[1]:.1f}, {offsets[2]:.1f})"
        )
        self._zero_lbl.setStyleSheet(f"color: {T.GREEN}; font-size: 11px;")

    def _reset_sm_defaults(self):
        """Reset all space mouse settings to factory defaults."""
        defaults = {
            "sm_sensitivity": 15.0,  "sm_deadzone": 100.0,
            "sm_z_threshold": 100.0, "sm_filter": 0.25,
            "sm_adapt": 0.003,       "sm_accel": True,
            "sm_accel_curve": 2.0,   "sm_z_mode": "ZOOM",
            "sm_orbit_enter_ms": 40, "sm_orbit_exit_ms": 80,
        }
        for k, v in defaults.items():
            self.serial.send({"action": "set", "key": k, "value": v})
        # Refresh sliders/toggles to match
        self.apply_config(defaults)
        self._zero_lbl.setText("Defaults restored")
        self._zero_lbl.setStyleSheet(f"color: {T.BLUE_LT}; font-size: 11px;")

    def apply_config(self, cfg):
        # Joystick
        if "joy_speed"    in cfg: self._speed.set_value(cfg["joy_speed"])
        if "joy_deadzone" in cfg: self._dz.set_value(cfg["joy_deadzone"])
        if "joy_invert_x" in cfg: self._ix_toggle.setChecked(cfg["joy_invert_x"])
        if "joy_invert_y" in cfg: self._iy_toggle.setChecked(cfg["joy_invert_y"])
        if "joy_sw"       in cfg: self._joy_sw_combo.setCurrentText(cfg["joy_sw"])
        # Space mouse
        sm_mapping = {
            "sm_sensitivity": self._sens, "sm_deadzone": self._sm_dz,
            "sm_z_threshold": self._zt,   "sm_filter":   self._filt,
            "sm_adapt": self._adapt,
            "sm_accel_curve": self._curve, "sm_orbit_enter_ms": self._enter,
            "sm_orbit_exit_ms": self._exit,
        }
        for k, widget in sm_mapping.items():
            if k in cfg: widget.set_value(cfg[k])
        if "sm_accel"  in cfg: self._accel_toggle.setChecked(cfg["sm_accel"])
        if "sm_z_mode" in cfg:
            if cfg["sm_z_mode"] == "PAN":
                self._zm_pan.setChecked(True)
            else:
                self._zm_zoom.setChecked(True)


class ProfilesTab(QWidget):
    mappings_changed = Signal()    # emitted when mappings change so MainWindow can sync

    def __init__(self, serial_mgr, parent=None):
        super().__init__(parent)
        self.serial    = serial_mgr
        self._cfg      = {}
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # ── Live status card ──
        status_card, status_inner = make_card("LIVE STATUS")
        status_row = QHBoxLayout()
        self._fg_label = QLabel("Foreground app:  —")
        self._fg_label.setStyleSheet(
            f"color: {T.TEXT}; font-size: 12px; font-family: 'Consolas', monospace;"
        )
        self._match_label = QLabel("")
        self._match_label.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 11px;"
        )
        status_row.addWidget(self._fg_label)
        status_row.addSpacing(20)
        status_row.addWidget(self._match_label)
        status_row.addStretch()

        self._auto_toggle = ToggleSwitch(True)
        self._auto_toggle.toggled.connect(self._on_auto_toggle)
        auto_lbl = QLabel("Auto-switch")
        auto_lbl.setStyleSheet(f"color: {T.TEXT_MID}; font-size: 11px;")
        status_row.addWidget(auto_lbl)
        status_row.addWidget(self._auto_toggle)
        status_inner.addLayout(status_row)
        layout.addWidget(status_card)

        # ── App mappings card ──
        map_card, map_inner = make_card("APP → LAYER MAPPINGS")
        desc = QLabel(
            "Assign applications to layers by executable name. "
            "When the matched app is in the foreground, the Pico switches to that layer automatically. "
            "Mappings are stored on the Pico and persist across reboots."
        )
        desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px; line-height: 1.6;")
        desc.setWordWrap(True)
        map_inner.addWidget(desc)

        self._rule_list = QListWidget()
        self._rule_list.setMinimumHeight(140)
        map_inner.addWidget(self._rule_list)

        # ── Add rule row ──
        add_row = QHBoxLayout()
        add_row.addWidget(QLabel("Application:"))

        self._app_combo = QComboBox()
        self._app_combo.setFixedWidth(220)
        self._app_combo.setEditable(True)
        self._app_combo.lineEdit().setPlaceholderText("select or type exe name")
        add_row.addWidget(self._app_combo)

        scan_btn = QPushButton("SCAN")
        scan_btn.setToolTip("Scan for currently running applications")
        scan_btn.setFixedWidth(60)
        scan_btn.clicked.connect(self._scan_apps)
        add_row.addWidget(scan_btn)

        add_row.addWidget(QLabel("→ Layer:"))
        self._layer_combo = QComboBox()
        self._layer_combo.setFixedWidth(160)
        self._layer_combo.addItem("0: Default")
        add_row.addWidget(self._layer_combo)

        add_btn    = accent_btn("ADD")
        remove_btn = danger_btn("REMOVE")
        add_btn.clicked.connect(self._add_rule)
        remove_btn.clicked.connect(self._remove_rule)
        add_row.addWidget(add_btn)
        add_row.addWidget(remove_btn)
        add_row.addStretch()
        map_inner.addLayout(add_row)

        # ── Default layer ──
        def_row = QHBoxLayout()
        def_row.addWidget(QLabel("Default layer (no match):"))
        self._default_combo = QComboBox()
        self._default_combo.setFixedWidth(160)
        self._default_combo.addItem("0: Default")
        self._default_combo.currentIndexChanged.connect(self._on_default_changed)
        def_row.addWidget(self._default_combo)
        def_row.addStretch()
        map_inner.addLayout(def_row)

        layout.addWidget(map_card)

        # ── Import/Export card ──
        ie_card, ie_inner = make_card("PROFILE IMPORT / EXPORT")
        ie_desc = QLabel("Import or export the full Pico configuration (all layers, keys, settings, and app mappings) as a JSON file.")
        ie_desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px;")
        ie_desc.setWordWrap(True)
        ie_inner.addWidget(ie_desc)
        ie_row = QHBoxLayout()
        self._import_btn = QPushButton("IMPORT PROFILE")
        self._export_btn = QPushButton("EXPORT PROFILE")
        ie_row.addWidget(self._import_btn)
        ie_row.addWidget(self._export_btn)
        ie_row.addStretch()
        ie_inner.addLayout(ie_row)
        layout.addWidget(ie_card)

        layout.addStretch()

    def _scan_apps(self):
        """Populate the app combo with currently running applications."""
        apps = _list_running_apps()
        current = self._app_combo.currentText()
        self._app_combo.clear()
        self._app_combo.addItems(apps)
        if current:
            self._app_combo.setCurrentText(current)

    def apply_config(self, cfg):
        """Update layer combos and rule list from Pico config."""
        self._cfg = cfg
        layers = cfg.get("layers", [])

        # Update layer combos
        for combo in (self._layer_combo, self._default_combo):
            prev = combo.currentIndex()
            combo.blockSignals(True)
            combo.clear()
            for i, l in enumerate(layers):
                combo.addItem(f"{i}: {l.get('name','')}")
            if 0 <= prev < combo.count():
                combo.setCurrentIndex(prev)
            combo.blockSignals(False)

        # Set default layer
        default_layer = cfg.get("default_layer", 0)
        if 0 <= default_layer < self._default_combo.count():
            self._default_combo.blockSignals(True)
            self._default_combo.setCurrentIndex(default_layer)
            self._default_combo.blockSignals(False)

        # Load mappings from Pico config
        self._refresh_list()

    def _refresh_list(self):
        self._rule_list.clear()
        mappings = self._cfg.get("app_mappings", [])
        layers = self._cfg.get("layers", [])
        for m in mappings:
            exe = m.get("exe", "")
            idx = m.get("layer", 0)
            lname = layers[idx]["name"] if idx < len(layers) else f"#{idx}"
            self._rule_list.addItem(f"  {exe}  →  Layer {idx}: {lname}")

    def _add_rule(self):
        exe = self._app_combo.currentText().strip().lower()
        if not exe:
            return
        layer_idx = self._layer_combo.currentIndex()
        mappings = self._cfg.setdefault("app_mappings", [])
        # Replace existing mapping for same exe
        mappings[:] = [m for m in mappings if m.get("exe","").lower() != exe]
        mappings.append({"exe": exe, "layer": layer_idx})
        self._refresh_list()
        self._app_combo.clearEditText()
        # Send to Pico
        self.serial.send({"action":"set","key":"app_mappings","value":mappings})
        self.mappings_changed.emit()

    def _remove_rule(self):
        row = self._rule_list.currentRow()
        mappings = self._cfg.get("app_mappings", [])
        if 0 <= row < len(mappings):
            mappings.pop(row)
            self._refresh_list()
            self.serial.send({"action":"set","key":"app_mappings","value":mappings})
            self.mappings_changed.emit()

    def _on_default_changed(self, idx):
        self._cfg["default_layer"] = idx
        self.serial.send({"action":"set","key":"default_layer","value":idx})
        self.mappings_changed.emit()

    def _on_auto_toggle(self, val):
        # Emitted to MainWindow to enable/disable AppWatcher
        self.mappings_changed.emit()

    def is_auto_enabled(self):
        return self._auto_toggle.isChecked()

    def update_foreground(self, exe):
        """Called by MainWindow when AppWatcher reports a foreground change."""
        self._fg_label.setText(f"Foreground app:  {exe}")

    def update_match(self, layer_idx, layer_name, exe):
        """Called by MainWindow when a mapping matches."""
        if exe:
            self._match_label.setText(
                f"→ Matched: Layer {layer_idx}: {layer_name}"
            )
            self._match_label.setStyleSheet(
                f"color: {T.GREEN}; font-size: 11px; font-weight: 600;"
            )
        else:
            self._match_label.setText("→ No match (using default layer)")
            self._match_label.setStyleSheet(
                f"color: {T.TEXT_DIM}; font-size: 11px;"
            )

    def get_mappings(self):
        return self._cfg.get("app_mappings", [])

    def get_default_layer(self):
        return self._cfg.get("default_layer", 0)


class VisualiserTab(QWidget):
    passthrough_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cfg            = {}
        self._active_layer_idx = 0
        self._passthrough    = False
        self._build()

    def _build(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner = QWidget()
        inner.setStyleSheet(f"background: {T.SURFACE};")
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # ── Passthrough toggle ─────────────────────────────────
        pt_card, pt_inner = make_card()
        pt_row = QHBoxLayout()
        pt_icon = QLabel("🔍")
        pt_icon.setStyleSheet("font-size: 18px; border: none;")
        pt_title = QLabel("PASSTHROUGH MODE")
        pt_title.setStyleSheet(
            f"color: {T.TEXT}; font-size: 12px; font-weight: 700; border: none;"
        )
        pt_desc = QLabel(
            "While active, key presses are shown here but no HID events are sent to the PC."
            " Encoders, joystick and space mouse are also suppressed."
        )
        pt_desc.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px; border: none;")
        pt_desc.setWordWrap(True)
        self._pt_toggle = ToggleSwitch(False)
        self._pt_toggle.toggled.connect(self._on_passthrough_toggle)
        self._pt_badge = QLabel("● LIVE")
        self._pt_badge.setStyleSheet(
            f"color: {T.GREEN}; font-size: 11px; font-weight: 700; border: none;"
        )
        pt_row.addWidget(pt_icon)
        pt_row.addSpacing(8)
        vv = QVBoxLayout()
        vv.addWidget(pt_title)
        vv.addWidget(pt_desc)
        pt_row.addLayout(vv, 1)
        pt_row.addWidget(self._pt_badge)
        pt_row.addWidget(self._pt_toggle)
        pt_inner.addLayout(pt_row)
        layout.addWidget(pt_card)

        # ── Status bar ─────────────────────────────────────────
        status_frame = QFrame()
        status_frame.setStyleSheet(
            f"background: {T.SURFACE3}; border-radius: 8px; border: 1px solid {T.BORDER};"
        )
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(14, 8, 14, 8)

        def status_item(prefix_text, attr_name):
            lbl_pre = QLabel(prefix_text)
            lbl_pre.setStyleSheet(
                f"color: {T.TEXT_DIM}; font-size: 10px; font-weight: 700;"
                f" letter-spacing: 1px; border: none; background: transparent;"
            )
            lbl_val = QLabel("—")
            lbl_val.setStyleSheet(
                f"color: {T.BLUE_LT}; font-size: 11px; font-weight: 700;"
                f" border: none; background: transparent;"
            )
            setattr(self, attr_name, lbl_val)
            status_layout.addWidget(lbl_pre)
            status_layout.addWidget(lbl_val)
            status_layout.addSpacing(20)

        status_item("LAYER:", "_vis_layer_lbl")
        status_item("ENC1:",  "_vis_enc1_lbl")
        status_item("ENC2:",  "_vis_enc2_lbl")
        status_layout.addStretch()
        self._zoom_ind = QLabel("⚡ ENC2 ZOOM")
        self._zoom_ind.setStyleSheet(
            f"color: {T.SURFACE3}; font-size: 10px; font-weight: 700;"
            f" border: none; background: transparent;"
        )
        status_layout.addWidget(self._zoom_ind)
        layout.addWidget(status_frame)

        # ── Key grid ───────────────────────────────────────────
        keys_card, keys_inner = make_card("LIVE KEY PRESS")
        grid = QGridLayout()
        grid.setSpacing(4)
        self._vis_btns = []
        for i in range(25):
            r, c = divmod(i, 5)
            if i in GHOST_KEYS:
                btn = QLabel("")
                btn.setFixedSize(80, 50)
                btn.setStyleSheet(
                    f"background: {T.BG}; border: 1px dashed {T.BORDER};"
                    f" border-radius: 6px;"
                )
            else:
                btn = QLabel("·")
                btn.setFixedSize(80, 50)
                btn.setAlignment(Qt.AlignCenter)
                btn.setStyleSheet(
                    f"background: {T.SURFACE3}; color: {T.BORDER2};"
                    f" border-radius: 6px; font-size: 9px;"
                    f" font-family: 'Consolas', monospace; font-weight: 600;"
                )
            grid.addWidget(btn, r, c)
            self._vis_btns.append(btn)
        keys_inner.addLayout(grid)
        layout.addWidget(keys_card)

        # ── Space mouse bars ───────────────────────────────────
        sm_card, sm_inner = make_card("SPACE MOUSE  (live)")
        sm_row = QHBoxLayout()
        sm_row.setSpacing(20)
        self._sm_bars = {}
        self._sm_lbls = {}
        colors = {"X": T.BLUE, "Y": T.CYAN, "Z": T.GREEN}
        for axis in ("X", "Y", "Z"):
            col_w = QWidget()
            col_l = QVBoxLayout(col_w)
            col_l.setAlignment(Qt.AlignHCenter)
            col_l.setSpacing(4)
            axis_lbl = QLabel(axis)
            axis_lbl.setAlignment(Qt.AlignCenter)
            axis_lbl.setStyleSheet(f"color: {T.TEXT}; font-size: 11px; font-weight: 700;")
            bar = BarMeter(colors[axis])
            val_lbl = QLabel("0")
            val_lbl.setAlignment(Qt.AlignCenter)
            val_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 10px; font-family: 'Consolas';")
            col_l.addWidget(axis_lbl)
            col_l.addWidget(bar)
            col_l.addWidget(val_lbl)
            self._sm_bars[axis] = bar
            self._sm_lbls[axis] = val_lbl
            sm_row.addWidget(col_w)

        sm_row.addSpacing(20)
        state_col = QWidget()
        state_l = QVBoxLayout(state_col)
        state_l.setAlignment(Qt.AlignTop)
        self._orbit_ind = QLabel("● ORBIT")
        self._orbit_ind.setStyleSheet(f"color: {T.BORDER2}; font-size: 11px; font-weight: 700;")
        self._pan_ind = QLabel("● PAN")
        self._pan_ind.setStyleSheet(f"color: {T.BORDER2}; font-size: 11px; font-weight: 700;")
        state_l.addWidget(self._orbit_ind)
        state_l.addSpacing(8)
        state_l.addWidget(self._pan_ind)
        sm_row.addWidget(state_col)
        sm_row.addStretch()
        sm_inner.addLayout(sm_row)
        layout.addWidget(sm_card)

        # ── Joystick ───────────────────────────────────────────
        joy_card, joy_inner = make_card("JOYSTICK  (live)")
        joy_row = QHBoxLayout()
        self._joy_view = JoystickView()
        joy_row.addWidget(self._joy_view)
        joy_vals = QWidget()
        joy_vals_l = QVBoxLayout(joy_vals)
        joy_vals_l.setSpacing(4)
        for axis in ("X", "Y"):
            row = QHBoxLayout()
            row.addWidget(QLabel(f"{axis}:"))
            vl = QLabel("0")
            vl.setStyleSheet(f"color: {T.TEXT}; font-family: 'Consolas'; font-size: 11px;")
            setattr(self, f"_joy_{axis.lower()}_lbl", vl)
            row.addWidget(vl)
            row.addStretch()
            joy_vals_l.addLayout(row)
        joy_row.addWidget(joy_vals)
        joy_row.addStretch()
        joy_inner.addLayout(joy_row)
        layout.addWidget(joy_card)

        # ── Boot health ─────────────────────────────────────────
        health_card, health_inner = make_card("BOOT HEALTH")
        self._health_lbl = QLabel("—  connect to Pico to see boot health")
        self._health_lbl.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 11px; font-family: 'Consolas';"
        )
        health_inner.addWidget(self._health_lbl)
        layout.addWidget(health_card)
        layout.addStretch()

        scroll.setWidget(inner)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    def _on_passthrough_toggle(self, val):
        self._passthrough = val
        self._pt_badge.setText("⏸ PASSTHROUGH" if val else "● LIVE")
        self._pt_badge.setStyleSheet(
            f"color: {T.YELLOW if val else T.GREEN}; font-size: 11px;"
            f" font-weight: 700; border: none;"
        )
        self.passthrough_changed.emit(val)

    def apply_config(self, cfg):
        self._cfg = cfg
        self._refresh_key_labels()

    def set_layer(self, idx, name, enc1_mode="", enc2_mode=""):
        self._active_layer_idx = idx
        self._vis_layer_lbl.setText(f"[{idx}] {name}")
        self._vis_enc1_lbl.setText(enc1_mode or "—")
        self._vis_enc2_lbl.setText(enc2_mode or "—")
        # Refresh key labels for new layer
        self._refresh_key_labels()

    def _refresh_key_labels(self):
        """Update resting key labels to reflect current layer bindings."""
        layers = self._cfg.get("layers", [])
        idx = self._active_layer_idx
        if not layers or idx >= len(layers):
            return
        keys = layers[idx].get("keys", [])
        for i, btn in enumerate(self._vis_btns):
            if i in GHOST_KEYS:
                continue
            kc = keys[i] if i < len(keys) else {}
            tap = kc.get("tap", [])
            label = "+".join(tap) if tap else "·"
            color = T.TEXT_DIM if tap else T.BORDER2
            btn.setText(label)
            btn.setStyleSheet(
                f"background: {T.SURFACE3}; color: {color};"
                f" border-radius: 6px; font-size: 9px;"
                f" font-family: 'Consolas', monospace; font-weight: 600;"
            )

    def key_press(self, idx):
        if idx in GHOST_KEYS or idx >= len(self._vis_btns):
            return
        layers = self._cfg.get("layers", [])
        al = self._active_layer_idx
        tap_text = "·"
        if layers and al < len(layers):
            kc = layers[al]["keys"][idx]
            tap = kc.get("tap", [])
            kt  = kc.get("key_type", "normal")
            if kt == "mo":
                tap_text = f"MO{kc.get('mo_layer', 0)}"
            elif kt == "mouse_hold":
                tap_text = f"🖱{kc.get('mouse_button','?')[:1]}"
            elif tap:
                tap_text = "+".join(tap)
            else:
                tap_text = "·"
        btn = self._vis_btns[idx]
        btn.setText(tap_text)
        btn.setStyleSheet(
            f"background: {T.BLUE}; color: white;"
            f" border-radius: 6px; font-size: 9px;"
            f" font-family: 'Consolas', monospace; font-weight: 600;"
        )

    def key_release(self, idx):
        if idx in GHOST_KEYS or idx >= len(self._vis_btns):
            return
        # Restore resting state label
        layers = self._cfg.get("layers", [])
        al = self._active_layer_idx
        tap_text = "·"
        color = T.BORDER2
        if layers and al < len(layers):
            kc = layers[al]["keys"][idx]
            tap = kc.get("tap", [])
            if tap:
                tap_text = "+".join(tap)
                color = T.TEXT_DIM
        btn = self._vis_btns[idx]
        btn.setText(tap_text)
        btn.setStyleSheet(
            f"background: {T.SURFACE3}; color: {color};"
            f" border-radius: 6px; font-size: 9px;"
            f" font-family: 'Consolas', monospace; font-weight: 600;"
        )

    def update_sm(self, x, y, z):
        for axis, val in (("X", x), ("Y", y), ("Z", z)):
            self._sm_bars[axis].set_value(val)
            self._sm_lbls[axis].setText(f"{val:.0f}")
        dz = self._cfg.get("sm_deadzone", 100.0)
        zt = self._cfg.get("sm_z_threshold", 100.0)
        orbiting = abs(x) > dz or abs(y) > dz
        z_active = abs(z) > zt
        self._orbit_ind.setStyleSheet(
            f"color: {T.BLUE_LT if orbiting else T.BORDER2}; font-size: 11px; font-weight: 700;"
        )
        self._pan_ind.setStyleSheet(
            f"color: {T.CYAN if z_active else T.BORDER2}; font-size: 11px; font-weight: 700;"
        )
        for axis in ("X", "Y"):
            self._sm_bars[axis].set_deadzone(dz)
        self._sm_bars["Z"].set_deadzone(zt)

    def update_joy(self, x, y):
        self._joy_view.set_position(x, y)
        self._joy_x_lbl.setText(str(int(x)))
        self._joy_y_lbl.setText(str(int(y)))

    def set_zoom_override(self, active):
        self._zoom_ind.setStyleSheet(
            f"color: {T.YELLOW if active else T.SURFACE3};"
            f" font-size: 10px; font-weight: 700; border: none; background: transparent;"
        )

    def show_health(self, health):
        parts = [
            f"MLX: {'✓' if health.get('mlx') else '✗'}",
            f"OLED: {'✓' if health.get('oled') else '✗'}",
            f"Settings: {health.get('settings', '?')}",
            f"Enc1@boot: {health.get('enc1_pos', 0)}",
            f"Enc2@boot: {health.get('enc2_pos', 0)}",
            f"Layers: {health.get('layers', 0)}",
        ]
        self._health_lbl.setText("  |  ".join(parts))


# ─────────────────────────────────────────────────────────────
#  10. MAIN WINDOW
# ─────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SpacePad Configurator")
        self.setMinimumSize(860, 760)
        self._cfg = {}
        self._unsaved = False
        self._really_quit = False

        self._settings = QSettings("SpacePad", "Configurator")

        self.serial = SerialManager(self)
        self.serial.message_received.connect(self._on_message)
        self.serial.connected.connect(self._on_connected)
        self.serial.disconnected.connect(self._on_disconnected)

        # App watcher (foreground process monitor)
        self._app_watcher = AppWatcher(self)
        self._app_watcher.layer_requested.connect(self._on_app_layer_requested)
        self._app_watcher.foreground_changed.connect(self._on_foreground_changed)
        self._app_watcher.no_match.connect(self._on_no_match)

        self._build_ui()
        self._build_tray_icon()
        self._refresh_ports()
        self._app_watcher.start()

        # Reconnect watchdog timer handled by SerialManager
        self._save_flash_timer = QTimer(self)
        self._save_flash_timer.setSingleShot(True)
        self._save_flash_timer.timeout.connect(self._reset_save_btn)

        # Keyboard shortcuts
        QShortcut(QKeySequence("Ctrl+S"), self, self._save_to_pico)

        # Auto-connect to last known port
        QTimer.singleShot(500, self._try_auto_connect)

    def _build_tray_icon(self):
        """Create system tray icon with right-click menu."""
        self._tray = QSystemTrayIcon(self)
        self._tray.setIcon(self._make_tray_icon())
        self._tray.setToolTip("SpacePad — Disconnected")
        self._tray.activated.connect(self._on_tray_activated)

        tray_menu = QMenu()
        self._tray_open_action = QAction("Open SpacePad", self)
        self._tray_open_action.triggered.connect(self._show_window)
        tray_menu.addAction(self._tray_open_action)

        tray_menu.addSeparator()

        self._tray_layer_label = QAction("Layer: —", self)
        self._tray_layer_label.setEnabled(False)
        tray_menu.addAction(self._tray_layer_label)

        self._tray_app_label = QAction("App: —", self)
        self._tray_app_label.setEnabled(False)
        tray_menu.addAction(self._tray_app_label)

        tray_menu.addSeparator()

        self._tray_auto_action = QAction("Auto-switching enabled", self)
        self._tray_auto_action.setCheckable(True)
        self._tray_auto_action.setChecked(True)
        self._tray_auto_action.triggered.connect(self._on_tray_auto_toggle)
        tray_menu.addAction(self._tray_auto_action)

        self._tray_minimized_action = QAction("Launch minimized to tray", self)
        self._tray_minimized_action.setCheckable(True)
        self._tray_minimized_action.setChecked(
            self._settings.value("launch_minimized", False, type=bool)
        )
        self._tray_minimized_action.triggered.connect(self._on_launch_minimized_toggle)
        tray_menu.addAction(self._tray_minimized_action)

        self._tray_startup_action = QAction("Start with Windows", self)
        self._tray_startup_action.setCheckable(True)
        self._tray_startup_action.setChecked(self._is_startup_enabled())
        self._tray_startup_action.triggered.connect(self._on_startup_toggle)
        tray_menu.addAction(self._tray_startup_action)

        tray_menu.addSeparator()

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self._real_quit)
        tray_menu.addAction(quit_action)

        self._tray.setContextMenu(tray_menu)
        self._tray.show()

    @staticmethod
    def _make_tray_icon(connected=False):
        """Generate a small tray icon — blue grid when idle, green when connected."""
        img = QImage(64, 64, QImage.Format_ARGB32)
        img.fill(QColor(0, 0, 0, 0))
        p = QPainter(img)
        p.setRenderHint(QPainter.Antialiasing)
        color = QColor(T.GREEN) if connected else QColor(T.BLUE)
        for r in range(3):
            for c in range(3):
                p.setBrush(color)
                p.setPen(Qt.NoPen)
                p.drawRoundedRect(6 + c * 19, 6 + r * 19, 14, 14, 3, 3)
        p.end()
        return QIcon(QPixmap.fromImage(img))

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Title bar ───────────────────────────────────────
        title_bar = QFrame()
        title_bar.setStyleSheet(
            f"background: {T.BG}; border-bottom: 1px solid {T.BORDER};"
        )
        title_bar.setFixedHeight(58)
        tb_layout = QHBoxLayout(title_bar)
        tb_layout.setContentsMargins(20, 0, 20, 0)

        brand = QLabel()
        brand.setStyleSheet("border: none;")
        brand_html = (
            f'<span style="color:{T.BLUE_LT};font-size:18px;font-weight:800;">'
            f'SpacePad</span>'
            f'<span style="color:{T.TEXT};font-size:16px;"> Configurator</span>'
        )
        brand.setText(brand_html)
        tb_layout.addWidget(brand)
        tb_layout.addStretch()

        self._status_dot = QLabel("●")
        self._status_dot.setStyleSheet(f"color: {T.RED}; font-size: 16px; border: none;")
        self._status_lbl = QLabel("Disconnected")
        self._status_lbl.setStyleSheet(f"color: {T.TEXT_DIM}; font-size: 11px; border: none;")
        tb_layout.addWidget(self._status_lbl)
        tb_layout.addWidget(self._status_dot)
        root.addWidget(title_bar)

        # ── Connection bar ──────────────────────────────────
        conn_bar = QFrame()
        conn_bar.setStyleSheet(
            f"background: {T.SURFACE}; border-bottom: 1px solid {T.BORDER};"
        )
        conn_bar.setFixedHeight(46)
        cb_layout = QHBoxLayout(conn_bar)
        cb_layout.setContentsMargins(16, 0, 16, 0)
        cb_layout.setSpacing(8)

        port_lbl = QLabel("PORT")
        port_lbl.setStyleSheet(
            f"color: {T.TEXT_DIM}; font-size: 10px; font-weight: 700;"
            f" letter-spacing: 1px; border: none;"
        )
        self._port_combo = QComboBox()
        self._port_combo.setFixedWidth(180)
        refresh_btn = QPushButton("↻")
        refresh_btn.setFixedWidth(34)
        refresh_btn.clicked.connect(self._refresh_ports)
        self._conn_btn = QPushButton("CONNECT")
        self._conn_btn.setProperty("accent", True)
        self._conn_btn.setFixedWidth(110)
        self._conn_btn.clicked.connect(self._toggle_connect)

        cb_layout.addWidget(port_lbl)
        cb_layout.addWidget(self._port_combo)
        cb_layout.addWidget(refresh_btn)
        cb_layout.addWidget(self._conn_btn)
        cb_layout.addStretch()

        self._save_btn   = QPushButton("SAVE TO PICO")
        self._save_btn.setStyleSheet(
            f"background: {T.BLUE_DIM}; color: {T.BLUE_LT}; border: 1px solid {T.BLUE_DIM};"
            f" border-radius: 6px; padding: 5px 16px; font-weight: 700; font-size: 11px;"
        )
        self._save_btn.clicked.connect(self._save_to_pico)

        cb_layout.addWidget(self._save_btn)
        root.addWidget(conn_bar)

        # ── Tabs ────────────────────────────────────────────
        self._tabs = QTabWidget()
        self._tabs.setDocumentMode(True)

        self._tab_matrix  = MatrixTab(self.serial)
        self._tab_layers  = LayersTab(self.serial)
        self._tab_input   = InputTab(self.serial)
        self._tab_profiles= ProfilesTab(self.serial)
        self._tab_profiles.mappings_changed.connect(self._sync_app_watcher)
        self._tab_profiles._import_btn.clicked.connect(self._import_profile)
        self._tab_profiles._export_btn.clicked.connect(self._export_profile)
        self._tab_vis     = VisualiserTab()
        self._tab_vis.passthrough_changed.connect(self._on_passthrough_changed)

        tab_defs = [
            ("  Matrix  ",       self._tab_matrix),
            ("  Layers  ",       self._tab_layers),
            ("  Input  ",        self._tab_input),
            ("  Profiles  ",     self._tab_profiles),
            ("  Visualiser  ",   self._tab_vis),
        ]
        for name, widget in tab_defs:
            self._tabs.addTab(widget, name)

        # Wrap tabs in a container with background
        tab_container = QWidget()
        tab_container.setStyleSheet(f"background: {T.SURFACE};")
        tc_layout = QVBoxLayout(tab_container)
        tc_layout.setContentsMargins(0, 0, 0, 0)
        tc_layout.addWidget(self._tabs)
        root.addWidget(tab_container, 1)

        # ── Status bar ──────────────────────────────────────
        self.statusBar().showMessage("—")

    # ── Connections ──────────────────────────────────────────

    def _refresh_ports(self):
        ports = self.serial.list_ports()
        current = self._port_combo.currentText()
        self._port_combo.clear()
        self._port_combo.addItems(ports)
        if current in ports:
            self._port_combo.setCurrentText(current)

    def _try_auto_connect(self):
        last = self._settings.value("last_port", "")
        if last and not self.serial.is_connected():
            ports = self.serial.list_ports()
            if last in ports:
                self._port_combo.setCurrentText(last)
                self._toggle_connect()
                self._log(f"Auto-connecting to {last}…")

    def _toggle_connect(self):
        if self.serial.is_connected():
            self.serial.send({"action":"unsubscribe"})
            self.serial.disconnect_port()
        else:
            port = self._port_combo.currentText()
            if not port:
                self._log("No port selected.")
                return
            self._settings.setValue("last_port", port)
            self.serial.connect_port(port)
            QTimer.singleShot(400, lambda: (
                self.serial.send({"action":"subscribe"}),
                self.serial.send({"action":"get_config"}),
            ))

    def _on_connected(self, port):
        self._status_dot.setStyleSheet(f"color: {T.GREEN}; font-size: 16px; border: none;")
        self._status_lbl.setText(f"Connected  {port}")
        self._conn_btn.setText("DISCONNECT")
        self._conn_btn.setStyleSheet(
            f"background: {T.RED_DIM}; color: {T.RED}; border: 1px solid {T.RED_DIM};"
            f" border-radius: 6px; padding: 5px 14px; font-weight: 700;"
        )
        self._tray.setIcon(self._make_tray_icon(connected=True))
        self._tray.setToolTip(f"SpacePad — Connected ({port})")
        self._log(f"Connected to {port}")

    def _on_disconnected(self):
        if self._tab_vis._passthrough:
            self._tab_vis._pt_toggle.setChecked(False)
        self._tab_input._zero_lbl.setText("")
        self._status_dot.setStyleSheet(f"color: {T.RED}; font-size: 16px; border: none;")
        self._status_lbl.setText("Disconnected")
        self._conn_btn.setText("CONNECT")
        self._conn_btn.setProperty("accent", True)
        self._conn_btn.setStyleSheet("")
        self._tray.setIcon(self._make_tray_icon(connected=False))
        self._tray.setToolTip("SpacePad — Disconnected")
        self._log("Disconnected — watching for reconnect…")

    # ── Tray icon ────────────────────────────────────────────

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._show_window()

    def _show_window(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def _on_tray_auto_toggle(self, checked):
        self._app_watcher.set_enabled(checked)
        self._tab_profiles._auto_toggle.setChecked(checked)

    def _on_launch_minimized_toggle(self, checked):
        self._settings.setValue("launch_minimized", checked)

    _STARTUP_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
    _STARTUP_APP_NAME = "SpacePad"

    def _get_startup_command(self):
        """Build the startup command for the current execution mode."""
        import os
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller .exe
            return f'"{sys.executable}" --minimized'
        else:
            # Running as Python script
            script = os.path.abspath(sys.argv[0])
            return f'"{sys.executable}" "{script}" --minimized'

    def _is_startup_enabled(self):
        """Check if SpacePad is in the Windows startup registry."""
        if sys.platform != "win32":
            return False
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self._STARTUP_REG_KEY, 0, winreg.KEY_READ)
            try:
                winreg.QueryValueEx(key, self._STARTUP_APP_NAME)
                return True
            except FileNotFoundError:
                return False
            finally:
                winreg.CloseKey(key)
        except Exception:
            return False

    def _on_startup_toggle(self, checked):
        """Add or remove SpacePad from Windows startup."""
        if sys.platform != "win32":
            return
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self._STARTUP_REG_KEY, 0, winreg.KEY_WRITE)
            try:
                if checked:
                    winreg.SetValueEx(key, self._STARTUP_APP_NAME, 0, winreg.REG_SZ,
                                      self._get_startup_command())
                    self._settings.setValue("launch_minimized", True)
                    self._tray_minimized_action.setChecked(True)
                    self._log("SpacePad added to Windows startup (minimized)")
                else:
                    try:
                        winreg.DeleteValue(key, self._STARTUP_APP_NAME)
                    except FileNotFoundError:
                        pass
                    self._log("SpacePad removed from Windows startup")
            finally:
                winreg.CloseKey(key)
        except Exception as e:
            self._log(f"Startup toggle failed: {e}")
            self._tray_startup_action.setChecked(not checked)

    def _real_quit(self):
        if self._unsaved:
            ret = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes that will be lost if you quit.\n\n"
                "Save to Pico before quitting?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            )
            if ret == QMessageBox.Save:
                self._save_to_pico()
                QTimer.singleShot(1000, self._force_quit)
                return
            elif ret == QMessageBox.Cancel:
                return
        self._force_quit()

    def _force_quit(self):
        self._really_quit = True
        self._app_watcher.stop()
        self._app_watcher.wait(1000)
        self.serial.disconnect_port()
        self._tray.hide()
        QApplication.instance().quit()

    def closeEvent(self, event):
        if self._really_quit:
            event.accept()
            return
        # Minimize to tray instead of closing
        event.ignore()
        self.hide()
        self._tray.showMessage(
            "SpacePad",
            "Running in the background. Right-click the tray icon to quit.",
            QSystemTrayIcon.Information,
            2000,
        )

    # ── App watcher integration ──────────────────────────────

    def _sync_app_watcher(self):
        """Sync app watcher with current mappings from Profiles tab."""
        mappings = self._tab_profiles.get_mappings()
        default  = self._tab_profiles.get_default_layer()
        enabled  = self._tab_profiles.is_auto_enabled()
        self._app_watcher.set_mappings(mappings, default)
        self._app_watcher.set_enabled(enabled)
        self._tray_auto_action.setChecked(enabled)
        self._mark_unsaved()

    @Slot(int, str)
    def _on_app_layer_requested(self, layer_idx, exe):
        """AppWatcher wants to switch layers."""
        if self.serial.is_connected():
            self.serial.send({"action":"set_active_layer","index":layer_idx})
            layers = self._cfg.get("layers", [])
            lname = layers[layer_idx]["name"] if layer_idx < len(layers) else f"#{layer_idx}"
            self._tray_layer_label.setText(f"Layer: {layer_idx}: {lname}")
            self._tray_app_label.setText(f"App: {exe}" if exe else "App: (default)")
            self._tray.setToolTip(f"SpacePad — Layer {layer_idx}: {lname}" + (f" ({exe})" if exe else ""))
            self._tab_profiles.update_match(layer_idx, lname, exe)

    @Slot(str)
    def _on_foreground_changed(self, exe):
        self._tab_profiles.update_foreground(exe)
        self._tray_app_label.setText(f"App: {exe}")

    @Slot()
    def _on_no_match(self):
        self._tab_profiles.update_match(0, "", "")

    # ── Unsaved changes tracking ─────────────────────────────

    def _mark_unsaved(self):
        if not self._unsaved:
            self._unsaved = True
            self._update_title()

    def _mark_saved(self):
        self._unsaved = False
        self._update_title()

    def _update_title(self):
        layer = self._cfg.get("active_layer", 0)
        layers = self._cfg.get("layers", [])
        lname = layers[layer]["name"] if layer < len(layers) else ""
        dot = "  ●" if self._unsaved else ""
        layer_part = f"  —  [{layer}] {lname}" if lname else ""
        self.setWindowTitle(f"SpacePad Configurator{layer_part}{dot}")

    # ── Message handler ───────────────────────────────────────

    @Slot(dict)
    def _on_message(self, msg):
        ev = msg.get("event","")

        if ev == "config":
            self._cfg = msg.get("data",{})
            self._apply_config()

        elif ev in ("ack_layers","ack_add_layer","ack_remove_layer","ack_rename_layer"):
            self.serial.send({"action":"get_config"})

        elif ev == "layer_changed":
            idx  = msg.get("index",0)
            name = msg.get("name","")
            self._on_layer_changed(idx, name)

        elif ev == "key_press":
            idx = msg.get("index")
            if idx is not None:
                self._tab_matrix.flash_key(idx)
                self._tab_vis.key_press(idx)

        elif ev == "key_release":
            idx = msg.get("index")
            if idx is not None:
                self._tab_vis.key_release(idx)

        elif ev == "sm_data":
            self._tab_vis.update_sm(
                msg.get("x",0), msg.get("y",0), msg.get("z",0)
            )

        elif ev == "joy_data":
            self._tab_vis.update_joy(msg.get("x",0), msg.get("y",0))

        elif ev == "enc2_zoom_override":
            self._tab_vis.set_zoom_override(msg.get("active",False))

        elif ev == "zeroed":
            self._tab_input.set_zeroed(msg.get("offsets",[0,0,0]))
            self._log("Space mouse zeroed.")

        elif ev == "passthrough_on":
            self._log("⏸ Passthrough confirmed — HID suppressed on Pico")

        elif ev == "passthrough_off":
            self._log("▶ Passthrough off — HID active")

        elif ev == "saved":
            self._save_btn.setText("SAVED  ✓")
            self._save_btn.setStyleSheet(
                f"background: {T.GREEN_DIM}; color: {T.GREEN}; border: 1px solid {T.GREEN_DIM};"
                f" border-radius: 6px; padding: 5px 16px; font-weight: 700; font-size: 11px;"
            )
            self._save_flash_timer.start(2000)
            self._mark_saved()
            self._log("✓ Saved to Pico flash.")

        elif ev == "save_failed":
            self._log("✗ Save failed — check Pico filesystem (boot.py needed)")
            QMessageBox.warning(
                self, "Save Failed",
                "The Pico could not write settings.json.\n\n"
                "You need a boot.py on the Pico root containing:\n\n"
                "  import storage\n"
                "  storage.remount('/', readonly=False)\n\n"
                "Copy boot.py to CIRCUITPY and reboot the Pico."
            )

        elif ev == "save_error":
            detail = msg.get("detail", "unknown error")
            self._log(f"✗ Save error: {detail}")
            QMessageBox.warning(self, "Save Error", f"Save failed:\n\n{detail}")

        elif ev == "boot_complete":
            self._tab_vis.show_health(msg.get("health",{}))
            self._log(f"Pico boot complete — {msg.get('health',{}).get('settings','?')}")

        elif ev == "joy_calibrated":
            cx, cy = msg.get("cx", 32768), msg.get("cy", 32768)
            self._log(f"Joystick calibrated — centre X:{cx}  Y:{cy}")

        elif ev == "ack":
            self._mark_unsaved()
            self._log(f"✓ {msg.get('key')} = {msg.get('value')}")

        elif ev == "ack_layer_prop":
            # Silent — no log noise for sm_active toggle changes
            pass

        elif "error" in msg:
            self._log(f"✗ {msg}")

    def _apply_config(self, from_boot=False):
        self._tab_matrix.apply_config(self._cfg)
        self._tab_layers.apply_config(self._cfg)
        self._tab_input.apply_config(self._cfg)
        self._tab_profiles.apply_config(self._cfg)
        self._tab_vis.apply_config(self._cfg)
        self._sync_app_watcher()
        self._mark_saved()
        self._log("Config loaded from Pico.")

    def _on_layer_changed(self, idx, name):
        self._cfg["active_layer"] = idx
        layers = self._cfg.get("layers",[])
        enc1 = enc2 = ""
        sm_active = False
        if layers and idx < len(layers):
            l = layers[idx]
            enc1 = ENCODER_MODES.get(l.get("enc1_mode",""),"")
            enc2 = ENCODER_MODES.get(l.get("enc2_mode",""),"")
            sm_active = l.get("sm_active", False)
        self._tab_matrix.set_active_layer(idx)
        self._tab_layers.set_active(idx, name)
        self._tab_layers.set_sm_active(sm_active)   # keep sm toggle in sync
        self._tab_vis.set_layer(idx, name, enc1, enc2)
        self._update_title()

    def _on_passthrough_changed(self, val):
        action = "passthrough_on" if val else "passthrough_off"
        self.serial.send({"action": action})
        self._log(f"{'⏸ Passthrough ON — HID suppressed' if val else '▶ Passthrough OFF — HID active'}")

    def _reset_save_btn(self):
        self._save_btn.setText("SAVE TO PICO")
        self._save_btn.setStyleSheet(
            f"background: {T.BLUE_DIM}; color: {T.BLUE_LT}; border: 1px solid {T.BLUE_DIM};"
            f" border-radius: 6px; padding: 5px 16px; font-weight: 700; font-size: 11px;"
        )

    # ── Profile import/export ─────────────────────────────────

    def _export_profile(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Profile", "", "JSON Profile (*.json);;All Files (*)"
        )
        if path:
            try:
                with open(path,"w") as f:
                    json.dump(self._cfg, f, indent=2)
                self._log(f"Profile exported to {path}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", str(e))

    def _import_profile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Profile", "", "JSON Profile (*.json);;All Files (*)"
        )
        if not path: return
        try:
            with open(path) as f:
                data = json.load(f)
            ok, err = validate_config(data)
            if not ok:
                QMessageBox.critical(
                    self, "Invalid Profile",
                    f"This profile cannot be imported:\n\n{err}"
                )
                return
            if "layers" in data:
                self.serial.send({"action":"set_layers","layers":data["layers"]})
            for k, v in data.items():
                if k not in ("layers","active_layer"):
                    self.serial.send({"action":"set","key":k,"value":v})
            self.serial.send({"action":"save"})
            self._log(f"Profile imported from {path}")
        except Exception as e:
            QMessageBox.critical(self, "Import Failed", str(e))

    def _save_to_pico(self):
        self.serial.send({"action":"save"})
        self._log("Saving to Pico flash…")

    def _log(self, text):
        ts = time.strftime("%H:%M:%S")
        self.statusBar().showMessage(f"[{ts}]  {text}")


# ─────────────────────────────────────────────────────────────
#  11. ENTRY POINT
# ─────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("SpacePad Configurator")
    app.setOrganizationName("SpacePad")
    app.setStyleSheet(QSS)

    # Prevent quitting when last window is hidden to tray
    app.setQuitOnLastWindowClosed(False)

    window = MainWindow()

    # Check for --minimized flag or saved preference
    settings = QSettings("SpacePad", "Configurator")
    launch_minimized = "--minimized" in sys.argv or settings.value("launch_minimized", False, type=bool)

    if launch_minimized:
        # Start hidden — tray icon is already visible
        window.hide()
    else:
        window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
