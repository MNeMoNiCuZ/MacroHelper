import copy
import json
import os
import random
import re
import sys
import threading
import time
from typing import Dict, List, Optional

import pyttsx3

if sys.platform.startswith("win"):
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=0")
    os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.window=false;qt.qpa.windows=false")

from PyQt6.QtCore import Qt, QTimer, QPoint, QSize, pyqtSignal
from PyQt6.QtGui import QAction, QIcon, QColor, QCursor
from PyQt6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QFrame, QPushButton, QTabBar, QScrollArea,
    QTableWidget, QTableWidgetItem, QLineEdit, QSpinBox, QSlider, QLabel, QCheckBox, QComboBox, QSystemTrayIcon,
    QMenu, QMessageBox, QGridLayout, QSizePolicy, QStyle, QColorDialog
)


def rp(path: str) -> str:
    return os.path.join(getattr(sys, "_MEIPASS", os.path.abspath(".")), path)


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_speech_lock = threading.Lock()


def speak(text: str, vol: float, rate: int, voice_id: Optional[str]):
    def _run():
        try:
            with _speech_lock:
                e = pyttsx3.init()
                e.setProperty("volume", vol)
                e.setProperty("rate", max(100, min(1000, int(rate))))
                if voice_id:
                    e.setProperty("voice", voice_id)
                e.say(text)
                e.runAndWait()
        except Exception:
            pass
    threading.Thread(target=_run, daemon=True).start()

def voices() -> List[dict]:
    out = []
    try:
        e = pyttsx3.init()
        for v in e.getProperty("voices") or []:
            out.append({"id": v.id, "name": getattr(v, "name", v.id)})
    except Exception:
        pass
    return out


def _screen_for_window_pos(pos: QPoint, w: int, h: int):
    probe = QPoint(pos.x() + max(1, w // 2), pos.y() + max(1, h // 2))
    screen = QApplication.screenAt(probe) or QApplication.screenAt(pos) or QApplication.primaryScreen()
    return screen


def clamp_window_pos(pos: QPoint, w: int, h: int) -> QPoint:
    screen = _screen_for_window_pos(pos, w, h)
    if not screen:
        return pos
    g = screen.availableGeometry()
    max_x = max(g.left(), g.right() - w + 1)
    max_y = max(g.top(), g.bottom() - h + 1)
    x = min(max(pos.x(), g.left()), max_x)
    y = min(max(pos.y(), g.top()), max_y)
    return QPoint(x, y)


def fit_window_to_screen(win: QWidget, preferred_pos: Optional[QPoint] = None):
    p = preferred_pos if preferred_pos is not None else win.pos()
    screen = _screen_for_window_pos(p, win.width(), win.height())
    if not screen:
        return
    g = screen.availableGeometry()
    nw = min(win.width(), g.width())
    nh = min(win.height(), g.height())
    if nw != win.width() or nh != win.height():
        win.resize(nw, nh)
    win.move(clamp_window_pos(p, win.width(), win.height()))


def parse_time_expr(text: str):
    parts = [p.strip() for p in re.split(r"[,.]+", str(text)) if p.strip() != ""]
    vals = [int(x) for x in parts]
    if len(vals) < 1:
        raise ValueError("invalid time expr")
    if vals[-1] == 0:
        if len(vals) < 2 or vals[-2] <= 0 or any(v < 0 for v in vals[:-1]):
            raise ValueError("invalid terminal time expr")
        return vals[:-1], vals[-2], True
    if vals[-1] <= 0 or any(v < 0 for v in vals[:-1]):
        raise ValueError("invalid time expr")
    return vals[:-1], vals[-1], False


def parse_text_seq(text: str) -> List[str]:
    s = str(text or "").strip()
    if not s:
        return []
    out = [p.strip() for p in re.split(r"[,.]+", s) if p.strip() != ""]
    return out if out else [s]


def _fmt_seconds(s: int) -> str:
    m, sec = divmod(int(s), 60)
    if m > 0 and sec > 0:
        return f"{m}m {sec:02d}s"
    elif m > 0:
        return f"{m}m"
    return f"{s}s"


def format_time_description(time_expr: str) -> str:
    try:
        delays, interval, terminal = parse_time_expr(time_expr)
    except Exception:
        return ""
    lines = []
    if delays:
        cumul = delays[0]
        lines.append(f"First alarm at {_fmt_seconds(cumul)}")
        if len(delays) > 1:
            remaining = delays[1:]
            if len(set(remaining)) == 1 and len(remaining) > 1:
                cumul += remaining[0] * len(remaining)
                lines.append(f"Then {len(remaining)} more every {_fmt_seconds(remaining[0])}")
            else:
                for i, d in enumerate(remaining):
                    cumul += d
                    lines.append(f"  #{i+2} at {_fmt_seconds(cumul)} (+{_fmt_seconds(d)})")
    if terminal:
        lines.append("Stops after last alarm")
    else:
        if delays:
            lines.append(f"Then repeats every {_fmt_seconds(interval)}")
        else:
            lines.append(f"Repeats every {_fmt_seconds(interval)}")
    return "\n".join(lines)


class DragBtn(QPushButton):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.dragging = False
        self.last = QPoint()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            self.last = e.globalPosition().toPoint()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.MouseButton.LeftButton:
            cur = e.globalPosition().toPoint()
            d = cur - self.last
            if d.manhattanLength() > 1:
                self.dragging = True
                win = self.window()
                target = clamp_window_pos(win.pos() + d, win.width(), win.height())
                win.move(target)
                self.last = cur
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton and self.dragging:
            self.setDown(False)
            e.accept()
            return
        super().mouseReleaseEvent(e)


class DragStrip(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dragging = False
        self.last = QPoint()
        self.setCursor(Qt.CursorShape.SizeAllCursor)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = True
            self.last = e.globalPosition().toPoint()
            e.accept()
            return
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self.dragging and (e.buttons() & Qt.MouseButton.LeftButton):
            cur = e.globalPosition().toPoint()
            d = cur - self.last
            if d.manhattanLength() > 0:
                win = self.window()
                win.move(clamp_window_pos(win.pos() + d, win.width(), win.height()))
                self.last = cur
            e.accept()
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            e.accept()
            return
        super().mouseReleaseEvent(e)


class DragHost(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dragging = False
        self.last = QPoint()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = True
            self.last = e.globalPosition().toPoint()
            e.accept()
            return
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self.dragging and (e.buttons() & Qt.MouseButton.LeftButton):
            cur = e.globalPosition().toPoint()
            d = cur - self.last
            if d.manhattanLength() > 0:
                win = self.window()
                win.move(clamp_window_pos(win.pos() + d, win.width(), win.height()))
                self.last = cur
            e.accept()
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            e.accept()
            return
        super().mouseReleaseEvent(e)


class DragTabBar(QTabBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dragging = False
        self.last = QPoint()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            self.last = e.globalPosition().toPoint()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.MouseButton.LeftButton:
            cur = e.globalPosition().toPoint()
            d = cur - self.last
            if d.manhattanLength() > 1:
                self.dragging = True
                win = self.window()
                win.move(clamp_window_pos(win.pos() + d, win.width(), win.height()))
                self.last = cur
                e.accept()
                return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton and self.dragging:
            e.accept()
            return
        super().mouseReleaseEvent(e)


class Overlay(QWidget):
    restore = pyqtSignal()
    toggled = pyqtSignal(str, bool)
    start_stop = pyqtSignal()
    build_changed = pyqtSignal(str)

    def __init__(self):
        super().__init__(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.btns: Dict[str, DragBtn] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(4)

        self.top = QHBoxLayout()
        self.top.setContentsMargins(0, 0, 0, 0)
        self.top.setSpacing(4)
        self.up = DragBtn("")
        self.up.clicked.connect(self.restore.emit)
        self.play = DragBtn("")
        self.play.clicked.connect(self.start_stop.emit)
        root.addLayout(self.top)

        self.build_combo = QComboBox()
        self.build_combo.currentTextChanged.connect(self._on_build_changed)
        self.build_combo.setVisible(False)
        root.addWidget(self.build_combo)

        self.host = QWidget()
        self.lay = QVBoxLayout(self.host)
        self.lay.setContentsMargins(0, 0, 0, 0)
        self.lay.setSpacing(4)
        root.addWidget(self.host)

        self.set_align("left")
        self.update_icons(False)

    def _on_build_changed(self, text):
        if text:
            self.build_changed.emit(text)

    def set_align(self, side: str):
        while self.top.count():
            it = self.top.takeAt(0)
            w = it.widget()
            if w:
                w.setParent(None)
        if side == "right":
            self.top.addWidget(self.play)
            self.top.addWidget(self.up)
        else:
            self.top.addWidget(self.up)
            self.top.addWidget(self.play)

    def update_icons(self, running: bool):
        self.up.setText("\uE70E")
        self.play.setText("\u25A0" if running else "\u25B6")

    def update_builds(self, build_names: List[str], active: str, show: bool):
        self.build_combo.blockSignals(True)
        self.build_combo.clear()
        for name in build_names:
            self.build_combo.addItem(name)
        idx = self.build_combo.findText(active)
        if idx >= 0:
            self.build_combo.setCurrentIndex(idx)
        self.build_combo.setVisible(show and len(build_names) > 1)
        self.build_combo.blockSignals(False)

    def apply_styles(self, w: int, h: int, fs: int, gap: int):
        self.top.setSpacing(4)
        self.lay.setSpacing(max(0, int(gap)))
        ac = self.property("accent_color") or "#00b7ff"
        fc = self.property("font_color") or "#ffffff"
        s = (
            f"QPushButton{{font-size:{max(11, fs-1)}px;font-weight:700;color:{fc};background:#2d313a;"
            f"border:1px solid {ac};border-radius:6px;padding:0px;text-align:center;}}"
            f"QPushButton:hover{{background:#3b404c;}}"
            f"QPushButton:checked{{background:{ac};border:2px solid {ac};}}"
        )
        play_s = (
            f"QPushButton{{font-size:{max(11, fs-1)}px;font-weight:700;color:{fc};background:#2d313a;"
            f"border:1px solid {ac};border-radius:6px;padding:0px;text-align:center;}}"
            f"QPushButton:hover{{background:#3b404c;}}"
        )
        self.up.setStyleSheet(s)
        self.play.setStyleSheet(play_s)
        gap = self.top.spacing()
        total_w = max(20, int(w))
        half_l = max(10, (total_w - gap) // 2)
        half_r = max(10, total_w - gap - half_l)
        self.up.setFixedSize(half_l, h)
        self.play.setFixedSize(half_r, h)
        icon_font = max(7, min(int(h * 0.62), int(min(half_l, half_r) * 0.72)))
        play_icon_font = max(9, int(icon_font * 1.35))
        self.up.setStyleSheet(s + f"QPushButton{{font-family:'Segoe MDL2 Assets';font-size:{icon_font}px;}}")
        self.play.setStyleSheet(play_s + f"QPushButton{{font-size:{play_icon_font}px;}}")
        combo_s = (
            f"QComboBox{{font-size:{max(11, fs-1)}px;font-weight:700;color:{fc};background:#2d313a;"
            f"border:1px solid {ac};border-radius:6px;padding:2px 6px;}}"
            f"QComboBox::drop-down{{border:none;}}"
            f"QComboBox::down-arrow{{image:none;}}"
            f"QComboBox QAbstractItemView{{background:#2d313a;color:{fc};border:1px solid {ac};selection-background-color:{ac};}}"
        )
        self.build_combo.setStyleSheet(combo_s)
        self.build_combo.setFixedWidth(total_w)
        self.build_combo.setFixedHeight(max(20, h - 4))
        for b in self.btns.values():
            b.setStyleSheet(s)
            b.setFixedSize(w, h)

    def rebuild(self, actions: List[dict], active: List[str], w: int, h: int, fs: int, gap: int):
        while self.lay.count():
            it = self.lay.takeAt(0)
            wd = it.widget()
            if wd:
                wd.deleteLater()
        self.btns = {}

        for a in actions:
            aid = a["id"]
            b = DragBtn(a["name"])
            b.setCheckable(True)
            b.setChecked(aid in active)
            b.setFixedSize(w, h)
            b.clicked.connect(lambda checked, x=aid: self.toggled.emit(x, checked))
            self.lay.addWidget(b, 0, Qt.AlignmentFlag.AlignLeft)
            self.btns[aid] = b
        self.lay.addStretch()
        self.apply_styles(w, h, fs, gap)
        self.adjustSize()

    def sync(self, active: List[str]):
        for aid, b in self.btns.items():
            b.blockSignals(True)
            b.setChecked(aid in active)
            b.blockSignals(False)


TOOLTIP_STYLE = """
    QToolTip {
        background-color: #23262b;
        color: #e0e0e0;
        border: 1px solid #3b404c;
        border-radius: 8px;
        padding: 8px 12px;
        font-size: 13px;
        font-family: 'Segoe UI', Arial, sans-serif;
    }
"""

SLIDER_STYLE = """
    QSlider {
        min-height: 32px;
    }
    QSlider::groove:horizontal {
        height: 10px;
        background: #3b404c;
        border-radius: 5px;
    }
    QSlider::handle:horizontal {
        width: 18px;
        height: 32px;
        background: #0c84da;
        border-radius: 5px;
        margin: -11px 0;
    }
    QSlider::sub-page:horizontal {
        background: #0c84da;
        border-radius: 5px;
        height: 10px;
    }
"""


class SettingsWin(QWidget):
    def __init__(self, m: "Main"):
        super().__init__(None, Qt.WindowType.Window)
        self.m = m
        self.setWindowTitle("Settings")
        self.setWindowIcon(m.icon)
        self.setMinimumSize(800, 680)
        self.resize(960, 760)
        self.upd = False
        self.row_ids: List[str] = []

        section_hdr = (
            "font-size:13px; font-weight:700; color:#8899aa;"
            "border-bottom:1px solid #3b404c; padding-bottom:3px; margin-top:2px;"
        )
        add_input_h = 34

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(6)

        # ════════════════════════════════════════
        #  GENERAL SETTINGS
        # ════════════════════════════════════════

        # ── Voice ──
        lbl_voice = QLabel("Voice"); lbl_voice.setStyleSheet(section_hdr)
        root.addWidget(lbl_voice)
        row_voice = QGridLayout()
        row_voice.setHorizontalSpacing(10)
        row_voice.setVerticalSpacing(6)
        self.vol = QSlider(Qt.Orientation.Horizontal); self.vol.setRange(0, 100); self.vol.valueChanged.connect(m.update_vol)
        self.vol.setToolTip("Volume of spoken alarm announcements (0–100%)")
        self.rate = QSlider(Qt.Orientation.Horizontal); self.rate.setRange(100, 500); self.rate.valueChanged.connect(m.update_rate)
        self.rate.setToolTip("Speech rate of the TTS voice\nHigher = faster speech")
        for s in [self.vol, self.rate]:
            s.setStyleSheet(SLIDER_STYLE)
            s.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.rand = QSpinBox(); self.rand.setRange(0, 500); self.rand.valueChanged.connect(m.update_rate_rand)
        self.rand.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.rand.setToolTip("Randomly vary the speech rate by ±N each alarm fire\nAdds variety to avoid a monotone voice (0 = no randomization)")
        self.voice = QComboBox(); [self.voice.addItem(v["name"], v["id"]) for v in m.voices]
        self.voice.currentIndexChanged.connect(lambda: m.update_voice(self.voice.currentData()))
        self.voice.setToolTip("Select the text-to-speech voice for alarm announcements")
        self.volv = QLabel(); self.ratev = QLabel()
        row_voice.addWidget(QLabel("Volume"), 0, 0); row_voice.addWidget(self.vol, 0, 1); row_voice.addWidget(self.volv, 0, 2)
        row_voice.addWidget(QLabel("Audio Pitch"), 0, 3); row_voice.addWidget(self.rate, 0, 4); row_voice.addWidget(self.ratev, 0, 5)
        row_voice.addWidget(QLabel("Pitch Randomize"), 0, 6); row_voice.addWidget(self.rand, 0, 7)
        row_voice.addWidget(QLabel("Voice"), 1, 0); row_voice.addWidget(self.voice, 1, 1, 1, 7)
        row_voice.setColumnStretch(1, 2); row_voice.setColumnStretch(4, 2)
        root.addLayout(row_voice)

        # ── Appearance ──
        lbl_appear = QLabel("Appearance"); lbl_appear.setStyleSheet(section_hdr)
        root.addWidget(lbl_appear)
        row_size = QGridLayout()
        row_size.setHorizontalSpacing(10)
        row_size.setVerticalSpacing(6)
        self.w = QSlider(Qt.Orientation.Horizontal); self.w.setRange(10, 420); self.w.valueChanged.connect(m.update_w)
        self.w.setToolTip("Width of alarm buttons in the overlay and main window")
        self.h = QSlider(Qt.Orientation.Horizontal); self.h.setRange(10, 90); self.h.valueChanged.connect(m.update_h)
        self.h.setToolTip("Height of alarm buttons in the overlay and main window")
        self.fs = QSlider(Qt.Orientation.Horizontal); self.fs.setRange(2, 50); self.fs.valueChanged.connect(m.update_fs)
        self.fs.setToolTip("Font size for button labels")
        self.gap = QSlider(Qt.Orientation.Horizontal); self.gap.setRange(0, 40); self.gap.valueChanged.connect(m.update_gap)
        self.gap.setToolTip("Vertical spacing in pixels between alarm buttons")
        for s in [self.w, self.h, self.fs, self.gap]:
            s.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            s.setMinimumWidth(80)
            s.setStyleSheet(SLIDER_STYLE)
        self.wv = QLabel(); self.hv = QLabel(); self.fsv = QLabel(); self.gapv = QLabel()
        row_size.addWidget(QLabel("Width"), 0, 0); row_size.addWidget(self.w, 0, 1); row_size.addWidget(self.wv, 0, 2)
        row_size.addWidget(QLabel("Height"), 0, 3); row_size.addWidget(self.h, 0, 4); row_size.addWidget(self.hv, 0, 5)
        row_size.addWidget(QLabel("Font"), 0, 6); row_size.addWidget(self.fs, 0, 7); row_size.addWidget(self.fsv, 0, 8)
        row_size.addWidget(QLabel("Gap"), 0, 9); row_size.addWidget(self.gap, 0, 10); row_size.addWidget(self.gapv, 0, 11)
        row_size.setColumnStretch(1, 2); row_size.setColumnStretch(4, 2); row_size.setColumnStretch(7, 2); row_size.setColumnStretch(10, 2)
        self.align = QComboBox(); self.align.addItem("Arrow Left", "left"); self.align.addItem("Arrow Right", "right")
        self.align.setToolTip("Side of the overlay where the close-overlay arrow button appears")
        self.align.currentIndexChanged.connect(lambda: m.update_align(self.align.currentData()))
        tabs_btn = QPushButton("Customize Tabs")
        tabs_btn.setToolTip("Add, remove, rename, or reorder the tabs in the main window")
        tabs_btn.clicked.connect(m.open_tabs_dialog)
        row_size.addWidget(QLabel("Arrow"), 1, 0); row_size.addWidget(self.align, 1, 1)
        row_size.addWidget(tabs_btn, 1, 3, 1, 3)
        root.addLayout(row_size)

        # ── Options ──
        lbl_opts = QLabel("Options"); lbl_opts.setStyleSheet(section_hdr)
        root.addWidget(lbl_opts)
        opts_grid = QGridLayout()
        opts_grid.setHorizontalSpacing(20)
        opts_grid.setVerticalSpacing(2)
        self.mintray = QCheckBox("Minimize to tray")
        self.mintray.setToolTip("When minimized, hide to the system tray\ninstead of appearing in the taskbar")
        self.mintray.stateChanged.connect(lambda: m.set_min_tray(self.mintray.isChecked()))
        lbl_mintray = QLabel("Hide to tray when minimized.")
        lbl_mintray.setStyleSheet("color:#8899aa; font-size:11px;")
        self.overlay_build = QCheckBox("Show Build in Overlay")
        self.overlay_build.setToolTip("Display a build selector dropdown in the compact overlay window\nOnly visible when multiple builds exist for the current race")
        self.overlay_build.stateChanged.connect(lambda: m.set_include_build_in_overlay(self.overlay_build.isChecked()))
        lbl_overlay_build = QLabel("Build dropdown in overlay.")
        lbl_overlay_build.setStyleSheet("color:#8899aa; font-size:11px;")
        self.overlay_lock = QCheckBox("Lock Overlay Buttons")
        self.overlay_lock.setToolTip("In overlay mode, disable all alarm toggle buttons\nOnly PLAY/STOP and Close Overlay remain clickable\nPrevents accidental alarm toggling while playing")
        self.overlay_lock.stateChanged.connect(lambda: m.set_overlay_buttons_locked(self.overlay_lock.isChecked()))
        lbl_overlay_lock = QLabel("Disable action buttons in overlay.")
        lbl_overlay_lock.setStyleSheet("color:#8899aa; font-size:11px;")
        self.initial_delay_spin = QSpinBox()
        self.initial_delay_spin.setRange(0, 300)
        self.initial_delay_spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.initial_delay_spin.setFixedWidth(52)
        self.initial_delay_spin.setToolTip("Seconds to count down before alarms start when PLAY is pressed.\nThe button shows the countdown. Click again to cancel.\nSet to 0 to start instantly.")
        self.initial_delay_spin.valueChanged.connect(lambda v: m.set_initial_delay(v))
        delay_ctrl = QWidget()
        delay_ctrl_l = QHBoxLayout(delay_ctrl)
        delay_ctrl_l.setContentsMargins(0, 0, 0, 0)
        delay_ctrl_l.setSpacing(4)
        lbl_delay_lbl = QLabel("Initial Delay (s):")
        lbl_delay_lbl.setToolTip("Seconds to count down before PLAY starts all alarms\nUseful to give yourself time to get in-game before alarms begin")
        delay_ctrl_l.addWidget(lbl_delay_lbl)
        delay_ctrl_l.addWidget(self.initial_delay_spin)
        delay_ctrl_l.addStretch()
        lbl_delay_desc = QLabel("Countdown before PLAY\nactivates alarms (0 = instant).")
        lbl_delay_desc.setStyleSheet("color:#8899aa; font-size:11px;")
        opts_grid.addWidget(self.mintray, 0, 0)
        opts_grid.addWidget(lbl_mintray, 1, 0)
        opts_grid.addWidget(self.overlay_build, 0, 1)
        opts_grid.addWidget(lbl_overlay_build, 1, 1)
        opts_grid.addWidget(self.overlay_lock, 0, 2)
        opts_grid.addWidget(lbl_overlay_lock, 1, 2)
        opts_grid.addWidget(delay_ctrl, 0, 3)
        opts_grid.addWidget(lbl_delay_desc, 1, 3)
        for c in range(4):
            opts_grid.setColumnStretch(c, 1)
        root.addLayout(opts_grid)

        # ════════════════════════════════════════
        #  RACE SETTINGS
        # ════════════════════════════════════════

        self.title = QLabel()
        self.title.setStyleSheet("font-size:15px; font-weight:700; padding-top:4px; padding-bottom:2px; border-bottom:2px solid #3b404c;")
        root.addWidget(self.title)

        # ── Colors (per race) ──
        row_colors = QHBoxLayout()
        row_colors.setSpacing(12)
        self.color_btn = QPushButton("Accent Color"); self.color_btn.clicked.connect(m.pick_accent_color)
        self.color_btn.setToolTip("Choose the accent/highlight color for buttons and borders\nEach race can have its own accent color")
        self.color_sw = QLabel(); self.color_sw.setFixedSize(18, 18)
        self.fcolor_btn = QPushButton("Font Color"); self.fcolor_btn.clicked.connect(m.pick_font_color)
        self.fcolor_btn.setToolTip("Choose the text color for all buttons\nEach race can have its own font color")
        self.fcolor_sw = QLabel(); self.fcolor_sw.setFixedSize(18, 18)
        row_colors.addWidget(self.color_btn); row_colors.addWidget(self.color_sw)
        row_colors.addWidget(self.fcolor_btn); row_colors.addWidget(self.fcolor_sw)
        row_colors.addStretch()
        root.addLayout(row_colors)

        # ── Builds ──
        lbl_builds = QLabel("Builds"); lbl_builds.setStyleSheet(section_hdr)
        root.addWidget(lbl_builds)
        row_builds = QHBoxLayout()
        row_builds.setSpacing(8)
        row_builds.addWidget(QLabel("Build:"))
        self.build_combo = QComboBox()
        self.build_combo.setMinimumWidth(140)
        self.build_combo.currentTextChanged.connect(self._on_build_changed)
        row_builds.addWidget(self.build_combo, 1)
        self.build_name_input = QLineEdit(); self.build_name_input.setPlaceholderText("New build name")
        self.build_name_input.setToolTip("Enter a name for a new build\nBuilds store separate alarm configurations\nfor the currently selected race")
        row_builds.addWidget(self.build_name_input, 1)
        add_build = QPushButton("Add Build"); add_build.clicked.connect(self._add_build)
        add_build.setToolTip("Create a new empty build\nwith the name entered to the left")
        clone_build = QPushButton("Clone Build"); clone_build.clicked.connect(self._clone_build)
        clone_build.setToolTip("Duplicate the current build\nwith a new name entered to the left\nCopies all alarms from the current build")
        del_build = QPushButton("Delete Build"); del_build.clicked.connect(self._delete_build)
        del_build.setToolTip("Delete the currently selected build\nCannot delete the last remaining build")
        row_builds.addWidget(add_build); row_builds.addWidget(clone_build); row_builds.addWidget(del_build)
        root.addLayout(row_builds)

        # ── Alarms ──
        lbl_alarms = QLabel("Alarms"); lbl_alarms.setStyleSheet(section_hdr)
        root.addWidget(lbl_alarms)
        self.tbl = QTableWidget(); self.tbl.setColumnCount(7); self.tbl.setHorizontalHeaderLabels(["Button Name", "Spoken Text", "Time (s)", "Play", "Up", "Down", "Remove"])
        hdr = self.tbl.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, hdr.ResizeMode.Stretch); hdr.setSectionResizeMode(1, hdr.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, hdr.ResizeMode.Stretch)
        hdr.setSectionResizeMode(3, hdr.ResizeMode.ResizeToContents); hdr.setSectionResizeMode(4, hdr.ResizeMode.ResizeToContents); hdr.setSectionResizeMode(5, hdr.ResizeMode.ResizeToContents); hdr.setSectionResizeMode(6, hdr.ResizeMode.ResizeToContents)
        hi0 = QTableWidgetItem("Button Name"); hi0.setToolTip("Display name shown on the button\nUse commas for sequential names that\ncycle with each alarm firing")
        hi1 = QTableWidgetItem("Spoken Text"); hi1.setToolTip("Text spoken aloud when the alarm fires\nUse commas for sequential texts that\ncycle with each alarm firing")
        hi2 = QTableWidgetItem("Time (s)"); hi2.setToolTip("Alarm timing in seconds\nExamples:\n  29 = Repeat every 29 seconds\n  140,29 = Wait 2m 20s, then every 29s\n  45,90,0 = At 45s, then 90s, then stop")
        self.tbl.setHorizontalHeaderItem(0, hi0); self.tbl.setHorizontalHeaderItem(1, hi1); self.tbl.setHorizontalHeaderItem(2, hi2)
        self.tbl.cellChanged.connect(self.cell_changed)
        root.addWidget(self.tbl, 1)

        ar = QHBoxLayout()
        ar.setSpacing(4)
        self.n = QLineEdit(); self.n.setPlaceholderText("New button name"); self.n.setMinimumHeight(add_input_h)
        self.n.setToolTip("Display name for the button\nUse commas to cycle names with each alarm\nExample: Pool, Tech")
        self.say = QLineEdit(); self.say.setPlaceholderText("Spoken text"); self.say.setMinimumHeight(add_input_h)
        self.say.setToolTip("Text spoken when alarm fires\nUse commas to cycle texts with each alarm\nExample: Poool, Tech Lab")
        self.sec = QLineEdit(); self.sec.setPlaceholderText("Time e.g. 120,29 or 120,29,0"); self.sec.setText("120"); self.sec.setMinimumHeight(add_input_h)
        self.sec.setToolTip("Alarm timing in seconds\nExamples:\n  29 = Repeat every 29 seconds\n  140,29 = Wait 2m 20s, then every 29s\n  45,90,90,0 = At 45s, 90s, 90s, then stop")
        add = DragBtn("Add"); add.clicked.connect(self.add_action); add.setMinimumHeight(add_input_h)
        add.setToolTip("Add a new alarm with the settings\nentered in the fields to the left")
        ar.addWidget(self.n, 3); ar.addWidget(self.say, 3); ar.addWidget(self.sec, 3); ar.addWidget(add)
        root.addLayout(ar)
        self.voice_refresh_timer = QTimer(self); self.voice_refresh_timer.timeout.connect(self.refresh_voice_list); self.voice_refresh_timer.start(4000)
        self.refresh()

    def _on_build_changed(self, text):
        if text and not self.upd:
            self.m.switch_build(text)
            self.refresh()

    def _add_build(self):
        name = self.build_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing Name", "Enter a build name.")
            return
        if not self.m.add_build(name):
            QMessageBox.warning(self, "Duplicate", f"Build '{name}' already exists for {self.m.active_race}.")
            return
        self.build_name_input.clear()
        self.refresh()

    def _clone_build(self):
        name = self.build_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Missing Name", "Enter a name for the cloned build.")
            return
        if not self.m.clone_build(name):
            QMessageBox.warning(self, "Duplicate", f"Build '{name}' already exists for {self.m.active_race}.")
            return
        self.build_name_input.clear()
        self.refresh()

    def _delete_build(self):
        if not self.m.delete_build():
            QMessageBox.warning(self, "Cannot Delete", "Cannot delete the last build.")
            return
        self.refresh()

    def refresh_voice_list(self):
        prev = self.voice.currentData() or self.m.voice_id
        self.m.refresh_voices()
        self.voice.blockSignals(True)
        self.voice.clear()
        for v in self.m.voices:
            self.voice.addItem(v["name"], v["id"])
        idx = self.voice.findData(prev)
        self.voice.setCurrentIndex(idx if idx >= 0 else (0 if self.voice.count() > 0 else -1))
        self.voice.blockSignals(False)

    def refresh(self):
        m = self.m
        self.refresh_voice_list()
        self.title.setText(f"{m.active_race} Settings")
        self.vol.setValue(int(m.volume * 100)); self.volv.setText(f"{int(m.volume*100)}%")
        self.w.setValue(m.bw); self.wv.setText(str(m.bw))
        self.h.setValue(m.bh); self.hv.setText(str(m.bh))
        self.fs.setValue(m.fs); self.fsv.setText(str(m.fs))
        self.rate.setValue(m.voice_rate); self.ratev.setText(str(m.voice_rate)); self.rand.setValue(m.voice_rate_rand)
        self.mintray.setChecked(m.min_to_tray)
        self.overlay_build.setChecked(m.include_build_in_overlay)
        self.overlay_lock.setChecked(m.overlay_buttons_locked)
        self.initial_delay_spin.setValue(m.initial_delay)
        self.gap.setValue(m.btn_gap); self.gapv.setText(str(m.btn_gap))
        self.color_sw.setStyleSheet(f"background:{m.accent_color}; border:1px solid #777; border-radius:3px;")
        self.fcolor_sw.setStyleSheet(f"background:{m.font_color}; border:1px solid #777; border-radius:3px;")
        idx = self.align.findData(m.arrow_align); self.align.setCurrentIndex(max(0, idx))
        if m.voice_id:
            idx = self.voice.findData(m.voice_id); self.voice.setCurrentIndex(idx if idx >= 0 else 0)

        self.upd = True
        self.build_combo.blockSignals(True)
        self.build_combo.clear()
        builds = m.builds.get(m.active_race, [])
        for b in builds:
            self.build_combo.addItem(b["name"])
        active_build = m.active_build.get(m.active_race, "Default")
        bidx = self.build_combo.findText(active_build)
        if bidx >= 0:
            self.build_combo.setCurrentIndex(bidx)
        self.build_combo.blockSignals(False)
        self.upd = False

        acts = m.actions[m.active_race]
        self.upd = True; self.tbl.setRowCount(len(acts)); self.row_ids = []
        for r, a in enumerate(acts):
            self.row_ids.append(a["id"])
            name_item = QTableWidgetItem(a.get("name_expr", a["name"]))
            name_parts = a.get("name_parts", [])
            if len(name_parts) > 1:
                name_item.setToolTip(f"Button label cycles through:\n" + "\n".join(f"  {i+1}. {p}" for i, p in enumerate(name_parts)))
            else:
                name_item.setToolTip(f"Button label: {a.get('name_expr', a['name'])}")
            self.tbl.setItem(r, 0, name_item)

            spoken_item = QTableWidgetItem(a.get("spoken_expr", a.get("spoken", a["name"])))
            spoken_parts = a.get("spoken_parts", [])
            if len(spoken_parts) > 1:
                spoken_item.setToolTip(f"Voice cycles through:\n" + "\n".join(f"  {i+1}. {p}" for i, p in enumerate(spoken_parts)))
            else:
                spoken_item.setToolTip(f"Voice says: {a.get('spoken_expr', a.get('spoken', a['name']))}")
            self.tbl.setItem(r, 1, spoken_item)

            delays = [int(x) for x in a.get("initial_delays", []) if int(x) >= 0]
            vals = [str(x) for x in delays] + [str(a["interval"])]
            if a.get("terminal_zero", False):
                vals.append("0")
            time_text = ",".join(vals) if vals else str(a["interval"])
            it = QTableWidgetItem(time_text)
            desc = format_time_description(a.get("time_expr", time_text))
            it.setToolTip(desc if desc else time_text)
            self.tbl.setItem(r, 2, it)
            pl = DragBtn("\u25B6"); pl.setFixedSize(28, 22); pl.setToolTip("Preview: speak this alarm's text now"); pl.clicked.connect(lambda checked=False, aid=a["id"]: self.play_spoken(aid))
            up = DragBtn("\u2191"); up.setFixedSize(28, 22); up.clicked.connect(lambda checked=False, aid=a["id"]: self.move_up(aid))
            dn = DragBtn("\u2193"); dn.setFixedSize(28, 22); dn.clicked.connect(lambda checked=False, aid=a["id"]: self.move_down(aid))
            rm = DragBtn("X"); rm.setFixedSize(28, 22); rm.clicked.connect(lambda checked=False, aid=a["id"]: self.rm(aid))
            self.tbl.setCellWidget(r, 3, pl)
            self.tbl.setCellWidget(r, 4, up)
            self.tbl.setCellWidget(r, 5, dn)
            self.tbl.setCellWidget(r, 6, rm)
        self.upd = False

    def cell_changed(self, r: int, c: int):
        if self.upd or r < 0 or r >= len(self.row_ids):
            return
        aid = self.row_ids[r]; a = self.m.find(aid)
        if not a:
            return
        if c == 0:
            nv = (self.tbl.item(r, 0).text() or "").strip()
            if not nv: self.upd = True; self.tbl.item(r, 0).setText(a.get("name_expr", a["name"])); self.upd = False; return
            a["name_expr"] = nv
            a["name_parts"] = parse_text_seq(nv)
            a["name"] = (a["name_parts"][0] if a["name_parts"] else nv)
        if c == 1:
            sv = (self.tbl.item(r, 1).text() or "").strip()
            if not sv: self.upd = True; self.tbl.item(r, 1).setText(a.get("spoken_expr", a.get("spoken", a["name"]))); self.upd = False; return
            a["spoken_expr"] = sv
            a["spoken_parts"] = parse_text_seq(sv)
            a["spoken"] = (a["spoken_parts"][0] if a["spoken_parts"] else sv)
        if c == 2:
            raw = (self.tbl.item(r, 2).text() or "").strip()
            try:
                dly, interval, terminal = parse_time_expr(raw)
            except Exception:
                vals = [str(x) for x in a.get("initial_delays", [])] + [str(a["interval"])]
                if a.get("terminal_zero", False):
                    vals.append("0")
                cur = ",".join(vals) if vals else str(a["interval"])
                self.upd = True; self.tbl.item(r, 2).setText(cur); self.upd = False; return
            a["interval"] = interval; a["initial_delays"] = dly; a["terminal_zero"] = terminal; a["time_expr"] = raw; self.m.restart_if_active(aid)
        self.m.rebuild_main(); self.m.rebuild_overlay(); self.m.save_timer.start(200)

    def add_action(self):
        n = self.n.text().strip()
        if not n: QMessageBox.warning(self, "Missing Name", "Enter an action name."); return
        spoken = (self.say.text() or n).strip()
        raw = (self.sec.text() or "").strip()
        self.m.add_action(n, spoken, raw); self.n.clear(); self.say.clear(); self.sec.setText("120"); self.refresh()

    def rm(self, aid: str):
        self.m.remove_action(aid); self.refresh()

    def play_spoken(self, aid: str):
        a = self.m.find(aid)
        if not a:
            return
        m = self.m
        delta = random.randint(-m.voice_rate_rand, m.voice_rate_rand) if m.voice_rate_rand > 0 else 0
        speak(a.get("spoken", a.get("name", "")), m.volume, max(100, min(1000, m.voice_rate + delta)), m.voice_id)

    def move_up(self, aid: str):
        moved = self.m.reorder_action(aid, -1)
        self.refresh()
        if moved:
            QTimer.singleShot(0, lambda: self._move_cursor_to_cell(aid, 4))

    def move_down(self, aid: str):
        moved = self.m.reorder_action(aid, 1)
        self.refresh()
        if moved:
            QTimer.singleShot(0, lambda: self._move_cursor_to_cell(aid, 5))

    def _move_cursor_to_cell(self, aid: str, col: int):
        idx = next((i for i, x in enumerate(self.row_ids) if x == aid), -1)
        if idx < 0:
            return
        widget = self.tbl.cellWidget(idx, col)
        if widget:
            pos = widget.mapToGlobal(QPoint(widget.width() // 2, widget.height() // 2))
            QCursor.setPos(pos)


class TabsDialog(QWidget):
    def __init__(self, m: "Main"):
        super().__init__(None, Qt.WindowType.Window)
        self.setWindowTitle("Customize Tabs")
        self.setWindowIcon(m.icon)
        self.setMinimumWidth(500)
        self.m = m
        self.races = list(m.races)
        self.names = list(m.tab_names)
        self.abbrevs = list(m.tab_abbrevs)
        self._building = False

        lay = QVBoxLayout(self)
        lay.setSpacing(8)

        info = QLabel("Customize the tabs shown in the main window.\nDisplay Name: shown on the tab.  Short Name: used when the window is narrow.")
        info.setStyleSheet("color:#8899aa; font-size:11px;")
        info.setWordWrap(True)
        lay.addWidget(info)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(5)
        self.tbl.setHorizontalHeaderLabels(["Display Name", "Short Name", "↑", "↓", "Remove"])
        hdr = self.tbl.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, hdr.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, hdr.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, hdr.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, hdr.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, hdr.ResizeMode.ResizeToContents)
        self.tbl.cellChanged.connect(self._cell_changed)
        lay.addWidget(self.tbl, 1)
        self._rebuild_table()

        add_row = QHBoxLayout()
        add_row.setSpacing(6)
        self.new_name = QLineEdit(); self.new_name.setPlaceholderText("New tab name")
        self.new_name.returnPressed.connect(self._add)
        self.new_abbrev = QLineEdit(); self.new_abbrev.setPlaceholderText("Short"); self.new_abbrev.setFixedWidth(60); self.new_abbrev.setMaxLength(4)
        self.new_abbrev.setToolTip("Short name shown when the window is too narrow for the full name (max 4 chars)")
        add_btn = QPushButton("Add Tab"); add_btn.clicked.connect(self._add)
        add_row.addWidget(QLabel("Name:")); add_row.addWidget(self.new_name, 2)
        add_row.addWidget(QLabel("Short:")); add_row.addWidget(self.new_abbrev); add_row.addWidget(add_btn)
        lay.addLayout(add_row)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("OK"); ok_btn.clicked.connect(self._ok)
        cancel_btn = QPushButton("Cancel"); cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(ok_btn); btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

    def _rebuild_table(self):
        self._building = True
        self.tbl.setRowCount(len(self.names))
        for r in range(len(self.names)):
            ni = QTableWidgetItem(self.names[r])
            ai = QTableWidgetItem(self.abbrevs[r])
            self.tbl.setItem(r, 0, ni)
            self.tbl.setItem(r, 1, ai)
            up_btn = QPushButton("↑"); up_btn.setFixedSize(28, 22)
            up_btn.clicked.connect(lambda checked=False, i=r: self._move(i, -1))
            dn_btn = QPushButton("↓"); dn_btn.setFixedSize(28, 22)
            dn_btn.clicked.connect(lambda checked=False, i=r: self._move(i, 1))
            rm_btn = QPushButton("✕"); rm_btn.setFixedSize(28, 22)
            rm_btn.setEnabled(len(self.names) > 1)
            rm_btn.setToolTip("Remove this tab (all its builds and alarms will be lost)")
            rm_btn.clicked.connect(lambda checked=False, i=r: self._remove(i))
            self.tbl.setCellWidget(r, 2, up_btn)
            self.tbl.setCellWidget(r, 3, dn_btn)
            self.tbl.setCellWidget(r, 4, rm_btn)
        self._building = False

    def _cell_changed(self, r: int, c: int):
        if self._building:
            return
        if c == 0 and r < len(self.names):
            val = (self.tbl.item(r, 0).text() or "").strip()
            if val:
                self.names[r] = val
        elif c == 1 and r < len(self.abbrevs):
            val = (self.tbl.item(r, 1).text() or "").strip()
            self.abbrevs[r] = val or self.names[r][:1]

    def _move(self, i: int, delta: int):
        j = i + delta
        if 0 <= j < len(self.races):
            for lst in (self.races, self.names, self.abbrevs):
                lst[i], lst[j] = lst[j], lst[i]
            self._rebuild_table()

    def _remove(self, i: int):
        if len(self.races) <= 1:
            return
        race = self.races[i]
        has_data = bool(self.m.builds.get(race, [{}])[0].get("actions"))
        if has_data:
            confirm = QMessageBox.question(self, "Remove Tab",
                f"Remove tab '{self.names[i]}'?\nAll its builds and alarms will be permanently deleted.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel)
            if confirm != QMessageBox.StandardButton.Yes:
                return
        self.races.pop(i); self.names.pop(i); self.abbrevs.pop(i)
        self._rebuild_table()

    def _add(self):
        name = self.new_name.text().strip()
        if not name:
            return
        if name in self.races:
            QMessageBox.warning(self, "Duplicate", f"A tab named '{name}' already exists.")
            return
        abbrev = self.new_abbrev.text().strip() or name[:1]
        self.races.append(name); self.names.append(name); self.abbrevs.append(abbrev)
        self.new_name.clear(); self.new_abbrev.clear()
        self._rebuild_table()

    def _ok(self):
        for r in range(len(self.names)):
            item0 = self.tbl.item(r, 0); item1 = self.tbl.item(r, 1)
            if item0:
                val = item0.text().strip()
                if val: self.names[r] = val
            if item1:
                val = item1.text().strip()
                self.abbrevs[r] = val or self.names[r][:1]
        if not all(n.strip() for n in self.names):
            QMessageBox.warning(self, "Invalid", "All tabs must have a display name.")
            return
        self.m.apply_tab_config(self.races, self.names, self.abbrevs)
        self.close()


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MacroHelper")
        self.setMinimumSize(120, 180)
        self.setWindowFlags(Qt.WindowType.Window)

        self.races = ["Zerg", "Protoss", "Terran"]; self.active_race = "Zerg"
        self.tab_names = list(self.races)
        self.tab_abbrevs = ["Z", "P", "T"]
        self.volume = 1.0; self.fs = 14; self.bw = 240; self.bh = 42; self.voice_rate = 200; self.voice_rate_rand = 0
        self.voices = voices(); self.voice_id = self.voices[0]["id"] if self.voices else None
        self.arrow_align = "left"; self.min_to_tray = True; self.btn_gap = 4; self.accent_color = "#00b7ff"; self.font_color = "#ffffff"; self.quitting = False; self.next_id = 1000
        self.include_build_in_overlay = False
        self.overlay_buttons_locked = False
        self.initial_delay = 0
        self._countdown_timer: Optional[QTimer] = None
        self._countdown_remaining = 0
        self.race_colors: Dict[str, dict] = {r: {"accent": "#00b7ff", "font": "#ffffff"} for r in self.races}
        self.resize(max(120, self.bw + 26), 420)

        default_actions = self.defaults()
        self.builds: Dict[str, List[dict]] = {}
        self.active_build: Dict[str, str] = {}
        for race in self.races:
            self.builds[race] = [{"name": "Default", "actions": default_actions[race]}]
            self.active_build[race] = "Default"
        self.actions: Dict[str, list] = {}
        self._sync_actions_from_builds()

        self.timers: Dict[str, QTimer] = {}; self.timer_meta: Dict[str, dict] = {}; self.main_btns: Dict[str, QPushButton] = {}
        self.overlay_hidden_ids: set[str] = set()
        self.ov: Optional[Overlay] = None; self.sw: Optional[SettingsWin] = None
        self.cfg_dir = app_dir(); self.cfg = os.path.join(self.cfg_dir, "config.json")
        self.save_timer = QTimer(self); self.save_timer.setSingleShot(True); self.save_timer.timeout.connect(self.save)
        self.progress_timer = QTimer(self); self.progress_timer.timeout.connect(self.refresh_progress_styles); self.progress_timer.start(120)
        self.ui(); self.tray(); self.apply_theme(); self.apply_font(); self.load()

    def _sync_actions_from_builds(self):
        for race in self.races:
            build_name = self.active_build.get(race, "Default")
            builds = self.builds.get(race, [])
            found = False
            for b in builds:
                if b["name"] == build_name:
                    self.actions[race] = b["actions"]
                    found = True
                    break
            if not found and builds:
                self.actions[race] = builds[0]["actions"]
                self.active_build[race] = builds[0]["name"]
            elif not builds:
                defs = self.defaults()
                self.builds[race] = [{"name": "Default", "actions": defs[race]}]
                self.active_build[race] = "Default"
                self.actions[race] = self.builds[race][0]["actions"]

    def refresh_voices(self):
        latest = voices()
        if not latest:
            return
        self.voices = latest
        if not any(v["id"] == self.voice_id for v in self.voices):
            self.voice_id = self.voices[0]["id"]

    def defaults(self):
        return {
            "Zerg": [{"id": "z_inject", "name": "Inject", "name_expr": "Inject", "name_parts": ["Inject"], "spoken": "Inject", "spoken_expr": "Inject", "spoken_parts": ["Inject"], "interval": 29, "initial_delays": [], "terminal_zero": False, "time_expr": "29"}, {"id": "z_upgrades", "name": "Upgrades", "name_expr": "Upgrades", "name_parts": ["Upgrades"], "spoken": "Upgrades", "spoken_expr": "Upgrades", "spoken_parts": ["Upgrades"], "interval": 120, "initial_delays": [], "terminal_zero": False, "time_expr": "120"}],
            "Protoss": [{"id": "p_probe", "name": "Probe", "name_expr": "Probe", "name_parts": ["Probe"], "spoken": "Probe", "spoken_expr": "Probe", "spoken_parts": ["Probe"], "interval": 12, "initial_delays": [], "terminal_zero": False, "time_expr": "12"}, {"id": "p_chrono", "name": "Chrono Boost", "name_expr": "Chrono Boost", "name_parts": ["Chrono Boost"], "spoken": "Chrono", "spoken_expr": "Chrono", "spoken_parts": ["Chrono"], "interval": 20, "initial_delays": [], "terminal_zero": False, "time_expr": "20"}, {"id": "p_upgrades", "name": "Upgrades", "name_expr": "Upgrades", "name_parts": ["Upgrades"], "spoken": "Upgrades", "spoken_expr": "Upgrades", "spoken_parts": ["Upgrades"], "interval": 120, "initial_delays": [], "terminal_zero": False, "time_expr": "120"}],
            "Terran": [{"id": "t_scv", "name": "SCV", "name_expr": "SCV", "name_parts": ["SCV"], "spoken": "SCV", "spoken_expr": "SCV", "spoken_parts": ["SCV"], "interval": 12, "initial_delays": [], "terminal_zero": False, "time_expr": "12"}, {"id": "t_mule", "name": "MULE / Scan", "name_expr": "MULE / Scan", "name_parts": ["MULE / Scan"], "spoken": "Mule", "spoken_expr": "Mule", "spoken_parts": ["Mule"], "interval": 64, "initial_delays": [], "terminal_zero": False, "time_expr": "64"}, {"id": "t_supply", "name": "Supply Depot", "name_expr": "Supply Depot", "name_parts": ["Supply Depot"], "spoken": "Supply", "spoken_expr": "Supply", "spoken_parts": ["Supply"], "interval": 21, "initial_delays": [], "terminal_zero": False, "time_expr": "21"}, {"id": "t_upgrades", "name": "Upgrades", "name_expr": "Upgrades", "name_parts": ["Upgrades"], "spoken": "Upgrades", "spoken_expr": "Upgrades", "spoken_parts": ["Upgrades"], "interval": 120, "initial_delays": [], "terminal_zero": False, "time_expr": "120"}],
        }

    def ui(self):
        root = QWidget(); self.setCentralWidget(root)
        lay = QVBoxLayout(root); lay.setContentsMargins(6, 6, 6, 6); lay.setSpacing(5)
        self.tabs = DragTabBar()
        [self.tabs.addTab(n) for n in self.tab_names]
        self.tabs.setExpanding(True)
        self.tabs.setUsesScrollButtons(False)
        self.tabs.setDrawBase(False)
        self.tabs.currentChanged.connect(self.race_changed)
        lay.addWidget(self.tabs)

        self.build_combo = QComboBox()
        self.build_combo.currentTextChanged.connect(self._on_build_combo_changed)
        lay.addWidget(self.build_combo)
        self._refresh_build_combo()

        top_wrap = DragHost()
        top_wrap_l = QHBoxLayout(top_wrap)
        top_wrap_l.setContentsMargins(0, 0, 0, 0)
        top_wrap_l.setSpacing(0)
        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(2)
        self.hide_btn = DragBtn(""); self.hide_btn.clicked.connect(self.enter_overlay)
        self.settings_btn = DragBtn(""); self.settings_btn.clicked.connect(self.open_settings)
        self.start_btn = DragBtn(""); self.start_btn.clicked.connect(self.toggle_start_stop)
        self.top_row = top
        top.addWidget(self.hide_btn); top.addWidget(self.start_btn); top.addWidget(self.settings_btn)
        top_wrap_l.addLayout(top)
        lay.addWidget(top_wrap)
        self.scr = QScrollArea(); self.scr.setWidgetResizable(True); self.host = QWidget(); self.bl = QVBoxLayout(self.host); self.bl.setContentsMargins(0, 0, 0, 0); self.bl.setSpacing(self.btn_gap)
        self.scr.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scr.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scr.setFrameShape(QFrame.Shape.NoFrame)
        self.scr.setWidget(self.host); lay.addWidget(self.scr); self.rebuild_main()

    def _on_build_combo_changed(self, text):
        if text and hasattr(self, 'save_timer'):
            self.switch_build(text)

    def _refresh_build_combo(self):
        self.build_combo.blockSignals(True)
        self.build_combo.clear()
        builds = self.builds.get(self.active_race, [])
        for b in builds:
            self.build_combo.addItem(b["name"])
        active_name = self.active_build.get(self.active_race, "Default")
        idx = self.build_combo.findText(active_name)
        if idx >= 0:
            self.build_combo.setCurrentIndex(idx)
        self.build_combo.setVisible(len(builds) > 1)
        self.build_combo.blockSignals(False)

    def switch_build(self, name: str):
        race = self.active_race
        builds = self.builds.get(race, [])
        if not any(b["name"] == name for b in builds):
            return
        if self.active_build.get(race) == name:
            return
        self._cancel_countdown()
        self.stop_all()
        self.active_build[race] = name
        self._sync_actions_from_builds()
        self._refresh_build_combo()
        self.rebuild_main(); self.rebuild_overlay()
        if self.sw and self.sw.isVisible():
            self.sw.refresh()
        self.save_timer.start(200)

    def add_build(self, name: str) -> bool:
        race = self.active_race
        builds = self.builds.get(race, [])
        if any(b["name"] == name for b in builds):
            return False
        builds.append({"name": name, "actions": []})
        self.stop_all()
        self.active_build[race] = name
        self._sync_actions_from_builds()
        self._refresh_build_combo()
        self.rebuild_main(); self.rebuild_overlay()
        self.save_timer.start(200)
        return True

    def clone_build(self, name: str) -> bool:
        race = self.active_race
        builds = self.builds.get(race, [])
        if any(b["name"] == name for b in builds):
            return False
        current_actions = self.actions.get(race, [])
        cloned = copy.deepcopy(current_actions)
        builds.append({"name": name, "actions": cloned})
        self.stop_all()
        self.active_build[race] = name
        self._sync_actions_from_builds()
        self._refresh_build_combo()
        self.rebuild_main(); self.rebuild_overlay()
        self.save_timer.start(200)
        return True

    def delete_build(self) -> bool:
        race = self.active_race
        builds = self.builds.get(race, [])
        if len(builds) <= 1:
            return False
        current_name = self.active_build.get(race, "Default")
        builds[:] = [b for b in builds if b["name"] != current_name]
        self.stop_all()
        self.active_build[race] = builds[0]["name"]
        self._sync_actions_from_builds()
        self._refresh_build_combo()
        self.rebuild_main(); self.rebuild_overlay()
        self.save_timer.start(200)
        return True

    def set_include_build_in_overlay(self, on: bool):
        self.include_build_in_overlay = on
        self.rebuild_overlay()
        self.save_timer.start(200)

    def set_overlay_buttons_locked(self, locked: bool):
        self.overlay_buttons_locked = locked
        if self.ov:
            for b in self.ov.btns.values():
                b.setEnabled(not locked)
        self.save_timer.start(200)

    def set_initial_delay(self, v: int):
        self.initial_delay = max(0, int(v))
        self.save_timer.start(200)

    def open_tabs_dialog(self):
        self._tabs_dlg = TabsDialog(self)
        self._tabs_dlg.show(); self._tabs_dlg.raise_(); self._tabs_dlg.activateWindow()

    def apply_tab_config(self, new_races: List[str], new_names: List[str], new_abbrevs: List[str]):
        if not new_races:
            return
        self._cancel_countdown()
        for aid in list(self.timers.keys()):
            self.stop(aid)
        for r in [r for r in self.races if r not in new_races]:
            self.builds.pop(r, None); self.actions.pop(r, None)
            self.race_colors.pop(r, None); self.active_build.pop(r, None)
        for r in new_races:
            if r not in self.builds:
                self.builds[r] = [{"name": "Default", "actions": []}]
            if r not in self.active_build:
                self.active_build[r] = "Default"
            if r not in self.race_colors:
                self.race_colors[r] = {"accent": "#00b7ff", "font": "#ffffff"}
        self.races = new_races
        self.tab_names = new_names
        self.tab_abbrevs = new_abbrevs
        if self.active_race not in self.races:
            self.active_race = self.races[0]
        self._sync_actions_from_builds()
        self.tabs.blockSignals(True)
        while self.tabs.count():
            self.tabs.removeTab(0)
        for name in self.tab_names:
            self.tabs.addTab(name)
        self.tabs.setCurrentIndex(self.races.index(self.active_race))
        self.tabs.blockSignals(False)
        rc = self.race_colors.get(self.active_race, {})
        self.accent_color = rc.get("accent", self.accent_color)
        self.font_color = rc.get("font", self.font_color)
        self._refresh_build_combo()
        self.apply_theme(); self.apply_btn_style()
        self.rebuild_main(); self.rebuild_overlay()
        if self.sw and self.sw.isVisible():
            self.sw.refresh()
        self.update_start_stop_ui()
        self.update_compact_controls()
        self.save_timer.start(200)

    def tray(self):
        icon = QIcon(rp("icon.ico")) if os.path.exists(rp("icon.ico")) else QApplication.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.icon = icon
        self.setWindowIcon(icon); self.tr = QSystemTrayIcon(icon, self); self.tr.activated.connect(self.tr_activated)
        m = QMenu(); a1 = QAction("Open", self); a1.triggered.connect(self.restore); a2 = QAction("Hide GUI", self); a2.triggered.connect(self.enter_overlay); a3 = QAction("Quit", self); a3.triggered.connect(self.quit)
        m.addAction(a1); m.addAction(a2); m.addSeparator(); m.addAction(a3); self.tr.setContextMenu(m)
        if self.min_to_tray:
            self.tr.show()

    def find(self, aid: str):
        for r in self.races:
            for a in self.actions[r]:
                if a["id"] == aid: return a
        return None

    def race_for_action(self, aid: str) -> Optional[str]:
        for r in self.races:
            for a in self.actions[r]:
                if a["id"] == aid:
                    return r
        return None

    def _seq_value(self, parts: List[str], idx: int, fallback: str) -> str:
        if not parts:
            return fallback
        i = max(0, min(idx, len(parts) - 1))
        return parts[i]

    def _action_display_name(self, a: dict, aid: str) -> str:
        m = self.timer_meta.get(aid)
        idx = int(m.get("fires", 0)) if m else 0
        parts = a.get("name_parts") or parse_text_seq(a.get("name_expr", a.get("name", "")))
        return self._seq_value(parts, idx, a.get("name", ""))

    def refresh_action_button_texts(self):
        for a in self.actions[self.active_race]:
            aid = a["id"]
            txt = self._action_display_name(a, aid)
            if aid in self.main_btns:
                self.main_btns[aid].setText(txt)
            if self.ov and aid in self.ov.btns:
                self.ov.btns[aid].setText(txt)

    def reorder_action(self, aid: str, delta: int) -> bool:
        arr = self.actions[self.active_race]
        idx = next((i for i, x in enumerate(arr) if x["id"] == aid), -1)
        if idx < 0:
            return False
        ni = max(0, min(len(arr) - 1, idx + delta))
        if ni == idx:
            return False
        item = arr.pop(idx)
        arr.insert(ni, item)
        self.rebuild_main(); self.rebuild_overlay(); self.save_timer.start(200)
        return True

    def rebuild_main(self):
        while self.bl.count():
            i = self.bl.takeAt(0); w = i.widget()
            if w: w.deleteLater()
        self.main_btns = {}
        avail_w = max(self.bw, self.centralWidget().width() - 12) if self.centralWidget() else self.bw
        for a in self.actions[self.active_race]:
            aid = a["id"]; b = DragBtn(self._action_display_name(a, aid)); b.setCheckable(True); b.setChecked(aid in self.timers); b.setFixedSize(avail_w, self.bh)
            b.clicked.connect(lambda checked, x=aid: self.toggle(x, checked)); self.bl.addWidget(b); self.main_btns[aid] = b
        self.apply_btn_style(); self.auto_fit_window()

    def auto_fit_window(self):
        count = max(1, len(self.actions[self.active_race]))
        list_h = (count * self.bh) + (max(0, count - 1) * self.btn_gap)
        self.scr.setFixedHeight(max(20, list_h))
        target_w = max(120, self.bw + 12)
        tab_h = self.tabs.sizeHint().height()
        build_h = (self.build_combo.sizeHint().height() + 5) if self.build_combo.isVisible() else 0
        top_h = self.bh
        target_h = max(140, 12 + tab_h + build_h + 5 + top_h + 5 + list_h + 12)
        self.setMinimumWidth(target_w)
        self.setMinimumHeight(target_h)
        if not self.isMaximized():
            self.resize(target_w, target_h)

    def rebuild_overlay(self):
        if not self.ov: return
        self.ov.setProperty("accent_color", self.accent_color)
        self.ov.setProperty("font_color", self.font_color)
        build_names = [b["name"] for b in self.builds.get(self.active_race, [])]
        active_build = self.active_build.get(self.active_race, "Default")
        self.ov.update_builds(build_names, active_build, self.include_build_in_overlay)
        vis = []
        for a in self.actions[self.active_race]:
            if a["id"] in self.overlay_hidden_ids:
                continue
            x = dict(a)
            x["name"] = self._action_display_name(a, a["id"])
            vis.append(x)
        self.ov.set_align(self.arrow_align)
        self.ov.rebuild(vis, list(self.timers.keys()), self.bw, self.bh, self.fs, self.btn_gap)
        for b in self.ov.btns.values():
            b.setEnabled(not self.overlay_buttons_locked)
        self.update_start_stop_ui()

    def apply_btn_style(self):
        c = self.accent_color
        fc = self.font_color
        s = f"QPushButton{{font-size:{max(11,self.fs-1)}px;font-weight:700;color:{fc};background:#2d313a;border:1px solid {c};border-radius:6px;padding:0px;text-align:center;}}QPushButton:hover{{background:#3b404c;}}QPushButton:checked{{background:{c};border:2px solid {c};}}"
        avail_w = max(self.bw, self.centralWidget().width() - 12) if self.centralWidget() else self.bw
        for b in self.main_btns.values():
            b.setStyleSheet(s); b.setFixedSize(avail_w, self.bh)
        self.hide_btn.setStyleSheet(s); self.hide_btn.setFixedHeight(self.bh)
        self.settings_btn.setStyleSheet(s); self.settings_btn.setFixedHeight(self.bh)
        self.start_btn.setStyleSheet(s); self.start_btn.setFixedHeight(self.bh)
        icon_style = s + "QPushButton{font-family:'Segoe MDL2 Assets';}"
        self.hide_btn.setStyleSheet(icon_style)
        self.settings_btn.setStyleSheet(icon_style)
        self.start_btn.setStyleSheet(s)
        self.update_control_icon_font()
        self.update_control_icons()
        combo_s = (
            f"QComboBox{{font-size:{max(11, self.fs-1)}px;font-weight:700;color:{fc};background:#2d313a;"
            f"border:1px solid {c};border-radius:6px;padding:2px 6px;}}"
            f"QComboBox::drop-down{{border:none;width:20px;}}"
            f"QComboBox::down-arrow{{image:none;}}"
            f"QComboBox QAbstractItemView{{background:#2d313a;color:{fc};border:1px solid {c};selection-background-color:{c};}}"
        )
        self.build_combo.setStyleSheet(combo_s)
        if self.ov: self.ov.apply_styles(self.bw, self.bh, self.fs, self.btn_gap)
        self.refresh_progress_styles(); self.update_start_stop_ui(); self.update_compact_controls()

    def _inactive_style(self) -> str:
        c = self.accent_color
        fc = self.font_color
        return f"QPushButton{{font-size:{max(11,self.fs-1)}px;font-weight:700;color:{fc};background:#2d313a;border:1px solid {c};border-radius:6px;padding:0px;text-align:center;}}QPushButton:hover{{background:#3b404c;}}"

    def _progress_style(self, p: float) -> str:
        p = max(0.0, min(1.0, p)); p2 = min(1.0, p + 0.001)
        c = self.accent_color
        fc = self.font_color
        return f"QPushButton{{font-size:{max(11,self.fs-1)}px;font-weight:700;color:{fc};background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {c}, stop:{p:.3f} {c}, stop:{p2:.3f} #2d313a, stop:1 #2d313a);border:1px solid {c};border-radius:6px;padding:0px;text-align:center;}}"

    def refresh_progress_styles(self):
        now = time.monotonic(); inactive = self._inactive_style()
        for aid, b in self.main_btns.items():
            m = self.timer_meta.get(aid)
            if not m: b.setStyleSheet(inactive); continue
            cycle = max(1.0, float(m.get("cycle", m["interval"])))
            elapsed = min(cycle, max(0.0, now - float(m["start"])))
            b.setStyleSheet(self._progress_style(elapsed / cycle))
        if self.ov:
            for aid, b in self.ov.btns.items():
                m = self.timer_meta.get(aid)
                if not m: b.setStyleSheet(inactive); continue
                cycle = max(1.0, float(m.get("cycle", m["interval"])))
                elapsed = min(cycle, max(0.0, now - float(m["start"])))
                b.setStyleSheet(self._progress_style(elapsed / cycle))

    def sync(self):
        for aid, b in self.main_btns.items(): b.blockSignals(True); b.setChecked(aid in self.timers); b.blockSignals(False)
        if self.ov: self.ov.sync(list(self.timers.keys()))

    def toggle(self, aid: str, on: bool):
        a = self.find(aid)
        if not a: return
        if on: self.start(a)
        else: self.stop(aid)

    def start(self, a: dict):
        aid = a["id"]
        self.overlay_hidden_ids.discard(aid)
        race = self.race_for_action(aid)
        if race:
            for running in list(self.timers.keys()):
                rr = self.race_for_action(running)
                if rr and rr != race:
                    self.stop(running)
        self.stop(aid)
        interval = max(1, int(a["interval"]))
        t = QTimer(self)
        t.timeout.connect(lambda x=aid: self._on_timer(x))
        self.timers[aid] = t
        self.timer_meta[aid] = {"start": time.monotonic(), "interval": interval, "cycle": interval, "pre_delays": [], "pre_idx": 0, "terminal_zero": False, "fires": 0}
        t.start(interval * 1000)
        self.sync(); self.refresh_progress_styles(); self.update_start_stop_ui(); self.refresh_action_button_texts()

    def start_with_initial_delay(self, a: dict):
        aid = a["id"]
        self.overlay_hidden_ids.discard(aid)
        race = self.race_for_action(aid)
        if race:
            for running in list(self.timers.keys()):
                rr = self.race_for_action(running)
                if rr and rr != race:
                    self.stop(running)
        self.stop(aid)
        interval = max(1, int(a["interval"]))
        delays = [int(x) for x in a.get("initial_delays", []) if int(x) >= 0]
        terminal = bool(a.get("terminal_zero", False))
        if (not delays) and a.get("time_expr"):
            try:
                dly, interval2, term2 = parse_time_expr(a.get("time_expr", ""))
                delays = dly
                interval = max(1, int(interval2))
                terminal = term2
            except Exception:
                pass
        first_delay = delays[0] if delays else 0
        t = QTimer(self)
        t.timeout.connect(lambda x=aid: self._on_timer(x))
        self.timers[aid] = t
        self.timer_meta[aid] = {"start": time.monotonic(), "interval": interval, "cycle": (first_delay if first_delay > 0 else interval), "pre_delays": delays, "pre_idx": 0, "terminal_zero": terminal, "fires": 0}
        t.start((first_delay if first_delay > 0 else interval) * 1000)
        self.sync(); self.refresh_progress_styles(); self.update_start_stop_ui(); self.refresh_action_button_texts()

    def _on_timer(self, aid: str):
        m = self.timer_meta.get(aid)
        t = self.timers.get(aid)
        if not m or not t:
            return
        pre = m.get("pre_delays", [])
        idx = int(m.get("pre_idx", 0))
        terminal = bool(m.get("terminal_zero", False))
        self.tick(aid)
        if isinstance(pre, list) and idx < len(pre):
            idx += 1
            m["pre_idx"] = idx
            if idx < len(pre):
                nxt = int(pre[idx])
            else:
                if terminal:
                    self.stop(aid)
                    self.overlay_hidden_ids.add(aid)
                    self.rebuild_overlay()
                    return
                nxt = int(m["interval"])
            m["start"] = time.monotonic()
            m["cycle"] = max(1, nxt)
            t.start(max(1, nxt) * 1000)
            return

    def stop(self, aid: str):
        t = self.timers.pop(aid, None)
        if t: t.stop(); t.deleteLater()
        self.timer_meta.pop(aid, None); self.sync(); self.refresh_progress_styles(); self.update_start_stop_ui(); self.refresh_action_button_texts()

    def tick(self, aid: str):
        a = self.find(aid)
        if not a: self.stop(aid); return
        m = self.timer_meta.get(aid)
        idx = int(m.get("fires", 0)) if m else 0
        if m:
            m["start"] = time.monotonic()
            m["fires"] = idx + 1
        sparts = a.get("spoken_parts") or parse_text_seq(a.get("spoken_expr", a.get("spoken", a.get("name", ""))))
        spoken = self._seq_value(sparts, idx, a.get("spoken", a["name"]))
        delta = random.randint(-self.voice_rate_rand, self.voice_rate_rand) if self.voice_rate_rand > 0 else 0
        speak(spoken, self.volume, max(100, min(1000, self.voice_rate + delta)), self.voice_id)
        self.refresh_action_button_texts()

    def start_all(self):
        if self.sw and self.sw.isVisible():
            self.sync_settings_time_values()
        if self.initial_delay > 0:
            self._begin_countdown(self.initial_delay)
        else:
            self._start_all_immediate()

    def _start_all_immediate(self):
        for a in self.actions[self.active_race]:
            self.start_with_initial_delay(a)

    def _begin_countdown(self, seconds: int):
        self._cancel_countdown()
        self._countdown_remaining = seconds
        self._countdown_timer = QTimer(self)
        self._countdown_timer.timeout.connect(self._countdown_tick)
        self._countdown_timer.start(1000)
        self._update_countdown_ui()

    def _countdown_tick(self):
        self._countdown_remaining -= 1
        if self._countdown_remaining <= 0:
            self._countdown_remaining = 0
            if self._countdown_timer:
                self._countdown_timer.stop()
                self._countdown_timer.deleteLater()
                self._countdown_timer = None
            self._start_all_immediate()
        else:
            self._update_countdown_ui()

    def _update_countdown_ui(self):
        txt = str(self._countdown_remaining)
        self.start_btn.setText(txt)
        if self.ov:
            self.ov.play.setText(txt)

    def _cancel_countdown(self):
        if self._countdown_timer:
            self._countdown_timer.stop()
            self._countdown_timer.deleteLater()
            self._countdown_timer = None
        if self._countdown_remaining > 0:
            self._countdown_remaining = 0
            self.update_start_stop_ui()

    def sync_settings_time_values(self):
        if not self.sw:
            return
        for r, aid in enumerate(self.sw.row_ids):
            a = self.find(aid)
            if not a:
                continue
            item = self.sw.tbl.item(r, 2)
            if not item:
                continue
            raw = (item.text() or "").strip()
            try:
                dly, interval, terminal = parse_time_expr(raw)
            except Exception:
                continue
            a["interval"] = interval
            a["initial_delays"] = dly
            a["terminal_zero"] = terminal
            a["time_expr"] = raw

    def stop_all(self):
        self._cancel_countdown()
        self.overlay_hidden_ids.clear()
        for a in list(self.actions[self.active_race]): self.stop(a["id"])
        self.rebuild_overlay()

    def all_active(self) -> bool:
        acts = self.actions[self.active_race]; return len(acts) > 0 and all(a["id"] in self.timers for a in acts)

    def any_active(self) -> bool:
        return any(a["id"] in self.timers for a in self.actions[self.active_race])

    def toggle_start_stop(self):
        if self._countdown_remaining > 0:
            self._cancel_countdown()
        elif self.any_active():
            self.stop_all()
        else:
            self.start_all()
        self.update_start_stop_ui()

    def update_start_stop_ui(self):
        if self._countdown_remaining > 0:
            return
        running = self.any_active()
        self.start_btn.setText("\u25A0" if running else "\u25B6")
        if self.ov:
            self.ov.update_icons(running)

    def update_control_icons(self):
        self.hide_btn.setText("\uE70D")
        self.settings_btn.setText("\uE713")

    def update_compact_controls(self):
        if not self.centralWidget():
            return
        avail = max(30, self.centralWidget().width() - 12)
        n = max(1, self.tabs.count())
        tab_w = max(12, int(avail / n))
        short = tab_w < 72
        labels = self.tab_abbrevs if short else self.tab_names
        for i in range(min(n, len(labels))):
            if self.tabs.tabText(i) != labels[i]:
                self.tabs.setTabText(i, labels[i])
        c = self.accent_color; fc = self.font_color
        self.tabs.setStyleSheet(
            "QTabBar{border:none;}"
            f"QTabBar::tab{{background:#2d313a;color:{fc};border:1px solid {c};padding:2px 0px;margin-right:0px;margin-bottom:0px;border-radius:6px;font-size:18px;font-weight:700;}}"
            f"QTabBar::tab:selected{{background:{c};color:{fc};border:1px solid {c};}}"
        )
        sp = self.top_row.spacing()
        w1 = max(10, int((avail - 2 * sp) / 3))
        w2 = w1
        w3 = max(10, int(avail - w1 - w2 - 2 * sp))
        self.hide_btn.setFixedWidth(w1)
        self.start_btn.setFixedWidth(w2)
        self.settings_btn.setFixedWidth(w3)
        for b in self.main_btns.values():
            b.setFixedSize(avail, self.bh)
        self.build_combo.setFixedWidth(avail)
        self.update_control_icon_font()

    def update_control_icon_font(self):
        icon_px = max(7, min(int(self.bh * 0.62), int(min(self.hide_btn.width(), self.start_btn.width(), self.settings_btn.width()) * 0.72)))
        play_icon_px = max(9, int(icon_px * 1.35))
        c = self.accent_color
        fc = self.font_color
        s = f"QPushButton{{font-size:{max(11,self.fs-1)}px;font-weight:700;color:{fc};background:#2d313a;border:1px solid {c};border-radius:6px;padding:0px;text-align:center;}}QPushButton:hover{{background:#3b404c;}}QPushButton:checked{{background:{c};border:2px solid {c};}}"
        icon_style = s + f"QPushButton{{font-family:'Segoe MDL2 Assets';font-size:{icon_px}px;}}"
        start_icon_style = s + f"QPushButton{{font-size:{play_icon_px}px;}}"
        self.hide_btn.setStyleSheet(icon_style)
        self.settings_btn.setStyleSheet(icon_style)
        self.start_btn.setStyleSheet(start_icon_style)

    def pick_accent_color(self):
        col = QColorDialog.getColor(QColor(self.accent_color), self, "Accent Color")
        if not col.isValid():
            return
        self.accent_color = col.name()
        self.race_colors[self.active_race]["accent"] = self.accent_color
        self.apply_theme(); self.apply_btn_style(); self.rebuild_overlay()
        if self.sw and self.sw.isVisible():
            self.sw.refresh()
        self.save_timer.start(200)

    def pick_font_color(self):
        col = QColorDialog.getColor(QColor(self.font_color), self, "Font Color")
        if not col.isValid():
            return
        self.font_color = col.name()
        self.race_colors[self.active_race]["font"] = self.font_color
        self.apply_theme(); self.apply_btn_style(); self.rebuild_overlay()
        if self.sw and self.sw.isVisible():
            self.sw.refresh()
        self.save_timer.start(200)

    def race_changed(self, i: int):
        if 0 <= i < len(self.races):
            self._cancel_countdown()
            old_race = self.active_race
            self.race_colors[old_race] = {"accent": self.accent_color, "font": self.font_color}
            self.active_race = self.races[i]
            rc = self.race_colors.get(self.active_race, {})
            self.accent_color = rc.get("accent", self.accent_color)
            self.font_color = rc.get("font", self.font_color)
            for aid in list(self.timers.keys()):
                rr = self.race_for_action(aid)
                if rr and rr != self.active_race:
                    self.stop(aid)
            self._refresh_build_combo()
            self.apply_theme(); self.apply_btn_style()
            self.rebuild_main(); self.rebuild_overlay()
            if self.sw and self.sw.isVisible(): self.sw.refresh()
            self.update_start_stop_ui(); self.save_timer.start(200)

    def add_action(self, name: str, spoken: str, time_expr: str):
        if any(a["name"].lower() == name.lower() for a in self.actions[self.active_race]): QMessageBox.warning(self, "Duplicate", f"'{name}' already exists in {self.active_race}."); return
        try:
            dly, sec, terminal = parse_time_expr(time_expr)
        except Exception:
            QMessageBox.warning(self, "Invalid Time", "Use: delay1,delay2,...,interval or ...,0 for terminal.")
            return
        aid = f"user_{self.next_id}"; self.next_id += 1
        norm = ",".join([str(x) for x in dly] + [("0" if terminal else str(sec))]) if dly else ("0" if terminal else str(sec))
        if terminal and (not dly):
            norm = f"{sec},0"
        nparts = parse_text_seq(name); sparts = parse_text_seq(spoken)
        was_all_active = self.all_active()
        self.actions[self.active_race].append({"id": aid, "name": (nparts[0] if nparts else name), "name_expr": name, "name_parts": nparts, "spoken": (sparts[0] if sparts else spoken), "spoken_expr": spoken, "spoken_parts": sparts, "interval": sec, "initial_delays": dly, "terminal_zero": terminal, "time_expr": norm})
        self.rebuild_main(); self.rebuild_overlay()
        if was_all_active:
            self.start_with_initial_delay(self.find(aid))
        self.save_timer.start(200)

    def remove_action(self, aid: str):
        acts = self.actions[self.active_race]
        acts[:] = [a for a in acts if a["id"] != aid]
        self.stop(aid); self.rebuild_main(); self.rebuild_overlay(); self.save_timer.start(200)

    def restart_if_active(self, aid: str):
        if aid in self.timers:
            a = self.find(aid)
            if a: self.start(a)

    def open_settings(self):
        self.refresh_voices()
        if self.sw is None: self.sw = SettingsWin(self)
        self.sw.refresh(); self.sw.show(); self.sw.raise_(); self.sw.activateWindow()

    def enter_overlay(self):
        if self.ov is None:
            self.ov = Overlay()
            self.ov.restore.connect(self.exit_overlay)
            self.ov.toggled.connect(self.toggle)
            self.ov.start_stop.connect(self.toggle_start_stop)
            self.ov.build_changed.connect(self.switch_build)
        self.rebuild_overlay()
        self.ov.move(clamp_window_pos(self.pos(), self.ov.width(), self.ov.height()))
        self.hide(); self.ov.show(); self.ov.raise_(); self.ov.activateWindow()
        fit_window_to_screen(self.ov)

    def exit_overlay(self):
        if self.ov:
            self.move(self.ov.pos())
            self.ov.hide()
        self.showNormal(); self.show(); self.raise_(); self.activateWindow()
        fit_window_to_screen(self)

    def update_vol(self, v: int): self.volume = v / 100.0; self.sw and self.sw.volv.setText(f"{v}%"); self.save_timer.start(200)
    def update_w(self, v: int): self.bw = v; self.sw and self.sw.wv.setText(str(v)); self.rebuild_main(); self.rebuild_overlay(); self.save_timer.start(200)
    def update_h(self, v: int): self.bh = v; self.sw and self.sw.hv.setText(str(v)); self.rebuild_main(); self.rebuild_overlay(); self.save_timer.start(200)
    def update_fs(self, v: int): self.fs = v; self.sw and self.sw.fsv.setText(str(v)); self.apply_font(); self.save_timer.start(200)
    def update_rate(self, v: int): self.voice_rate = max(100, min(1000, v)); self.sw and self.sw.ratev.setText(str(self.voice_rate)); self.save_timer.start(200)
    def update_rate_rand(self, v: int): self.voice_rate_rand = v; self.save_timer.start(200)
    def update_voice(self, vid: str): self.voice_id = vid; self.save_timer.start(200)
    def update_align(self, a: str): self.arrow_align = "right" if a == "right" else "left"; self.rebuild_overlay(); self.save_timer.start(200)
    def update_gap(self, v: int):
        self.btn_gap = max(0, int(v))
        if self.sw and hasattr(self.sw, "gapv"): self.sw.gapv.setText(str(self.btn_gap))
        if hasattr(self, "top_row"): self.top_row.setSpacing(4)
        if hasattr(self, "bl"): self.bl.setSpacing(self.btn_gap)
        self.apply_btn_style()
        self.auto_fit_window()
        self.rebuild_overlay(); self.save_timer.start(200)
    def set_min_tray(self, on: bool):
        self.min_to_tray = on
        if on:
            self.tr.show()
        else:
            self.tr.hide()
        self.save_timer.start(200)

    def apply_font(self):
        f = self.font(); f.setPointSize(self.fs); self.setFont(f); self.apply_btn_style()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.update_compact_controls()
        if hasattr(self, "save_timer"):
            self.save_timer.start(200)

    def tr_activated(self, r):
        if r in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick): self.restore()

    def restore(self):
        if self.ov and self.ov.isVisible(): self.ov.show(); self.ov.raise_(); self.ov.activateWindow()
        else: self.showNormal(); self.show(); self.raise_(); self.activateWindow()

    def quit(self):
        self.quitting = True; self.tr.hide(); self.ov and self.ov.close(); self.sw and self.sw.close(); self.close()

    def changeEvent(self, e):
        if e.type() == e.Type.WindowStateChange and self.isMinimized() and self.min_to_tray:
            QTimer.singleShot(0, self.hide); self.tr.isVisible() and self.tr.showMessage("MacroHelper", "Running in system tray.", QSystemTrayIcon.MessageIcon.Information, 900)
        super().changeEvent(e)

    def closeEvent(self, e):
        self.quitting = True; self.save(); self.tr.hide(); self.ov and self.ov.close(); self.sw and self.sw.close(); e.accept()
        QApplication.instance().quit()

    def _parse_actions_list(self, raw_list: list) -> list:
        ok = []
        for a in raw_list:
            aid = str(a.get("id", "")).strip(); n = str(a.get("name", "")).strip(); sec = int(a.get("interval", 0)); sp = str(a.get("spoken", n))
            ne = str(a.get("name_expr", n)).strip() or n
            se = str(a.get("spoken_expr", sp)).strip() or sp
            dly = a.get("initial_delays", [])
            if not isinstance(dly, list): dly = []
            dly = [int(x) for x in dly if int(x) >= 0]
            tex = str(a.get("time_expr", "")).strip()
            term = bool(a.get("terminal_zero", False))
            if tex:
                try:
                    d2, i2, t2 = parse_time_expr(tex)
                    dly, sec, term = d2, i2, t2
                except Exception:
                    pass
            if aid and n and sec > 0:
                texpr = tex if tex else (",".join([str(x) for x in dly] + ([("0")] if term else [str(sec)])) if dly else (f"{sec},0" if term else str(sec)))
                nparts = parse_text_seq(ne); sparts = parse_text_seq(se)
                ok.append({"id": aid, "name": (nparts[0] if nparts else n), "name_expr": ne, "name_parts": nparts, "spoken": (sparts[0] if sparts else sp), "spoken_expr": se, "spoken_parts": sparts, "interval": sec, "initial_delays": dly, "terminal_zero": term, "time_expr": texpr})
        return ok

    def load(self):
        if not os.path.exists(self.cfg): return
        try:
            with open(self.cfg, "r", encoding="utf-8") as f: c = json.load(f)
            p = c.get("pos", [120, 120]); self.move(p[0], p[1]); self.volume = float(c.get("volume", 1.0)); self.fs = int(c.get("font_size", self.fs))
            self.bw = int(c.get("button_width", self.bw)); self.bh = int(c.get("button_height", self.bh)); self.voice_rate = max(100, min(1000, int(c.get("voice_rate", self.voice_rate)))); self.voice_rate_rand = int(c.get("voice_rate_rand", self.voice_rate_rand))
            self.voice_id = c.get("voice_id", self.voice_id); self.arrow_align = c.get("arrow_align", self.arrow_align); self.min_to_tray = bool(c.get("minimize_to_tray", True)); self.btn_gap = int(c.get("button_gap", self.btn_gap)); self.accent_color = str(c.get("accent_color", self.accent_color)); self.font_color = str(c.get("font_color", self.font_color))
            rc = c.get("race_colors")
            if isinstance(rc, dict):
                for r in self.races:
                    if r in rc and isinstance(rc[r], dict):
                        self.race_colors[r] = {"accent": str(rc[r].get("accent", "#00b7ff")), "font": str(rc[r].get("font", "#ffffff"))}
            else:
                for r in self.races:
                    self.race_colors[r] = {"accent": self.accent_color, "font": self.font_color}
            self.include_build_in_overlay = bool(c.get("include_build_in_overlay", False))
            self.overlay_buttons_locked = bool(c.get("overlay_buttons_locked", False))
            self.initial_delay = int(c.get("initial_delay", 0))
            saved_races = c.get("races")
            if isinstance(saved_races, list) and saved_races:
                self.races = [str(r).strip() for r in saved_races if str(r).strip()]
                raw_names = c.get("tab_names", [])
                self.tab_names = [str(n) for n in raw_names] if isinstance(raw_names, list) and len(raw_names) == len(self.races) else list(self.races)
                raw_abbrevs = c.get("tab_abbrevs", [])
                self.tab_abbrevs = [str(a) for a in raw_abbrevs] if isinstance(raw_abbrevs, list) and len(raw_abbrevs) == len(self.races) else [n[:1] for n in self.tab_names]
                self.tabs.blockSignals(True)
                while self.tabs.count():
                    self.tabs.removeTab(0)
                for name in self.tab_names:
                    self.tabs.addTab(name)
                self.tabs.blockSignals(False)
                rc2 = c.get("race_colors", {})
                for r in self.races:
                    if r not in self.race_colors:
                        entry = rc2.get(r, {}) if isinstance(rc2, dict) else {}
                        self.race_colors[r] = {"accent": str(entry.get("accent", "#00b7ff")), "font": str(entry.get("font", "#ffffff"))}
            ww = int(c.get("window_width", self.width()))
            wh = int(c.get("window_height", self.height()))
            self.resize(max(120, ww), max(180, wh))

            builds_data = c.get("builds_by_race")
            if isinstance(builds_data, dict):
                for r in self.races:
                    raw_builds = builds_data.get(r, [])
                    parsed_builds = []
                    for rb in raw_builds:
                        bname = str(rb.get("name", "")).strip()
                        bactions = self._parse_actions_list(rb.get("actions", []))
                        if bname:
                            parsed_builds.append({"name": bname, "actions": bactions})
                    if parsed_builds:
                        self.builds[r] = parsed_builds
                active_builds = c.get("active_build_by_race", {})
                for r in self.races:
                    ab = str(active_builds.get(r, "Default")).strip()
                    if any(b["name"] == ab for b in self.builds.get(r, [])):
                        self.active_build[r] = ab
            else:
                acts = c.get("actions_by_race")
                if isinstance(acts, dict):
                    for r in self.races:
                        raw = acts.get(r, [])
                        ok = self._parse_actions_list(raw)
                        if ok:
                            self.builds[r] = [{"name": "Default", "actions": ok}]
                            self.active_build[r] = "Default"

            self._sync_actions_from_builds()
            self.next_id = int(max(self.next_id, c.get("next_action_id", self.next_id))); self.active_race = c.get("active_race", self.active_race)
            ar_colors = self.race_colors.get(self.active_race, {})
            self.accent_color = ar_colors.get("accent", self.accent_color)
            self.font_color = ar_colors.get("font", self.font_color)
            if self.active_race in self.races: self.tabs.setCurrentIndex(self.races.index(self.active_race))
            if hasattr(self, "top_row"): self.top_row.setSpacing(4)
            if hasattr(self, "bl"): self.bl.setSpacing(self.btn_gap)
            self._refresh_build_combo()
            self.apply_theme(); self.apply_font(); self.rebuild_main()
            self.update_start_stop_ui()
            if self.min_to_tray:
                self.tr.show()
            else:
                self.tr.hide()
        except Exception as ex:
            print(f"Error loading config: {ex}")

    def save(self):
        builds_data = {}
        for r in self.races:
            builds_data[r] = []
            for b in self.builds.get(r, []):
                builds_data[r].append({"name": b["name"], "actions": b["actions"]})

        c = {
            "pos": [self.pos().x(), self.pos().y()], "volume": self.volume, "font_size": self.fs, "button_width": self.bw, "button_height": self.bh,
            "window_width": self.width(), "window_height": self.height(),
            "voice_rate": self.voice_rate, "voice_rate_rand": self.voice_rate_rand, "voice_id": self.voice_id, "arrow_align": self.arrow_align, "minimize_to_tray": self.min_to_tray, "button_gap": self.btn_gap, "accent_color": self.accent_color, "font_color": self.font_color,
            "race_colors": self.race_colors,
            "include_build_in_overlay": self.include_build_in_overlay,
            "overlay_buttons_locked": self.overlay_buttons_locked,
            "initial_delay": self.initial_delay,
            "races": self.races,
            "tab_names": self.tab_names,
            "tab_abbrevs": self.tab_abbrevs,
            "active_race": self.active_race,
            "builds_by_race": builds_data,
            "active_build_by_race": dict(self.active_build),
            "next_action_id": self.next_id,
        }
        try:
            if not os.path.exists(self.cfg_dir): os.makedirs(self.cfg_dir, exist_ok=True)
            target = self.cfg
            with open(target, "w", encoding="utf-8") as f: json.dump(c, f, indent=2)
            return
        except Exception:
            pass
        target = os.path.join(os.getcwd(), "config.json")
        try:
            with open(target, "w", encoding="utf-8") as f: json.dump(c, f, indent=2)
        except Exception as ex:
            print(f"Error saving config: {ex}")

    def apply_theme(self):
        c = self.accent_color
        fc = self.font_color
        self.setStyleSheet(TOOLTIP_STYLE + """
            QMainWindow,QWidget{background:#1d1f23;color:%s;font-family:'Segoe UI',Arial,sans-serif;}
            QLineEdit,QSpinBox,QTableWidget,QComboBox{background:#2d313a;border:1px solid #3b404c;border-radius:6px;color:%s;}
            QPushButton{background:#0c84da;border:none;border-radius:6px;color:%s;font-weight:700;padding:4px 8px;}
            QPushButton:hover{background:#19a0ff;} QPushButton:pressed{background:#0b6eb7;} QPushButton:checked{background:#00a6d8;}
            QScrollArea{border:none;background:transparent;}
            QTabBar{border:none;}
            QTabBar::tab{background:#2d313a;color:%s;border:1px solid %s;padding:2px 4px;margin-right:2px;margin-bottom:0px;border-radius:6px;font-size:18px;font-weight:700;}
            QTabBar::tab:selected{background:%s;color:%s;border:1px solid %s;}
        """ % (fc, fc, fc, fc, c, c, fc, c))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setStyle("Fusion")
    w = Main()
    w.show()
    sys.exit(app.exec())
