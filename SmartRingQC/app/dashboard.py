import os
import sys
import cv2
import uuid
import subprocess
import numpy as np
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QMessageBox,
    QSpinBox, QScrollArea, QFrame, QSizePolicy
)
from PyQt6.QtGui import QPixmap, QImage, QFont, QShortcut, QKeySequence
from PyQt6.QtCore import Qt, QTimer, QThread, QObject, pyqtSignal

try:
    from app.classifier import Classifier
    from app import db_handler, excel_writer
except ImportError:
    from classifier import Classifier
    import db_handler
    import excel_writer

# ================================================================
# CAMERA SLOTS CONFIGURATION
# ================================================================
CAMERA_SLOTS = [
    {
        "label":       "Webcam",
        "index":       0,
        "color":       "#546E7A",
        "hover":       "#78909C",
        "description": "HP Wide Vision HD Camera (built-in)",
    },
    {
        "label":       "Microscope USB",
        "index":       1,
        "color":       "#1565C0",
        "hover":       "#1976D2",
        "description": "USB Video capture card connected to microscope",
    },
    {
        "label":       "Microscope HDMI",
        "index":       2,
        "color":       "#6A1B9A",
        "hover":       "#7B1FA2",
        "description": "Microscope via HDMI → USB capture card → PC USB port",
    },
]

# ================================================================
# USB Camera — Open by DirectShow Device Name (most reliable)
# ================================================================
def open_camera_by_name(device_name, width=1280, height=720):
    """Opens a camera using its Windows DirectShow friendly name — same as VLC."""
    cap = cv2.VideoCapture(f"video={device_name}", cv2.CAP_DSHOW)
    if cap.isOpened():
        cap.set(cv2.CAP_PROP_FRAME_WIDTH,  width)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
        ret, _ = cap.read()
        if ret:
            return cap
        cap.release()
    return None


# ================================================================
# Background Camera Scanner — never blocks the main thread
# ================================================================
class CameraScanner(QObject):
    found = pyqtSignal(list)

    def run(self):
        results = []
        for i in range(10):
            try:
                cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
                if cap.isOpened():
                    ret, _ = cap.read()
                    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                    res = f"{w}×{h}" if (ret and w > 0) else "opened — no frame"
                    results.append((i, res))
                cap.release()
            except Exception:
                pass
        self.found.emit(results)


# ================================================================
# THEME
# ================================================================
BG_DEEP   = "#0B0D13"
BG_PANEL  = "#13151D"
BG_CARD   = "#1A1D26"
BG_INPUT  = "#1E2130"
BORDER    = "#252B3B"
ACCENT    = "#6366F1"
ACCENT2   = "#4F46E5"
GREEN     = "#22C55E"
RED       = "#EF4444"
ORANGE    = "#F97316"
BLUE      = "#3B82F6"
TEXT_PRI  = "#F1F5F9"
TEXT_SEC  = "#94A3B8"
TEXT_MUT  = "#475569"

# ================================================================
# Maximized Feed Window
# ================================================================
class MaximizedFeedWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Live Camera Feed - Maximized")
        self.resize(1024, 768)
        self.setStyleSheet(f"background-color: {BG_DEEP};")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet(f"background-color: #000000;")
        self.layout.addWidget(self.image_label)
        
        self.esc_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.esc_shortcut.activated.connect(self.close)
        
    def update_frame(self, pixmap):
        if not pixmap.isNull():
            self.image_label.setPixmap(pixmap.scaled(
                self.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.FastTransformation))


# ================================================================
# Main Dashboard
# ================================================================
class Dashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.classifier   = Classifier()
        self.waiting_images       = []
        self.current_image_index  = 0
        self.session_id   = f"#QC-{uuid.uuid4().hex[:4].upper()}"
        self._rec_blink   = True

        # Camera
        self.cap            = None
        self.camera_id      = CAMERA_SLOTS[0]["index"]
        self.active_slot    = 0
        self.wait_dir       = self.classifier.waiting_dir
        self._scanner_thread  = None
        self._scanner_worker  = None
        self.current_frame    = None
        self.max_feed_win     = None

        self.init_ui()
        self.init_camera()

        self.state_timer = QTimer(); self.state_timer.timeout.connect(self.refresh_state); self.state_timer.start(2000)
        self.camera_timer = QTimer(); self.camera_timer.timeout.connect(self.update_frame); self.camera_timer.start(30)
        self.clock_timer  = QTimer(); self.clock_timer.timeout.connect(self._update_clock); self.clock_timer.start(1000)
        self.rec_timer    = QTimer(); self.rec_timer.timeout.connect(self._blink_rec);      self.rec_timer.start(800)

        self.refresh_state()
        self._update_clock()
        self.setup_global_shortcuts()

    # ----------------------------------------------------------
    # Shortcuts
    # ----------------------------------------------------------
    def setup_global_shortcuts(self):
        def bind(k, m):
            s = QShortcut(QKeySequence(k), self)
            s.setContext(Qt.ShortcutContext.ApplicationShortcut)
            s.activated.connect(m)

        bind("A",                    self._safe_do_accept)
        bind("a",                    self._safe_do_accept)
        bind("R",                    self._safe_do_reject)
        bind("r",                    self._safe_do_reject)
        bind("W",                    self._safe_do_rework)
        bind("w",                    self._safe_do_rework)
        bind("S",                    self._safe_do_skip)
        bind("s",                    self._safe_do_skip)
        bind(Qt.Key.Key_Right,       self._safe_do_skip)
        bind(Qt.Key.Key_F5,          self._safe_do_refresh)
        bind(Qt.Key.Key_Delete,      self._safe_do_delete)
        bind(Qt.Key.Key_Space,       self._safe_do_capture)

    def _safe_do_accept(self):
        if not self.is_input_focused(): self.do_accept()
    def _safe_do_reject(self):
        if not self.is_input_focused(): self.do_reject()
    def _safe_do_rework(self):
        if not self.is_input_focused(): self.do_rework()
    def _safe_do_skip(self):
        if not self.is_input_focused(): self.do_skip()
    def _safe_do_refresh(self):
        if not self.is_input_focused(): self.refresh_state()
    def _safe_do_delete(self):
        if not self.is_input_focused(): self.do_delete()
    def _safe_do_capture(self):
        if not self.is_input_focused(): self.capture_image()

    # ----------------------------------------------------------
    # UI Construction
    # ----------------------------------------------------------
    def init_ui(self):
        self.setWindowTitle("SmartRingQC — Professional Quality Control")
        self.resize(1300, 800)
        self.setMinimumSize(1100, 700)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self.setStyleSheet(f"""
            QMainWindow, QWidget     {{ background-color: {BG_DEEP}; color: {TEXT_PRI}; font-family: 'Segoe UI', Arial, sans-serif; }}
            QLabel                   {{ background: transparent; }}
            QLineEdit                {{ background-color: {BG_INPUT}; color: {TEXT_PRI}; border: 1px solid {BORDER}; border-radius: 6px; padding: 8px 12px; font-size: 13px; }}
            QLineEdit:focus          {{ border-color: {ACCENT}; }}
            QComboBox                {{ background-color: {BG_INPUT}; color: {TEXT_PRI}; border: 1px solid {BORDER}; border-radius: 6px; padding: 8px 12px; font-size: 13px; }}
            QComboBox::drop-down     {{ border: none; width: 28px; }}
            QComboBox QAbstractItemView {{ background-color: {BG_CARD}; color: {TEXT_PRI}; border: 1px solid {BORDER}; selection-background-color: {ACCENT}; }}
            QScrollArea              {{ border: none; background: transparent; }}
            QScrollBar:horizontal    {{ height: 5px; background: {BG_CARD}; }}
            QScrollBar::handle:horizontal {{ background: {BORDER}; border-radius: 2px; }}
            QToolTip                 {{ background-color: {BG_CARD}; color: {TEXT_PRI}; border: 1px solid {BORDER}; padding: 4px 8px; }}
        """)

        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_navbar())
        root_layout.addWidget(self._hsep())

        content = QWidget()
        cl = QHBoxLayout(content)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(12)
        cl.addWidget(self._build_camera_panel(),   stretch=22)
        cl.addWidget(self._build_review_panel(),   stretch=28)
        cl.addWidget(self._build_controls_panel(), stretch=17)
        root_layout.addWidget(content, stretch=1)

        root_layout.addWidget(self._hsep())
        root_layout.addWidget(self._build_statusbar())

    def _hsep(self):
        f = QFrame(); f.setFixedHeight(1); f.setStyleSheet(f"background:{BORDER};"); return f

    def _panel(self):
        w = QWidget()
        w.setStyleSheet(f"QWidget {{ background-color:{BG_PANEL}; border-radius:10px; border:1px solid {BORDER}; }}")
        return w

    # ------------------------------------------------------------------
    # TOP NAVBAR
    # ------------------------------------------------------------------
    def _build_navbar(self):
        bar = QWidget(); bar.setFixedHeight(54)
        bar.setStyleSheet(f"background-color:{BG_PANEL}; border-bottom:1px solid {BORDER};")
        ly = QHBoxLayout(bar); ly.setContentsMargins(18, 0, 18, 0)

        # Logo
        ico = QLabel("⬡"); ico.setStyleSheet(f"color:{ACCENT}; font-size:22px; font-weight:bold;")
        ttl = QLabel("SmartRingQC"); ttl.setStyleSheet(f"color:{TEXT_PRI}; font-size:16px; font-weight:bold; letter-spacing:1px;")
        ly.addWidget(ico); ly.addWidget(ttl); ly.addStretch()

        # Center status badge
        self.nav_status_lbl = QLabel("● WAITING FOR REVIEW")
        self.nav_status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.nav_status_lbl.setStyleSheet(f"""
            color:{TEXT_PRI}; background-color:{BG_CARD}; border:1px solid {BORDER};
            border-radius:14px; padding:5px 20px; font-size:12px; font-weight:bold; letter-spacing:0.5px;
        """)
        ly.addWidget(self.nav_status_lbl); ly.addStretch()

        # Clock
        clk_lbl = QLabel("SYSTEM CLOCK")
        clk_lbl.setStyleSheet(f"color:{TEXT_MUT}; font-size:9px; letter-spacing:1px;")
        self.clock_display = QLabel("00:00:00")
        self.clock_display.setStyleSheet(f"color:{TEXT_PRI}; font-size:19px; font-weight:bold; letter-spacing:2px;")
        clk_col = QVBoxLayout(); clk_col.setSpacing(0)
        clk_col.addWidget(clk_lbl, alignment=Qt.AlignmentFlag.AlignCenter)
        clk_col.addWidget(self.clock_display, alignment=Qt.AlignmentFlag.AlignCenter)
        ly.addLayout(clk_col); ly.addSpacing(14)

        for icon in ["⚙", "⬡"]:
            b = QPushButton(icon); b.setFixedSize(32, 32); b.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            b.setStyleSheet(f"""
                QPushButton {{ background:{BG_CARD}; color:{TEXT_SEC}; border:1px solid {BORDER}; border-radius:16px; font-size:15px; }}
                QPushButton:hover {{ background:{BORDER}; color:{TEXT_PRI}; }}
            """)
            ly.addWidget(b); ly.addSpacing(4)
        return bar

    # ------------------------------------------------------------------
    # LEFT: CAMERA PANEL
    # ------------------------------------------------------------------
    def _build_camera_panel(self):
        panel = self._panel()
        ly = QVBoxLayout(panel); ly.setContentsMargins(10, 10, 10, 10); ly.setSpacing(7)

        # Camera header
        hdr = QHBoxLayout()
        self.rec_badge = QLabel("● LIVE CAMERA FEED")
        self.rec_badge.setStyleSheet(f"background:{RED}; color:white; font-size:10px; font-weight:bold; padding:3px 10px; border-radius:4px; letter-spacing:0.5px;")
        hdr.addWidget(self.rec_badge); hdr.addStretch()
        self.btn_maximize = QPushButton("⛶"); self.btn_maximize.setFixedSize(28, 28); self.btn_maximize.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_maximize.setStyleSheet(f"QPushButton{{background:{BG_CARD};color:{TEXT_SEC};border:1px solid {BORDER};border-radius:6px;font-size:13px;}} QPushButton:hover{{color:{TEXT_PRI};}}")
        self.btn_maximize.clicked.connect(self.toggle_maximized_feed)
        hdr.addWidget(self.btn_maximize); ly.addLayout(hdr)

        # Source tag
        self.cam_source_lbl = QLabel(f"SOURCE:  {CAMERA_SLOTS[0]['label']}  (index {CAMERA_SLOTS[0]['index']})")
        self.cam_source_lbl.setStyleSheet(f"color:{ACCENT}; background:{BG_CARD}; border:1px solid {BORDER}; border-radius:4px; padding:3px 10px; font-size:10px; font-weight:bold;")
        ly.addWidget(self.cam_source_lbl)

        # REC indicator
        self.rec_indicator = QLabel("⬤  REC  1080P • 60 FPS")
        self.rec_indicator.setStyleSheet(f"color:{RED}; font-size:10px; font-weight:bold; padding:1px 4px;")
        ly.addWidget(self.rec_indicator)

        # Live feed
        self.cam_lbl = QLabel("Initializing Camera…")
        self.cam_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cam_lbl.setStyleSheet(f"background:#090B11; color:{TEXT_MUT}; border:1px solid {BORDER}; border-radius:8px; font-size:13px;")
        self.cam_lbl.setMinimumSize(280, 280)
        self.cam_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        ly.addWidget(self.cam_lbl, stretch=1)

        # Camera slot buttons
        slot_row = QHBoxLayout(); slot_row.setSpacing(6)
        self.cam_slot_buttons = []
        for i, slot in enumerate(CAMERA_SLOTS):
            btn = QPushButton(slot["label"])
            btn.setCheckable(True); btn.setChecked(i == 0)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            btn.setToolTip(slot.get("description", slot["label"]))
            btn.setStyleSheet(self._cam_btn_style(slot, active=(i == 0)))
            btn.clicked.connect(lambda checked, idx=i: self.switch_camera_slot(idx))
            slot_row.addWidget(btn); self.cam_slot_buttons.append(btn)
        ly.addLayout(slot_row)

        # Capture row
        cap_row = QHBoxLayout(); cap_row.setSpacing(8)

        self.btn_scan = QPushButton("⊕")
        self.btn_scan.setFixedSize(40, 46); self.btn_scan.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_scan.setToolTip("Scan / detect all cameras")
        self.btn_scan.setStyleSheet(f"""
            QPushButton{{background:{BG_CARD};color:{TEXT_SEC};border:1px solid {BORDER};border-radius:8px;font-size:20px;font-weight:bold;}}
            QPushButton:hover{{background:{BORDER};color:{TEXT_PRI};}}
            QPushButton:disabled{{color:{TEXT_MUT};}}
        """)
        self.btn_scan.clicked.connect(self.scan_cameras)
        cap_row.addWidget(self.btn_scan)

        self.btn_capture = QPushButton("  📷  CAPTURE\n  [SPACE]")
        self.btn_capture.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_capture.setStyleSheet(f"""
            QPushButton{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {ACCENT2},stop:1 {ACCENT});color:white;font-size:13px;font-weight:bold;border-radius:8px;border:none;padding:4px;}}
            QPushButton:hover{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #5B5CF6,stop:1 #7071F7);}}
            QPushButton:pressed{{background:{ACCENT2};}}
        """)
        self.btn_capture.clicked.connect(self.capture_image)
        cap_row.addWidget(self.btn_capture, stretch=1)

        gear = QPushButton("⚙"); gear.setFixedSize(40, 46); gear.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        gear.setStyleSheet(f"QPushButton{{background:{BG_CARD};color:{TEXT_SEC};border:1px solid {BORDER};border-radius:8px;font-size:16px;}} QPushButton:hover{{background:{BORDER};color:{TEXT_PRI};}}")
        cap_row.addWidget(gear)
        ly.addLayout(cap_row)
        return panel

    # ------------------------------------------------------------------
    # MIDDLE: REVIEW PANEL
    # ------------------------------------------------------------------
    def _build_review_panel(self):
        panel = self._panel()
        ly = QVBoxLayout(panel); ly.setContentsMargins(10, 10, 10, 10); ly.setSpacing(7)

        # Header
        hdr = QHBoxLayout()
        self.badge_lbl = QLabel("Waiting For Review")
        self.badge_lbl.setStyleSheet(f"background:{BG_CARD}; color:{TEXT_PRI}; border:1px solid {BORDER}; border-radius:12px; padding:4px 14px; font-size:12px; font-weight:bold;")
        self.count_badge = QLabel("0"); self.count_badge.setFixedSize(26, 26)
        self.count_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.count_badge.setStyleSheet(f"background:{ACCENT}; color:white; border-radius:13px; font-size:11px; font-weight:bold;")
        hdr.addWidget(self.badge_lbl); hdr.addWidget(self.count_badge); hdr.addStretch()
        sess = QLabel(f"SESSION ID: {self.session_id}")
        sess.setStyleSheet(f"color:{TEXT_MUT}; font-size:10px; letter-spacing:0.5px;")
        hdr.addWidget(sess); ly.addLayout(hdr)

        # Main image
        self.image_lbl = QLabel("No Image Waiting")
        self.image_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_lbl.setStyleSheet(f"background:#090B11; color:{TEXT_MUT}; border:1px solid {BORDER}; border-radius:8px; font-size:14px;")
        self.image_lbl.setMinimumSize(300, 270)
        self.image_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        ly.addWidget(self.image_lbl, stretch=1)

        # Timestamp
        self.captured_lbl = QLabel("")
        self.captured_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.captured_lbl.setStyleSheet(f"color:{TEXT_MUT}; font-size:10px; padding-right:4px;")
        ly.addWidget(self.captured_lbl)

        # Thumbnail strip
        scroll = QScrollArea()
        scroll.setFixedHeight(74); scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background:transparent; border:none;")
        thumb_w = QWidget(); thumb_w.setStyleSheet("background:transparent;")
        self.thumb_layout = QHBoxLayout(thumb_w)
        self.thumb_layout.setContentsMargins(0, 0, 0, 0); self.thumb_layout.setSpacing(6)
        self.thumb_labels = []
        for i in range(7):
            t = QLabel(); t.setFixedSize(62, 62)
            t.setAlignment(Qt.AlignmentFlag.AlignCenter)
            t.setStyleSheet(f"background:{BG_CARD}; border:1px solid {BORDER}; border-radius:6px; color:{TEXT_MUT};")
            t.setCursor(Qt.CursorShape.PointingHandCursor)
            t.mousePressEvent = lambda e, idx=i: self._thumb_click(idx)
            self.thumb_layout.addWidget(t); self.thumb_labels.append(t)
        self.thumb_layout.addStretch()
        scroll.setWidget(thumb_w); ly.addWidget(scroll)

        # Action label + counter
        ar = QHBoxLayout()
        self.counter_lbl = QLabel("0 of 0 waiting")
        self.counter_lbl.setStyleSheet(f"color:{TEXT_MUT}; font-size:11px; letter-spacing:0.5px;")
        self.action_hint = QLabel("REVIEW ACTIONS")
        self.action_hint.setStyleSheet(f"color:{TEXT_MUT}; font-size:10px; letter-spacing:1px;")
        ar.addWidget(self.counter_lbl); ar.addStretch(); ar.addWidget(self.action_hint)
        ly.addLayout(ar)

        # Action buttons
        self.btn_accept  = self._action_btn("ACCEPT",  "A",   GREEN,      "✓")
        self.btn_reject  = self._action_btn("REJECT",  "R",   RED,        "✕")
        self.btn_rework  = self._action_btn("REWORK",  "W",   ORANGE,     "⟳")
        self.btn_skip    = self._action_btn("SKIP",    "S",   TEXT_SEC,   "⊳")
        self.btn_refresh = self._action_btn("REFRESH", "r",   BLUE,       "↻")
        self.btn_delete  = self._action_btn("DELETE",  "Del", "#B91C1C",  "🗑")
        self.btn_accept.clicked.connect(self.do_accept)
        self.btn_reject.clicked.connect(self.do_reject)
        self.btn_rework.clicked.connect(self.do_rework)
        self.btn_skip.clicked.connect(self.do_skip)
        self.btn_refresh.clicked.connect(lambda: self.refresh_state())
        self.btn_delete.clicked.connect(self.do_delete)
        act_row = QHBoxLayout(); act_row.setSpacing(5)
        for b in [self.btn_accept, self.btn_reject, self.btn_rework,
                  self.btn_skip, self.btn_refresh, self.btn_delete]:
            act_row.addWidget(b)
        ly.addLayout(act_row)
        return panel

    def _action_btn(self, label, key, color, icon):
        btn = QPushButton(f"{icon}\n{label}\n[{key}]")
        btn.setFixedHeight(62); btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        btn.setStyleSheet(f"""
            QPushButton{{background:{BG_CARD};color:{color};border:1px solid {color}44;border-radius:8px;font-size:10px;font-weight:bold;padding:4px;}}
            QPushButton:hover{{background:{color}1A;border-color:{color};}}
            QPushButton:pressed{{background:{color}33;}}
        """)
        return btn

    # ------------------------------------------------------------------
    # RIGHT: CONTROLS + STATS
    # ------------------------------------------------------------------
    def _build_controls_panel(self):
        panel = self._panel()
        ly = QVBoxLayout(panel); ly.setContentsMargins(14, 14, 14, 14); ly.setSpacing(10)

        # Controls header
        ly.addWidget(self._section_hdr("⚙  CONTROLS"))

        for field_lbl, attr, default in [
            ("OPERATOR ID", "operator_input", "OP-101"),
            ("BATCH ID",    "batch_input",    "BATCH-001"),
        ]:
            lbl = QLabel(field_lbl)
            lbl.setStyleSheet(f"color:{TEXT_MUT}; font-size:9px; letter-spacing:1px;")
            inp = QLineEdit(default)
            setattr(self, attr, inp)
            ly.addWidget(lbl); ly.addWidget(inp)

        # Reject reason
        rl = QLabel("REJECT REASON"); rl.setStyleSheet(f"color:{TEXT_MUT}; font-size:9px; letter-spacing:1px;")
        self.reason_combo = QComboBox()
        self.reason_combo.addItems(["Surface Scratch", "Scratch", "Crack", "Size Issue", "Polish", "Shape", "Other"])
        ly.addWidget(rl); ly.addWidget(self.reason_combo)

        # Report button
        self.btn_report = QPushButton("  📊  Generate Shift Report")
        self.btn_report.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_report.setStyleSheet(f"""
            QPushButton{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #7C3AED,stop:1 {ACCENT});color:white;font-size:12px;font-weight:bold;border-radius:8px;border:none;padding:10px;}}
            QPushButton:hover{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #8B5CF6,stop:1 #6366F1);}}
        """)
        self.btn_report.clicked.connect(self.generate_report)
        ly.addWidget(self.btn_report)

        ly.addWidget(self._hsep())

        # Stats header
        ly.addWidget(self._section_hdr("⚡  TODAY'S LIVE STATS"))

        # Stat rows
        self.stat_waiting,  r1 = self._stat_row("⏱", "Waiting Count",  TEXT_PRI)
        self.stat_accepted, r2 = self._stat_row("✓", "Accepted Count", GREEN)
        self.stat_rejected, r3 = self._stat_row("✕", "Rejected Count", RED)
        self.stat_rework,   r4 = self._stat_row("⟳", "Rework Count",   ORANGE)
        for r in [r1, r2, r3, r4]: ly.addLayout(r)

        # Acceptance rate card
        rate_card = QWidget()
        rate_card.setStyleSheet(f"""
            background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #1E3A8A,stop:1 #1D4ED8);
            border-radius:10px; border:1px solid #2563EB;
        """)
        rcl = QVBoxLayout(rate_card); rcl.setContentsMargins(12, 10, 12, 10)
        rl2 = QLabel("ACCEPTANCE RATE"); rl2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rl2.setStyleSheet("color:#93C5FD; font-size:9px; letter-spacing:1.5px; font-weight:bold; background:transparent;")
        self.stat_rate = QLabel("0.0%"); self.stat_rate.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stat_rate.setStyleSheet("color:white; font-size:28px; font-weight:bold; background:transparent;")
        rcl.addWidget(rl2); rcl.addWidget(self.stat_rate)
        ly.addWidget(rate_card)

        # Last refreshed
        self.stat_last_refresh = QLabel("LAST REFRESHED  —")
        self.stat_last_refresh.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.stat_last_refresh.setStyleSheet(f"color:{TEXT_MUT}; font-size:10px;")
        ly.addWidget(self.stat_last_refresh)
        ly.addStretch()
        return panel

    def _section_hdr(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color:{ACCENT}; font-size:11px; font-weight:bold; letter-spacing:1.5px;")
        return lbl

    def _stat_row(self, icon, label, color):
        row = QHBoxLayout()
        lbl = QLabel(f"{icon}  {label}"); lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:12px;")
        val = QLabel("0"); val.setStyleSheet(f"font-size:18px; font-weight:bold; color:{color};")
        val.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        row.addWidget(lbl); row.addStretch(); row.addWidget(val)
        return val, row

    # ------------------------------------------------------------------
    # BOTTOM STATUS BAR
    # ------------------------------------------------------------------
    def _build_statusbar(self):
        bar = QWidget(); bar.setFixedHeight(38)
        bar.setStyleSheet(f"background:{BG_PANEL}; border-top:1px solid {BORDER};")
        ly = QHBoxLayout(bar); ly.setContentsMargins(12, 0, 12, 0); ly.setSpacing(6)

        for label, slot_fn in [("⊕  Quick Scan", self.scan_cameras), ("⬡  Camera Picker", None)]:
            btn = QPushButton(label); btn.setFixedHeight(26); btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            btn.setStyleSheet(f"""
                QPushButton{{background:{BG_CARD};color:{TEXT_SEC};border:1px solid {BORDER};border-radius:5px;padding:0 12px;font-size:11px;font-weight:bold;}}
                QPushButton:hover{{color:{TEXT_PRI};border-color:{ACCENT};}}
            """)
            if slot_fn: btn.clicked.connect(slot_fn)
            ly.addWidget(btn)

        ly.addSpacing(10)
        self.cam_status_lbl = QLabel("CAMERA LINKED")
        self.cam_status_lbl.setStyleSheet(f"color:{GREEN}; font-size:10px; font-weight:bold; letter-spacing:0.5px;")
        ly.addWidget(self.cam_status_lbl)
        storage = QLabel("  STORAGE OK")
        storage.setStyleSheet(f"color:{GREEN}; font-size:10px; font-weight:bold; letter-spacing:0.5px;")
        ly.addWidget(storage)
        ly.addStretch()

        self.mini_acc = QLabel("Acc  0"); self.mini_acc.setStyleSheet(f"color:{GREEN}; font-size:11px; font-weight:bold;")
        self.mini_rej = QLabel("Rej  0");  self.mini_rej.setStyleSheet(f"color:{RED}; font-size:11px; font-weight:bold;")
        self.mini_rew = QLabel("Rew  0"); self.mini_rew.setStyleSheet(f"color:{ORANGE}; font-size:11px; font-weight:bold;")
        for w in [self.mini_acc, self.mini_rej, self.mini_rew]:
            ly.addWidget(w); ly.addSpacing(8)

        self.status_chip = QLabel("STATUS  STABLE")
        self.status_chip.setStyleSheet(f"background:{GREEN}22; color:{GREEN}; border:1px solid {GREEN}55; border-radius:4px; padding:2px 10px; font-size:10px; font-weight:bold;")
        ly.addWidget(self.status_chip)
        return bar

    # ------------------------------------------------------------------
    # Camera button style
    # ------------------------------------------------------------------
    def _cam_btn_style(self, slot, active=False):
        if active:
            return (f"QPushButton{{padding:6px 10px;font-weight:bold;border-radius:6px;font-size:11px;"
                    f"background-color:{slot['color']};color:white;border:1px solid white;}}"
                    f"QPushButton:hover{{background-color:{slot['hover']};}}")
        return (f"QPushButton{{padding:6px 10px;font-weight:bold;border-radius:6px;font-size:11px;"
                f"background-color:{BG_CARD};color:{TEXT_SEC};border:1px solid {BORDER};}}"
                f"QPushButton:hover{{background-color:{slot['color']}22;color:white;border-color:{slot['color']};}}")

    # ------------------------------------------------------------------
    # Clock
    # ------------------------------------------------------------------
    def _update_clock(self):
        self.clock_display.setText(datetime.now().strftime("%H:%M:%S"))

    # ------------------------------------------------------------------
    # REC blink
    # ------------------------------------------------------------------
    def _blink_rec(self):
        self._rec_blink = not self._rec_blink
        c = RED if self._rec_blink else "transparent"
        self.rec_indicator.setStyleSheet(f"color:{c}; font-size:10px; font-weight:bold; padding:1px 4px;")

    # ------------------------------------------------------------------
    # Thumbnails
    # ------------------------------------------------------------------
    def _thumb_click(self, idx):
        if idx < len(self.waiting_images):
            self.current_image_index = idx
            self.load_image()
            self._update_thumb_highlights()

    def _update_thumb_highlights(self):
        for i, t in enumerate(self.thumb_labels):
            if i < len(self.waiting_images):
                active = (i == self.current_image_index)
                border_color = ACCENT if active else BORDER
                border_width = "2px" if active else "1px"
                t.setStyleSheet(f"background:{BG_CARD}; border:{border_width} solid {border_color}; border-radius:6px; color:{TEXT_MUT};")
            else:
                t.setStyleSheet(f"background:{BG_CARD}; border:1px solid {BORDER}; border-radius:6px; color:{TEXT_MUT};")

    def _update_thumbnails(self):
        for i, t in enumerate(self.thumb_labels):
            if i < len(self.waiting_images):
                img_path = os.path.join(self.classifier.waiting_dir, self.waiting_images[i])
                px = QPixmap(img_path)
                if not px.isNull():
                    t.setPixmap(px.scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                          Qt.TransformationMode.FastTransformation))
                    t.setText("")
                    t.setToolTip(self.waiting_images[i])
                else:
                    t.clear(); t.setText("?")
            else:
                t.clear(); t.setText("")
        self._update_thumb_highlights()

    # ------------------------------------------------------------------
    # Camera init & switching
    # ------------------------------------------------------------------
    def init_camera(self, index=None):
        if self.cap is not None:
            self.cap.release(); self.cap = None

        slot = CAMERA_SLOTS[self.active_slot]

        if "Microscope USB" in slot["label"]:
            cap = open_camera_by_name("USB Video", width=1280, height=720)
            if cap is not None:
                self.cap = cap
                self.camera_id = "USB Video"
                self.cam_source_lbl.setText("SOURCE:  Microscope USB  (USB Video — Direct)")
                self._set_cam_status(True)
                return
            target = 1
        else:
            target = index if index is not None else self.camera_id

        self.camera_id = target
        self.cap = cv2.VideoCapture(target, cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH,  1280)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

        if not self.cap.isOpened():
            self.cam_lbl.setText(
                f"⚠  Camera not found (index {target})\n\n"
                "• Is the USB cable firmly plugged in?\n"
                "• Try a different USB port\n"
                "• Unplug, wait 3 s, replug\n"
                "• Click Scan / Quick Scan"
            )
            self.cam_lbl.setStyleSheet(
                f"background:#1A0A0A; color:#EF9A9A; border:2px solid {RED}; "
                f"border-radius:8px; font-size:12px; padding:18px;"
            )
            self._set_cam_status(False)
        else:
            self._set_cam_status(True)

    def _set_cam_status(self, ok):
        if ok:
            self.cam_status_lbl.setText("CAMERA LINKED")
            self.cam_status_lbl.setStyleSheet(f"color:{GREEN}; font-size:10px; font-weight:bold; letter-spacing:0.5px;")
            self.status_chip.setText("STATUS  STABLE")
            self.status_chip.setStyleSheet(f"background:{GREEN}22; color:{GREEN}; border:1px solid {GREEN}55; border-radius:4px; padding:2px 10px; font-size:10px; font-weight:bold;")
        else:
            self.cam_status_lbl.setText("NO CAMERA")
            self.cam_status_lbl.setStyleSheet(f"color:{RED}; font-size:10px; font-weight:bold; letter-spacing:0.5px;")
            self.status_chip.setText("STATUS  ERROR")
            self.status_chip.setStyleSheet(f"background:{RED}22; color:{RED}; border:1px solid {RED}55; border-radius:4px; padding:2px 10px; font-size:10px; font-weight:bold;")

    def switch_camera_slot(self, slot_index):
        self.active_slot = slot_index
        slot = CAMERA_SLOTS[slot_index]
        self.camera_id = slot["index"]
        for i, btn in enumerate(self.cam_slot_buttons):
            btn.setChecked(i == slot_index)
            btn.setStyleSheet(self._cam_btn_style(CAMERA_SLOTS[i], active=(i == slot_index)))
        self.cam_source_lbl.setText(f"SOURCE:  {slot['label']}  (index {slot['index']})")
        self.cam_lbl.setStyleSheet(f"background:#090B11; color:{TEXT_MUT}; border:1px solid {BORDER}; border-radius:8px;")
        self.cam_lbl.setText("Switching camera…")
        self.init_camera(slot["index"])

    # ------------------------------------------------------------------
    # Background camera scanner
    # ------------------------------------------------------------------
    def scan_cameras(self):
        self.btn_scan.setEnabled(False); self.btn_scan.setText("…")
        self._scanner_thread = QThread()
        self._scanner_worker = CameraScanner()
        self._scanner_worker.moveToThread(self._scanner_thread)
        self._scanner_thread.started.connect(self._scanner_worker.run)
        self._scanner_worker.found.connect(self._on_scan_done)
        self._scanner_worker.found.connect(self._scanner_thread.quit)
        self._scanner_thread.finished.connect(self._scanner_thread.deleteLater)
        self._scanner_thread.start()

    def _on_scan_done(self, results):
        self.btn_scan.setEnabled(True); self.btn_scan.setText("⊕")
        if not results:
            QMessageBox.warning(self, "No Cameras", "No cameras found.\n\n• Check USB cable\n• Try a different port"); return
        lines = ["Cameras detected:\n"]
        lines += [f"  Index {i}  —  {r}{'  ← webcam' if i==0 else ''}" for i, r in results]
        lines += ["", "Update CAMERA_SLOTS in dashboard.py → restart."]
        QMessageBox.information(self, "Camera Scan Results", "\n".join(lines))

    # ------------------------------------------------------------------
    # Frame update
    # ------------------------------------------------------------------
    def update_frame(self):
        if self.cap is None or not self.cap.isOpened(): return
        ret, frame = self.cap.read()
        if not ret: return
        self.current_frame = frame.copy()

        s = self.cam_lbl.styleSheet()
        if "1A0A0A" in s or "EF9A9A" in s:
            self.cam_lbl.setStyleSheet(f"background:#090B11; color:{TEXT_MUT}; border:1px solid {BORDER}; border-radius:8px;")

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        qt = QImage(rgb.data, w, h, ch * w, QImage.Format.Format_RGB888)
        px = QPixmap.fromImage(qt).scaled(
            self.cam_lbl.width(), self.cam_lbl.height(),
            Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.FastTransformation)
        self.cam_lbl.setPixmap(px)

        if self.max_feed_win is not None and self.max_feed_win.isVisible():
            full_px = QPixmap.fromImage(qt)
            self.max_feed_win.update_frame(full_px)

    def toggle_maximized_feed(self):
        if self.max_feed_win is None:
            self.max_feed_win = MaximizedFeedWindow()
        if self.max_feed_win.isVisible():
            self.max_feed_win.hide()
        else:
            self.max_feed_win.showMaximized()

    # ------------------------------------------------------------------
    # Capture
    # ------------------------------------------------------------------
    def capture_image(self):
        if self.current_frame is None: return
        try:
            self.cam_lbl.setStyleSheet(f"background:white; border:3px solid {GREEN}; border-radius:8px;")
            QTimer.singleShot(150, lambda: self.cam_lbl.setStyleSheet(
                f"background:#090B11; color:{TEXT_MUT}; border:1px solid {BORDER}; border-radius:8px;"))
            uid = uuid.uuid4().hex[:10]
            ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
            name = f"IMG_{ts}_{uid}.jpg"
            path = os.path.join(self.wait_dir, name)
            cv2.imwrite(path, self.current_frame)
            db_handler.insert_capture(name, path)
            excel_writer.add_capture(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), name, path)
            self.refresh_state()
        except Exception as e:
            QMessageBox.critical(self, "Capture Error",
                f"Could not capture image!\nMake sure inspection_log.xlsx is NOT open in Excel.\n\nError: {e}")

    # ------------------------------------------------------------------
    # Input focus guard
    # ------------------------------------------------------------------
    def is_input_focused(self):
        return isinstance(self.focusWidget(), (QLineEdit, QComboBox, QSpinBox))

    # ------------------------------------------------------------------
    # State refresh
    # ------------------------------------------------------------------
    def refresh_state(self, *args):
        try:
            files = os.listdir(self.classifier.waiting_dir)
        except FileNotFoundError:
            files = []
        new_waiting = sorted([f for f in files if f.lower().endswith((".png", ".jpg", ".jpeg"))])
        curr = self.get_current_image()
        self.waiting_images = new_waiting
        self.current_image_index = self.waiting_images.index(curr) if (curr and curr in self.waiting_images) else 0

        if self.waiting_images:
            if self.current_image_index >= len(self.waiting_images): self.current_image_index = 0
            self.load_image()
        else:
            self.image_lbl.clear(); self.image_lbl.setText("No Images Waiting")
            self.counter_lbl.setText("0 of 0 waiting")
            self.count_badge.setText("0")
            self.captured_lbl.setText("")
            self.action_hint.setText("REVIEW ACTIONS")

        self._update_thumbnails()

        try:
            s = db_handler.get_stats()
            self.stat_waiting.setText(str(s["waiting"]))
            self.stat_accepted.setText(str(s["accepted"]))
            self.stat_rejected.setText(str(s["rejected"]))
            self.stat_rework.setText(str(s.get("rework", 0)))
            self.stat_rate.setText(f"{s['rate']}%")
            self.mini_acc.setText(f"Acc  {s['accepted']}")
            self.mini_rej.setText(f"Rej  {s['rejected']}")
            self.mini_rew.setText(f"Rew  {s.get('rework', 0)}")
        except Exception:
            pass

        self.stat_last_refresh.setText(f"LAST REFRESHED  {datetime.now().strftime('%H:%M:%S')}")

        # Flash badge
        orig = self.badge_lbl.styleSheet()
        self.badge_lbl.setStyleSheet(
            f"background:{ACCENT}33; color:{TEXT_PRI}; border:1px solid {ACCENT}; border-radius:12px; padding:4px 14px; font-size:12px; font-weight:bold;")
        QTimer.singleShot(250, lambda: self.badge_lbl.setStyleSheet(orig))

        n = len(self.waiting_images)
        self.nav_status_lbl.setText(f"● WAITING FOR REVIEW  ({n})" if n > 0 else "✓ ALL CLEAR")

    def load_image(self):
        if not self.waiting_images: return
        name = self.waiting_images[self.current_image_index]
        path = os.path.join(self.classifier.waiting_dir, name)
        px = QPixmap(path)
        if not px.isNull():
            self.image_lbl.setPixmap(px.scaled(
                self.image_lbl.width(), self.image_lbl.height(),
                Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        n = len(self.waiting_images)
        self.counter_lbl.setText(f"{self.current_image_index + 1} of {n} waiting")
        self.count_badge.setText(str(n))
        self.action_hint.setText(f"REVIEW ACTIONS — {self.current_image_index + 1} OF {n} SELECTED")
        parts = name.replace("IMG_", "").split("_")
        try:
            ts = datetime.strptime(f"{parts[0]}_{parts[1]}", "%Y%m%d_%H%M%S")
            self.captured_lbl.setText(f"Captured: {ts.strftime('%H:%M:%S')}")
        except Exception:
            self.captured_lbl.setText(f"File: {name}")

    def get_current_image(self):
        if not self.waiting_images: return None
        if self.current_image_index < len(self.waiting_images):
            return self.waiting_images[self.current_image_index]
        return None

    # ------------------------------------------------------------------
    # Action handlers
    # ------------------------------------------------------------------
    def do_accept(self):
        img = self.get_current_image()
        if not img: return
        try:
            if self.classifier.accept(img, self.operator_input.text().strip() or "Unknown",
                                           self.batch_input.text().strip() or "Unknown"):
                self.refresh_state()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not accept image!\n\nError: {e}")

    def do_reject(self):
        img = self.get_current_image()
        if not img: return
        try:
            if self.classifier.reject(img, self.operator_input.text().strip() or "Unknown",
                                           self.batch_input.text().strip() or "Unknown",
                                           self.reason_combo.currentText()):
                self.refresh_state()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not reject image!\n\nError: {e}")

    def do_rework(self):
        img = self.get_current_image()
        if not img: return
        try:
            if self.classifier.rework(img, self.operator_input.text().strip() or "Unknown",
                                           self.batch_input.text().strip() or "Unknown"):
                self.refresh_state()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not move to Rework!\n\nError: {e}")

    def do_skip(self):
        if self.waiting_images:
            self.current_image_index = (self.current_image_index + 1) % len(self.waiting_images)
            self.load_image(); self._update_thumb_highlights()

    def do_delete(self):
        img = self.get_current_image()
        if not img: return
        if QMessageBox.question(self, "Confirm Delete", f"Permanently delete  {img}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                if self.classifier.delete(img): self.refresh_state()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not delete!\n\nError: {e}")

    def generate_report(self):
        stats = db_handler.get_stats()
        path  = excel_writer.generate_shift_report(stats, batch_data=db_handler.get_batch_stats())
        QMessageBox.information(self, "Report Generated", f"Shift report saved to:\n{path}")

    # ------------------------------------------------------------------
    # Clean shutdown
    # ------------------------------------------------------------------
    def closeEvent(self, event):
        for t in [self.camera_timer, self.state_timer, self.clock_timer, self.rec_timer]:
            t.stop()
        if self.cap: self.cap.release(); self.cap = None
        if self._scanner_thread and self._scanner_thread.isRunning():
            self._scanner_thread.quit(); self._scanner_thread.wait(2000)
        if self.max_feed_win: self.max_feed_win.close()
        super().closeEvent(event)
