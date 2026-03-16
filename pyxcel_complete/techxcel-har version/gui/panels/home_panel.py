"""
PyXcel — Home Panel
Welcome screen with file upload, quick actions, and system status.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QFrame,
    QSizePolicy, QScrollArea
)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent


# ── Background thread for Ollama checks (never blocks the UI) ────────────────
class _StatusChecker(QThread):
    done = Signal(bool, bool)   # (ollama_running, model_available)

    def run(self):
        from core.ollama_client import is_ollama_running, is_model_available
        running  = is_ollama_running()
        model_ok = is_model_available("llama3.1") if running else False
        self.done.emit(running, model_ok)


class HomePanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAcceptDrops(True)
        self._build_ui()
        self._start_status_check()

    # ── Build UI ─────────────────────────────────────────────
    def _build_ui(self):
        # Outer layout — zero margin, holds only the scroll area
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        # Inner content widget
        content = QWidget()
        content.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(24)

        # ── Header ──
        badge = QLabel("AI-POWERED SPREADSHEET SYSTEM")
        badge.setFixedHeight(24)
        badge.setAlignment(Qt.AlignLeft)
        badge.setStyleSheet("""
            background-color: #1e2035; color: #7c83ff;
            border-radius: 12px; padding: 3px 12px;
            font-size: 10px; font-weight: bold; letter-spacing: 1.5px;
        """)

        title = QLabel("Welcome to PyXcel")
        font = QFont(); font.setPointSize(24); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel(
            "Replace macros with natural language  ·  "
            "Generate formulas  ·  Clean data  ·  Chat with your spreadsheet"
        )
        subtitle.setStyleSheet("color: #555; font-size: 13px; letter-spacing: 0.2px;")
        subtitle.setWordWrap(True)

        layout.addWidget(badge)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        # ── Drop Zone ──
        self.drop_zone = self._build_drop_zone()
        layout.addWidget(self.drop_zone)

        # ── Quick Actions ──
        layout.addWidget(self._section_label("QUICK ACTIONS"))
        layout.addWidget(self._build_quick_actions())

        # ── System Status ──
        layout.addWidget(self._section_label("SYSTEM STATUS"))
        self.status_grid = self._build_status_cards()
        layout.addWidget(self.status_grid)

        layout.addSpacing(24)

        scroll.setWidget(content)
        outer.addWidget(scroll)

    # ── Drop Zone ────────────────────────────────────────────
    def _build_drop_zone(self):
        zone = QFrame()
        zone.setObjectName("drop_zone")
        zone.setFixedHeight(190)
        zone.setCursor(Qt.PointingHandCursor)
        zone.setStyleSheet("""
            QFrame#drop_zone {
                background-color: #1a1d2e;
                border: 2px dashed #2a2d3e;
                border-radius: 12px;
            }
            QFrame#drop_zone:hover {
                border-color: #7c83ff;
                background-color: #1e2035;
            }
        """)

        layout = QVBoxLayout(zone)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(10)

        icon = QLabel("📂")
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("font-size: 36px; background: transparent;")

        self.drop_label = QLabel("Drop your Excel file here")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet(
            "color: #555; font-size: 15px; background: transparent;"
        )

        hint = QLabel("or click to browse  ·  supports .xlsx  .xls")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet(
            "color: #3a3d5e; font-size: 12px; background: transparent;"
        )

        browse_btn = QPushButton("Browse Files")
        browse_btn.setFixedWidth(160)
        browse_btn.setFixedHeight(40)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse_file)

        layout.addWidget(icon)
        layout.addWidget(self.drop_label)
        layout.addWidget(hint)
        layout.addWidget(browse_btn, alignment=Qt.AlignCenter)

        zone.mousePressEvent = lambda e: self._browse_file()
        return zone

    # ── Quick Actions ────────────────────────────────────────
    def _build_quick_actions(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        actions = [
            ("⚙️",  "Macro Replacement", "Replace VBA macros with natural language",    2),
            ("🧮",  "Formula Generator", "Generate Excel formulas from plain English",   3),
            ("🧹",  "Data Cleaner",      "Clean & transform data with AI instructions",  4),
            ("💬",  "Chat with Data",    "Ask questions about your spreadsheet",         5),
            ("📈",  "KPI Cards",         "Auto-detect business KPIs from any sheet",     6),
        ]

        for icon, title, desc, panel_index in actions:
            layout.addWidget(self._action_card(icon, title, desc, panel_index), stretch=1)

        return widget

    def _action_card(self, icon, title, desc, panel_index):
        card = QFrame()
        card.setObjectName("card")
        card.setCursor(Qt.PointingHandCursor)
        card.setMinimumHeight(140)
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        card.setStyleSheet("""
            QFrame#card {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-radius: 12px;
            }
            QFrame#card:hover {
                border-color: #7c83ff;
                background-color: #1e2035;
            }
        """)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)

        icon_lbl = QLabel(icon)
        icon_lbl.setStyleSheet("font-size: 24px; background: transparent; border: none;")

        title_lbl = QLabel(title)
        title_lbl.setStyleSheet(
            "font-size: 13px; font-weight: bold; color: #e0e0e0; "
            "background: transparent; border: none;"
        )
        title_lbl.setWordWrap(True)

        desc_lbl = QLabel(desc)
        desc_lbl.setStyleSheet(
            "font-size: 11px; color: #555; background: transparent; border: none; "
            "line-height: 1.4;"
        )
        desc_lbl.setWordWrap(True)

        layout.addWidget(icon_lbl)
        layout.addWidget(title_lbl)
        layout.addWidget(desc_lbl)
        layout.addStretch()

        card.mousePressEvent = lambda e, i=panel_index: self.main_window.switch_panel(i)
        return card

    # ── System Status Cards ──────────────────────────────────
    def _build_status_cards(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        self.ollama_card  = self._status_card("Ollama",       "Checking...", "⬤", "#555")
        self.model_card   = self._status_card("LLaMA Model",  "Checking...", "⬤", "#555")
        self.file_card    = self._status_card("File Loaded",  "None",        "📄", "#555")
        self.offline_card = self._status_card("Offline Mode", "Active ✓",   "🔒", "#4caf81")

        for card in [self.ollama_card, self.model_card,
                     self.file_card,   self.offline_card]:
            layout.addWidget(card["frame"], stretch=1)

        return widget

    def _status_card(self, title, value, icon, value_color):
        frame = QFrame()
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        frame.setMinimumWidth(150)
        frame.setStyleSheet("""
            QFrame {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-radius: 12px;
            }
        """)

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(8)

        # Icon + title row
        title_row = QHBoxLayout()
        title_row.setSpacing(6)

        icon_lbl = QLabel(icon)
        icon_lbl.setStyleSheet("font-size: 12px; color: #555; background: transparent;")
        icon_lbl.setFixedWidth(18)

        title_lbl = QLabel(title)
        title_lbl.setStyleSheet(
            "font-size: 11px; color: #888; background: transparent; "
            "letter-spacing: 0.4px; font-weight: bold;"
        )

        title_row.addWidget(icon_lbl)
        title_row.addWidget(title_lbl)
        title_row.addStretch()

        # Divider line
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setStyleSheet("background-color: #2a2d3e; max-height: 1px; border: none;")

        # Value label — large and clearly visible
        value_lbl = QLabel(value)
        value_lbl.setStyleSheet(
            f"font-size: 15px; font-weight: bold; "
            f"color: {value_color}; background: transparent; letter-spacing: 0.3px;"
        )
        value_lbl.setWordWrap(True)

        layout.addLayout(title_row)
        layout.addWidget(divider)
        layout.addWidget(value_lbl)

        return {"frame": frame, "value": value_lbl}

    # ── Status Check (background thread, never blocks UI) ────
    def _start_status_check(self):
        self._run_status_check()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._run_status_check)
        self.timer.start(20000)

    def _run_status_check(self):
        self._checker = _StatusChecker(self)
        self._checker.done.connect(self._on_status_result)
        self._checker.start()
        self._update_file_card()

    def _on_status_result(self, ollama_ok: bool, model_ok: bool):
        if ollama_ok:
            self._set_card(self.ollama_card, "Running ✓", "#4caf81")
        else:
            self._set_card(self.ollama_card, "Offline ✗", "#f44336")

        if model_ok:
            self._set_card(self.model_card, "Ready ✓", "#4caf81")
        else:
            self._set_card(self.model_card, "Not Pulled ✗", "#f44336")

    def _update_file_card(self):
        if self.main_window.current_file:
            name = os.path.basename(self.main_window.current_file)
            self._set_card(self.file_card, name[:24], "#7c83ff")
        else:
            self._set_card(self.file_card, "None", "#555")

    def _set_card(self, card, text, color):
        card["value"].setText(text)
        card["value"].setStyleSheet(
            f"font-size: 15px; font-weight: bold; "
            f"color: {color}; background: transparent; letter-spacing: 0.3px;"
        )

    # Keep legacy alias
    def _check_status(self):
        self._run_status_check()

    # ── File Handling ────────────────────────────────────────
    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        self.main_window.current_file = path
        self.main_window._notify_panels_file_loaded(path)
        filename = os.path.basename(path)
        self.drop_label.setText(f"✅  {filename} loaded")
        self.drop_label.setStyleSheet(
            "color: #4caf81; font-size: 14px; background: transparent;"
        )
        self.main_window.file_label.setText(f"📄 {filename}")
        self.main_window.status_bar.showMessage(f"Loaded: {filename}")
        self.main_window.switch_panel(1)

    # ── Drag & Drop ──────────────────────────────────────────
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().endswith((".xlsx", ".xls")):
                event.acceptProposedAction()
                self.drop_zone.setStyleSheet("""
                    QFrame#drop_zone {
                        background-color: #1e2035;
                        border: 2px dashed #7c83ff;
                        border-radius: 12px;
                    }
                """)

    def dragLeaveEvent(self, event):
        self.drop_zone.setStyleSheet("""
            QFrame#drop_zone {
                background-color: #1a1d2e;
                border: 2px dashed #2a2d3e;
                border-radius: 12px;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.endswith((".xlsx", ".xls")):
                self._load_file(path)
        self.drop_zone.setStyleSheet("""
            QFrame#drop_zone {
                background-color: #1a1d2e;
                border: 2px dashed #2a2d3e;
                border-radius: 12px;
            }
        """)

    def on_file_loaded(self, path: str):
        """Called by main_window when file is loaded from sidebar."""
        filename = os.path.basename(path)
        self.drop_label.setText(f"✅  {filename} loaded")
        self.drop_label.setStyleSheet(
            "color: #4caf81; font-size: 14px; background: transparent;"
        )
        self._update_file_card()

    # ── Helper ───────────────────────────────────────────────
    def _section_label(self, text: str):
        label = QLabel(text)
        label.setStyleSheet(
            "color: #444; font-size: 10px; letter-spacing: 2px; padding-top: 8px;"
        )
        return label
