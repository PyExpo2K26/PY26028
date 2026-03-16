"""
PyXcel — Home Panel
Welcome screen with file upload, quick actions, and system status.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QFrame, QGridLayout,
    QSizePolicy, QSpacerItem
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent


class HomePanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAcceptDrops(True)  # Enable drag & drop
        self._build_ui()
        self._start_status_check()

    # ── Build UI ────────────────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(24)

        # ── Header ──
        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(6)

        badge = QLabel("AI-POWERED SPREADSHEET SYSTEM")
        badge.setObjectName("badge")
        badge.setFixedHeight(24)
        badge.setAlignment(Qt.AlignLeft)
        badge.setStyleSheet("""
            background-color: #1e2035;
            color: #7c83ff;
            border-radius: 12px;
            padding: 3px 12px;
            font-size: 10px;
            font-weight: bold;
            letter-spacing: 1px;
        """)

        title = QLabel("Welcome to PyXcel")
        title.setObjectName("panel_title")
        font = QFont()
        font.setPointSize(24)
        font.setBold(True)
        title.setFont(font)

        subtitle = QLabel(
            "Replace macros with natural language  ·  "
            "Generate formulas  ·  Clean data  ·  Chat with your spreadsheet"
        )
        subtitle.setObjectName("panel_subtitle")
        subtitle.setStyleSheet("color: #555; font-size: 13px;")

        header_layout.addWidget(badge)
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        layout.addWidget(header)

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

        layout.addStretch()

    # ── Drop Zone ────────────────────────────────────────────
    def _build_drop_zone(self):
        zone = QFrame()
        zone.setObjectName("drop_zone")
        zone.setFixedHeight(180)
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

        hint = QLabel("or click to browse  ·  supports .xlsx .xls")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet(
            "color: #3a3d5e; font-size: 12px; background: transparent;"
        )

        browse_btn = QPushButton("Browse Files")
        browse_btn.setFixedWidth(160)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse_file)

        layout.addWidget(icon)
        layout.addWidget(self.drop_label)
        layout.addWidget(hint)
        layout.addWidget(browse_btn, alignment=Qt.AlignCenter)

        # Make entire zone clickable
        zone.mousePressEvent = lambda e: self._browse_file()

        return zone

    # ── Quick Actions ────────────────────────────────────────
    def _build_quick_actions(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        actions = [
            ("⚙️",  "Macro Replacement",  "Replace VBA macros with\nnatural language",  2),
            ("🧮",  "Formula Generator",  "Generate Excel formulas\nfrom plain English",  3),
            ("🧹",  "Data Cleaner",       "Clean & transform data\nwith AI instructions", 4),
            ("💬",  "Chat with Data",     "Ask questions about\nyour spreadsheet",        5),
            ("📈",  "KPI Cards",          "Auto-detect business\nKPIs from any sheet",    6),
        ]

        for icon, title, desc, panel_index in actions:
            card = self._action_card(icon, title, desc, panel_index)
            layout.addWidget(card)

        return widget

    def _action_card(self, icon, title, desc, panel_index):
        card = QFrame()
        card.setObjectName("card")
        card.setCursor(Qt.PointingHandCursor)
        card.setFixedHeight(130)
        card.setStyleSheet("""
            QFrame#card {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-radius: 12px;
                padding: 16px;
            }
            QFrame#card:hover {
                border-color: #7c83ff;
                background-color: #1e2035;
            }
        """)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)

        icon_label = QLabel(icon)
        icon_label.setStyleSheet(
            "font-size: 22px; background: transparent; border: none;"
        )

        title_label = QLabel(title)
        title_label.setStyleSheet(
            "font-size: 13px; font-weight: bold; "
            "color: #e0e0e0; background: transparent; border: none;"
        )

        desc_label = QLabel(desc)
        desc_label.setStyleSheet(
            "font-size: 11px; color: #555; "
            "background: transparent; border: none;"
        )

        layout.addWidget(icon_label)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addStretch()

        # Click to switch panel
        card.mousePressEvent = lambda e, i=panel_index: \
            self.main_window.switch_panel(i)

        return card

    # ── Status Cards ─────────────────────────────────────────
    def _build_status_cards(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        # Ollama status
        self.ollama_card  = self._status_card("Ollama",       "Checking...", "⬤")
        self.model_card   = self._status_card("LLaMA Model",  "Checking...", "⬤")
        self.file_card    = self._status_card("File Loaded",  "None",        "📄")
        self.offline_card = self._status_card("Offline Mode", "Active ✓",   "🔒")

        layout.addWidget(self.ollama_card["frame"])
        layout.addWidget(self.model_card["frame"])
        layout.addWidget(self.file_card["frame"])
        layout.addWidget(self.offline_card["frame"])
        layout.addStretch()

        # Set offline card as always green
        self.offline_card["value"].setStyleSheet(
            "font-size: 14px; font-weight: bold; "
            "color: #4caf81; background: transparent;"
        )

        return widget

    def _status_card(self, title, value, icon):
        frame = QFrame()
        frame.setObjectName("card")
        frame.setFixedSize(190, 90)
        frame.setStyleSheet("""
            QFrame {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-radius: 12px;
            }
        """)

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)

        title_row = QHBoxLayout()
        icon_lbl  = QLabel(icon)
        icon_lbl.setStyleSheet(
            "font-size: 12px; color: #555; background: transparent;"
        )
        title_lbl = QLabel(title)
        title_lbl.setStyleSheet(
            "font-size: 11px; color: #555; background: transparent;"
        )
        title_row.addWidget(icon_lbl)
        title_row.addWidget(title_lbl)
        title_row.addStretch()

        value_lbl = QLabel(value)
        value_lbl.setStyleSheet(
            "font-size: 14px; font-weight: bold; "
            "color: #e0e0e0; background: transparent;"
        )

        layout.addLayout(title_row)
        layout.addWidget(value_lbl)

        return {"frame": frame, "value": value_lbl}

    # ── Status Check ─────────────────────────────────────────
    def _start_status_check(self):
        self._check_status()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._check_status)
        self.timer.start(15000)  # re-check every 15 seconds

    def _check_status(self):
        from core.ollama_client import is_ollama_running, is_model_available

        # Ollama
        if is_ollama_running():
            self.ollama_card["value"].setText("Running ✓")
            self.ollama_card["value"].setStyleSheet(
                "font-size: 14px; font-weight: bold; "
                "color: #4caf81; background: transparent;"
            )
        else:
            self.ollama_card["value"].setText("Offline ✗")
            self.ollama_card["value"].setStyleSheet(
                "font-size: 14px; font-weight: bold; "
                "color: #f44336; background: transparent;"
            )

        # Model
        if is_model_available("llama3.1"):
            self.model_card["value"].setText("Ready ✓")
            self.model_card["value"].setStyleSheet(
                "font-size: 14px; font-weight: bold; "
                "color: #4caf81; background: transparent;"
            )
        else:
            self.model_card["value"].setText("Not Pulled ✗")
            self.model_card["value"].setStyleSheet(
                "font-size: 14px; font-weight: bold; "
                "color: #f44336; background: transparent;"
            )

        # File
        if self.main_window.current_file:
            name = os.path.basename(self.main_window.current_file)
            self.file_card["value"].setText(name[:22])
            self.file_card["value"].setStyleSheet(
                "font-size: 13px; font-weight: bold; "
                "color: #7c83ff; background: transparent;"
            )
        else:
            self.file_card["value"].setText("None")
            self.file_card["value"].setStyleSheet(
                "font-size: 14px; font-weight: bold; "
                "color: #555; background: transparent;"
            )

    # ── File Handling ────────────────────────────────────────
    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        self.main_window.current_file = path
        self.main_window.load_file.__func__
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

    # ── Helper ───────────────────────────────────────────────
    def _section_label(self, text: str):
        label = QLabel(text)
        label.setStyleSheet(
            "color: #444; font-size: 10px; "
            "letter-spacing: 1px; padding-top: 8px;"
        )
        return label
