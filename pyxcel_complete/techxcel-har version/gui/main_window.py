"""
PyXcel — Main Window
Sidebar navigation + panel switcher.
"""
import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLabel, QPushButton, QStackedWidget, QFrame,
    QStatusBar, QFileDialog
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyXcel — AI Spreadsheet System")
        self.setMinimumSize(1200, 750)
        self.resize(1400, 860)
        self.current_file = None
        self._build_ui()
        self._check_ollama_status()

    # ── Build UI ────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.sidebar = self._build_sidebar()
        root_layout.addWidget(self.sidebar)

        self.stack = QStackedWidget()
        self.stack.setObjectName("content_area")
        root_layout.addWidget(self.stack)

        self._load_panels()

        self.status_bar = QStatusBar()
        self.status_bar.showMessage("Ready  |  No file loaded")
        self.setStatusBar(self.status_bar)

    # ── Sidebar ─────────────────────────────────────────────
    def _build_sidebar(self):
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(230)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # App title
        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setContentsMargins(16, 20, 16, 16)

        title = QLabel("PyXcel")
        title.setObjectName("app_title")
        font = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel("AI Spreadsheet System")
        subtitle.setObjectName("app_subtitle")

        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        layout.addWidget(title_widget)
        layout.addWidget(self._divider())

        # Ollama status
        self.ollama_indicator = QLabel("⬤  Checking Ollama...")
        self.ollama_indicator.setContentsMargins(16, 8, 16, 8)
        self.ollama_indicator.setStyleSheet("color:#555;font-size:11px;")
        layout.addWidget(self.ollama_indicator)
        layout.addWidget(self._divider())

        # ── Nav sections ──
        nav_items = [
            ("WORKSPACE", [
                ("🏠  Home",               0),
                ("📊  Spreadsheet View",   1),
            ]),
            ("AI FEATURES", [
                ("⚙️   Macro Replacement",  2),
                ("🧮  Formula Generator",  3),
                ("🧹  Data Cleaner",        4),
                ("💬  Chat with Data",      5),
                ("📈  KPI Cards",           6),
            ]),
            ("TOOLS", [
                ("🔢  Pivot Tables",        7),
                ("📉  Chart Creator",       8),
                ("📄  PDF Export",          9),
                ("🔀  File Merger",         10),
            ]),
        ]

        self.nav_buttons = []
        for section_name, items in nav_items:
            # Section label
            sec = QLabel(section_name)
            sec.setContentsMargins(16, 12, 16, 4)
            sec.setStyleSheet("color:#444;font-size:10px;letter-spacing:1px;")
            layout.addWidget(sec)

            for label, index in items:
                btn = QPushButton(label)
                btn.setObjectName("nav_btn")
                btn.setProperty("active", "false")
                btn.setCursor(Qt.PointingHandCursor)
                btn.clicked.connect(
                    lambda checked, i=index: self.switch_panel(i)
                )
                layout.addWidget(btn)
                self.nav_buttons.append((index, btn))

        layout.addStretch()
        layout.addWidget(self._divider())

        # File info
        self.file_label = QLabel("No file loaded")
        self.file_label.setContentsMargins(16, 10, 16, 4)
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color:#444;font-size:10px;")
        layout.addWidget(self.file_label)

        # Load file button
        load_btn = QPushButton("📂  Load Excel File")
        load_btn.setCursor(Qt.PointingHandCursor)
        load_btn.setStyleSheet("""
            QPushButton {
                margin: 8px;
                background-color: #1e2035;
                color: #7c83ff;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 9px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #252840;
                border-color: #7c83ff;
            }
        """)
        load_btn.clicked.connect(self.load_file)
        layout.addWidget(load_btn)

        spacer = QWidget(); spacer.setFixedHeight(10)
        layout.addWidget(spacer)

        return sidebar

    # ── Load Panels ─────────────────────────────────────────
    def _load_panels(self):
        from gui.panels.home_panel        import HomePanel
        from gui.panels.spreadsheet_panel import SpreadsheetPanel
        from gui.panels.macro_panel       import MacroPanel
        from gui.panels.formula_panel     import FormulaPanel
        from gui.panels.cleaner_panel     import CleanerPanel
        from gui.panels.chat_panel        import ChatPanel
        from gui.panels.kpi_panel         import KpiPanel
        from gui.panels.pivot_panel       import PivotPanel
        from gui.panels.chart_panel       import ChartPanel
        from gui.panels.pdf_panel         import PdfPanel
        from gui.panels.merger_panel      import MergerPanel

        self.home_panel        = HomePanel(self)
        self.spreadsheet_panel = SpreadsheetPanel(self)
        self.macro_panel       = MacroPanel(self)
        self.formula_panel     = FormulaPanel(self)
        self.cleaner_panel     = CleanerPanel(self)
        self.chat_panel        = ChatPanel(self)
        self.kpi_panel         = KpiPanel(self)
        self.pivot_panel       = PivotPanel(self)
        self.chart_panel       = ChartPanel(self)
        self.pdf_panel         = PdfPanel(self)
        self.merger_panel      = MergerPanel(self)

        for panel in [
            self.home_panel,        # 0
            self.spreadsheet_panel, # 1
            self.macro_panel,       # 2
            self.formula_panel,     # 3
            self.cleaner_panel,     # 4
            self.chat_panel,        # 5
            self.kpi_panel,         # 6
            self.pivot_panel,       # 7
            self.chart_panel,       # 8
            self.pdf_panel,         # 9
            self.merger_panel,      # 10
        ]:
            self.stack.addWidget(panel)

        self.switch_panel(0)

    # ── Panel Switcher ───────────────────────────────────────
    def switch_panel(self, index: int):
        self.stack.setCurrentIndex(index)
        for idx, btn in self.nav_buttons:
            btn.setProperty("active", "true" if idx == index else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    # ── File Loader ─────────────────────────────────────────
    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return

        self.current_file = path
        filename = os.path.basename(path)
        self.file_label.setText(f"📄 {filename}")
        self.status_bar.showMessage(f"Loaded: {filename}")
        self._notify_panels_file_loaded(path)
        self.switch_panel(1)

    def _notify_panels_file_loaded(self, path: str):
        panels = [
            self.spreadsheet_panel,
            self.macro_panel,
            self.formula_panel,
            self.cleaner_panel,
            self.chat_panel,
            self.kpi_panel,
            self.pivot_panel,
            self.chart_panel,
            self.pdf_panel,
            self.merger_panel,
        ]
        for panel in panels:
            if hasattr(panel, "on_file_loaded"):
                try:
                    panel.on_file_loaded(path)
                except Exception:
                    pass

    # ── Ollama Status ────────────────────────────────────────
    def _check_ollama_status(self):
        from core.ollama_client import is_ollama_running, is_model_available

        def check():
            if is_ollama_running():
                if is_model_available():
                    self.ollama_indicator.setText("⬤  LLaMA Ready")
                    self.ollama_indicator.setStyleSheet(
                        "color:#4caf81;font-size:11px;padding:0 16px;"
                    )
                else:
                    self.ollama_indicator.setText("⬤  Model Not Pulled")
                    self.ollama_indicator.setStyleSheet(
                        "color:#ff9800;font-size:11px;padding:0 16px;"
                    )
            else:
                self.ollama_indicator.setText("⬤  Ollama Offline")
                self.ollama_indicator.setStyleSheet(
                    "color:#f44336;font-size:11px;padding:0 16px;"
                )

        QTimer.singleShot(500, check)
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(check)
        self.status_timer.start(30000)

    # ── Helpers ─────────────────────────────────────────────
    def _divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(
            "background-color:#2a2d3e;max-height:1px;border:none;"
        )
        return line

    def set_status(self, message: str):
        self.status_bar.showMessage(message)
