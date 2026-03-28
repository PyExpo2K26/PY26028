"""
PyXcel — Main Window
Sidebar navigation + panel switcher + hamburger toggle with fade.
"""
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLabel, QPushButton, QStackedWidget, QFrame, QScrollArea,
    QStatusBar, QFileDialog, QApplication, QGraphicsOpacityEffect
)
from PySide6.QtCore import Qt, QTimer, QSettings, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QIcon, QPixmap


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyXcel — AI Spreadsheet System")
        self.setMinimumSize(1200, 750)
        self.resize(1400, 860)
        self.current_file = None
        self.theme = "dark"
        self.sidebar_visible = True
        self.settings = QSettings("KiTE Development Team", "PyXcel")

        # ── Logo path ───────────────────────────────────────
        self.logo_path = str(Path(__file__).resolve().parent / "assets" / "logo.png")

        # ── Window Title Bar Icon ───────────────────────────
        self.setWindowIcon(QIcon(self.logo_path))

        self._build_ui()
        self._setup_sidebar_animation()
        self._load_saved_theme()
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

        # ── Right side: topbar + content stack ─────────────
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        topbar = self._build_topbar()
        right_layout.addWidget(topbar)

        self.stack = QStackedWidget()
        self.stack.setObjectName("content_area")
        right_layout.addWidget(self.stack)

        root_layout.addWidget(right_widget)

        self._load_panels()

        self.status_bar = QStatusBar()
        self.status_bar.showMessage("Ready  |  No file loaded")
        self.setStatusBar(self.status_bar)

    # ── Topbar ───────────────────────────────────────────────
    def _build_topbar(self):
        topbar = QWidget()
        topbar.setObjectName("topbar")
        topbar.setFixedHeight(48)
        topbar.setStyleSheet("""
            #topbar {
                background-color: #1a1a2e;
                border-bottom: 1px solid #2a2a3e;
            }
        """)

        layout = QHBoxLayout(topbar)
        layout.setContentsMargins(10, 0, 16, 0)
        layout.setSpacing(8)

        # ── Hamburger button ────────────────────────────────
        self.hamburger_btn = QPushButton("☰")
        self.hamburger_btn.setFixedSize(36, 36)
        self.hamburger_btn.setCursor(Qt.PointingHandCursor)
        self.hamburger_btn.setToolTip("Toggle Sidebar")
        self.hamburger_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #cccccc;
                font-size: 20px;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #2a2a4a;
                color: white;
            }
            QPushButton:pressed {
                background-color: #3a3a5a;
            }
        """)
        self.hamburger_btn.clicked.connect(self.toggle_sidebar)
        layout.addWidget(self.hamburger_btn)

        # ── Quick action buttons (shown when sidebar is hidden) ──
        self.quick_actions = QWidget()
        qa_layout = QHBoxLayout(self.quick_actions)
        qa_layout.setContentsMargins(4, 0, 0, 0)
        qa_layout.setSpacing(5)

        quick_items = [
            (" Home",        0),
            (" Spreadsheet", 1),
            (" Formula",     2),
            (" Cleaner",     3),
            (" Chat",        4),
            (" KPI",         5),
            (" Pivot",       6),
            (" Charts",      7),
            (" PDF",         8),
            (" Merger",      9),
        ]

        for label, index in quick_items:
            btn = QPushButton(label)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background: #2a2a4a;
                    color: #cccccc;
                    font-size: 11px;
                    border: none;
                    border-radius: 5px;
                    padding: 4px 10px;
                }
                QPushButton:hover {
                    background-color: #3a3a6a;
                    color: white;
                }
                QPushButton:pressed {
                    background-color: #4a4a8a;
                }
            """)
            btn.clicked.connect(lambda checked, i=index: self.switch_panel(i))
            qa_layout.addWidget(btn)

        # Hidden by default (sidebar is visible on launch)
        self.quick_actions.setVisible(False)
        layout.addWidget(self.quick_actions)
        layout.addStretch()

        return topbar

    # ── Sidebar Fade Animation ───────────────────────────────
    def _setup_sidebar_animation(self):
        self.sidebar_effect = QGraphicsOpacityEffect(self.sidebar)
        self.sidebar.setGraphicsEffect(self.sidebar_effect)
        self.sidebar_effect.setOpacity(1.0)

        self.fade_anim = QPropertyAnimation(self.sidebar_effect, b"opacity")
        self.fade_anim.setDuration(250)
        self.fade_anim.setEasingCurve(QEasingCurve.InOutCubic)
        self.fade_anim.finished.connect(self._on_fade_finished)

    def toggle_sidebar(self):
        if self.sidebar_visible:
            # Fade OUT
            self.fade_anim.setStartValue(1.0)
            self.fade_anim.setEndValue(0.0)
            self.fade_anim.start()
        else:
            # Show sidebar then fade IN
            self.sidebar.setVisible(True)
            self.quick_actions.setVisible(False)
            self.fade_anim.setStartValue(0.0)
            self.fade_anim.setEndValue(1.0)
            self.fade_anim.start()

        self.sidebar_visible = not self.sidebar_visible

    def _on_fade_finished(self):
        if not self.sidebar_visible:
            # Sidebar faded out — collapse it and show quick actions
            self.sidebar.setVisible(False)
            self.quick_actions.setVisible(True)
            self.hamburger_btn.setText("☰")
        else:
            # Sidebar faded in
            self.hamburger_btn.setText("✕")

    # ── Sidebar ─────────────────────────────────────────────
    def _build_sidebar(self):
        content = QWidget()
        content.setObjectName("sidebar_content")
        content.setFixedWidth(230)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── App title with logo icon ────────────────────────
        title_widget = QWidget()
        title_layout = QVBoxLayout(title_widget)
        title_layout.setContentsMargins(16, 20, 16, 16)

        title_row = QWidget()
        title_row_layout = QHBoxLayout(title_row)
        title_row_layout.setContentsMargins(0, 0, 0, 0)
        title_row_layout.setSpacing(8)

        logo_label = QLabel()
        pixmap = QPixmap(self.logo_path)
        if not pixmap.isNull():
            pixmap = pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        else:
            logo_label.setText("🟦")
            logo_label.setStyleSheet("font-size: 24px;")

        title = QLabel("PyXcel")
        title.setObjectName("app_title")
        font = QFont()
        font.setPointSize(18)
        font.setBold(True)
        title.setFont(font)

        title_row_layout.addWidget(logo_label)
        title_row_layout.addWidget(title)
        title_row_layout.addStretch()

        subtitle = QLabel("AI Spreadsheet System")
        subtitle.setObjectName("app_subtitle")

        title_layout.addWidget(title_row)
        title_layout.addWidget(subtitle)
        layout.addWidget(title_widget)
        layout.addWidget(self._divider())

        # Ollama status
        self.ollama_indicator = QLabel("⬤  Checking Ollama...")
        self.ollama_indicator.setObjectName("ollama_indicator")
        self.ollama_indicator.setContentsMargins(16, 8, 16, 8)
        self.ollama_indicator.setProperty("status", "checking")
        layout.addWidget(self.ollama_indicator)
        layout.addWidget(self._divider())

        # ── Nav sections ──
        nav_items = [
            ("WORKSPACE", [
                ("  Home",               0),
                ("  Spreadsheet View",   1),
            ]),
            ("AI FEATURES", [
                ("  Formula Generator",  2),
                ("  Data Cleaner",        3),
                ("  Chat with Data",      4),
                ("  KPI Cards",           5),
            ]),
            ("TOOLS", [
                ("  Pivot Tables",        6),
                ("  Chart Creator",       7),
                ("  PDF Export",          8),
                ("  File Merger",         9),
            ]),
        ]

        self.nav_buttons = []
        for section_name, items in nav_items:
            sec = QLabel(section_name)
            sec.setObjectName("section_label")
            sec.setContentsMargins(16, 12, 16, 4)
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
        self.file_label.setObjectName("file_label")
        self.file_label.setContentsMargins(16, 10, 16, 4)
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

        load_btn = QPushButton("  Load Excel File")
        load_btn.setObjectName("btn_secondary")
        load_btn.setCursor(Qt.PointingHandCursor)
        load_btn.setProperty("compact", "true")
        load_btn.clicked.connect(self.load_file)
        layout.addWidget(load_btn)

        self.theme_btn = QPushButton("🌙  Dark Theme")
        self.theme_btn.setObjectName("theme_toggle")
        self.theme_btn.setProperty("compact", "true")
        self.theme_btn.setCursor(Qt.PointingHandCursor)
        self.theme_btn.clicked.connect(self.toggle_theme)
        layout.addWidget(self.theme_btn)

        spacer = QWidget()
        spacer.setFixedHeight(10)
        layout.addWidget(spacer)

        scroll = QScrollArea()
        scroll.setObjectName("sidebar")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setFixedWidth(230)
        scroll.setWidget(content)

        return scroll

    # ── Load Panels ─────────────────────────────────────────
    def _load_panels(self):
        from gui.panels.home_panel        import HomePanel
        from gui.panels.spreadsheet_panel import SpreadsheetPanel
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
            self.formula_panel,     # 2
            self.cleaner_panel,     # 3
            self.chat_panel,        # 4
            self.kpi_panel,         # 5
            self.pivot_panel,       # 6
            self.chart_panel,       # 7
            self.pdf_panel,         # 8
            self.merger_panel,      # 9
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
        self.file_label.setText(f" {filename}")
        self.status_bar.showMessage(f"Loaded: {filename}")
        self._notify_panels_file_loaded(path)
        self.switch_panel(1)

    def _notify_panels_file_loaded(self, path: str):
        panels = [
            self.spreadsheet_panel,
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
                    self._set_ollama_status("online")
                else:
                    self.ollama_indicator.setText("⬤  Model Not Pulled")
                    self._set_ollama_status("warning")
            else:
                self.ollama_indicator.setText("⬤  Ollama Offline")
                self._set_ollama_status("offline")

        QTimer.singleShot(500, check)
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(check)
        self.status_timer.start(30000)

    # ── Helpers ─────────────────────────────────────────────
    def _divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setObjectName("divider")
        return line

    def set_status(self, message: str):
        self.status_bar.showMessage(message)

    def _set_ollama_status(self, status: str):
        self.ollama_indicator.setProperty("status", status)
        self.ollama_indicator.style().unpolish(self.ollama_indicator)
        self.ollama_indicator.style().polish(self.ollama_indicator)

    def _load_saved_theme(self):
        saved = self.settings.value("ui/theme", "dark")
        self.apply_theme(saved if saved in {"dark", "light"} else "dark")

    def toggle_theme(self):
        self.apply_theme("light" if self.theme == "dark" else "dark")

    def apply_theme(self, theme_name: str):
        theme_files = {
            "dark": "styles.qss",
            "light": "styles_light.qss",
        }
        qss_name = theme_files.get(theme_name, "styles.qss")
        qss_path = Path(__file__).resolve().parent / qss_name

        if qss_path.exists():
            app = QApplication.instance()
            if app is not None:
                with qss_path.open("r", encoding="utf-8") as f:
                    app.setStyleSheet(f.read())

            self.theme = theme_name
            self.settings.setValue("ui/theme", theme_name)
            self.theme_btn.setText(
                "  Switch to Dark" if theme_name == "light" else "  Switch to Light"
            )