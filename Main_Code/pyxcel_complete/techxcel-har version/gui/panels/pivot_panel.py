"""
PyXcel — Pivot Table Panel
Generate pivot tables manually or via AI natural language.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QPlainTextEdit,
    QFrame, QTabWidget, QScrollArea, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


AGG_FUNCTIONS = ["sum", "mean", "count", "max", "min", "median"]

PIVOT_EXAMPLES = [
    "Total Revenue by Region",
    "Average Salary by Department",
    "Count of Orders by Status",
    "Total Sales by Product and Region",
    "Maximum Profit by Category",
    "Average Performance Score by Department",
]


class PivotPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.sheet_names  = []
        self.columns      = []
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea{background:transparent;border:none;}")

        inner = QWidget()
        inner.setStyleSheet("background:transparent;")
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(16)

        # ── Header ──
        badge = QLabel("PIVOT TABLE GENERATOR")
        badge.setStyleSheet("background-color:#1e2035;color:#7c83ff;border-radius:12px;padding:3px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;")
        badge.setFixedHeight(24)

        title = QLabel("Pivot Table Generator")
        font  = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel("Create pivot tables manually by selecting columns, or let AI generate one from a plain-English description.")
        subtitle.setStyleSheet("color:#555;font-size:12px;")
        subtitle.setWordWrap(True)

        layout.addWidget(badge)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self._divider())

        # ── File Status ──
        self.file_status = QLabel("⚠️  No file loaded")
        self.file_status.setStyleSheet("color:#ff9800;font-size:12px;background:#1e1a0e;border-radius:6px;padding:8px 12px;")
        layout.addWidget(self.file_status)

        # ── Sheet Selector ──
        sheet_row = QHBoxLayout()
        sheet_label = QLabel("Source sheet:")
        sheet_label.setStyleSheet("color:#888;font-size:12px;")
        sheet_label.setFixedWidth(110)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setFixedWidth(200)
        self.sheet_combo.setStyleSheet(self._combo_style())
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        self.sheet_combo.addItem("Sheet1")
        sheet_row.addWidget(sheet_label)
        sheet_row.addWidget(self.sheet_combo)
        sheet_row.addStretch()
        layout.addLayout(sheet_row)

        # ── Tabs: Manual / AI ──
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                background: #1a1d2e;
            }
            QTabBar::tab {
                background: #13151f;
                color: #555;
                border: 1px solid #2a2d3e;
                border-radius: 6px;
                padding: 8px 20px;
                margin-right: 4px;
                font-size: 12px;
            }
            QTabBar::tab:selected {
                background: #1e2035;
                color: #7c83ff;
                border-color: #7c83ff;
            }
        """)

        self.tabs.addTab(self._build_manual_tab(), "⚙️  Manual")
        self.tabs.addTab(self._build_ai_tab(),     "🤖  AI Natural Language")
        layout.addWidget(self.tabs)

        # ── Result ──
        layout.addWidget(self._section_label("RESULT"))
        self.result_box = QPlainTextEdit()
        self.result_box.setReadOnly(True)
        self.result_box.setMinimumHeight(150)
        self.result_box.setPlaceholderText("Pivot table result will appear here...")
        self.result_box.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-left: 3px solid #7c83ff;
                border-radius: 8px;
                padding: 12px;
                color: #c0c4ff;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
            }
        """)
        layout.addWidget(self.result_box)
        layout.addStretch()

        scroll.setWidget(inner)
        outer.addWidget(scroll)

    # ── Manual Tab ───────────────────────────────────────────
    def _build_manual_tab(self):
        widget = QWidget()
        widget.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        # Row selector
        row1 = QHBoxLayout(); row1.setSpacing(12)
        lbl1 = QLabel("Row (Index):"); lbl1.setStyleSheet("color:#888;font-size:12px;"); lbl1.setFixedWidth(110)
        self.index_combo = QComboBox(); self.index_combo.setFixedWidth(180); self.index_combo.setStyleSheet(self._combo_style())
        row1.addWidget(lbl1); row1.addWidget(self.index_combo); row1.addStretch()

        # Value selector
        row2 = QHBoxLayout(); row2.setSpacing(12)
        lbl2 = QLabel("Values:"); lbl2.setStyleSheet("color:#888;font-size:12px;"); lbl2.setFixedWidth(110)
        self.value_combo = QComboBox(); self.value_combo.setFixedWidth(180); self.value_combo.setStyleSheet(self._combo_style())
        row2.addWidget(lbl2); row2.addWidget(self.value_combo); row2.addStretch()

        # Aggregation
        row3 = QHBoxLayout(); row3.setSpacing(12)
        lbl3 = QLabel("Aggregation:"); lbl3.setStyleSheet("color:#888;font-size:12px;"); lbl3.setFixedWidth(110)
        self.agg_combo = QComboBox(); self.agg_combo.setFixedWidth(180); self.agg_combo.setStyleSheet(self._combo_style())
        for func in AGG_FUNCTIONS:
            self.agg_combo.addItem(func)
        row3.addWidget(lbl3); row3.addWidget(self.agg_combo); row3.addStretch()

        # Columns (optional)
        row4 = QHBoxLayout(); row4.setSpacing(12)
        lbl4 = QLabel("Columns (opt):"); lbl4.setStyleSheet("color:#888;font-size:12px;"); lbl4.setFixedWidth(110)
        self.cols_combo = QComboBox(); self.cols_combo.setFixedWidth(180); self.cols_combo.setStyleSheet(self._combo_style())
        self.cols_combo.addItem("— None —")
        row4.addWidget(lbl4); row4.addWidget(self.cols_combo); row4.addStretch()

        # Run button
        self.manual_status = QLabel("Ready")
        self.manual_status.setStyleSheet("color:#555;font-size:11px;")
        self.manual_btn = QPushButton("📊  Generate Pivot Table")
        self.manual_btn.setFixedWidth(210)
        self.manual_btn.setCursor(Qt.PointingHandCursor)
        self.manual_btn.clicked.connect(self._run_manual_pivot)
        self.manual_btn.setEnabled(False)
        self.manual_btn.setStyleSheet(self._btn_style())

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.manual_status)
        btn_row.addStretch()
        btn_row.addWidget(self.manual_btn)

        layout.addLayout(row1)
        layout.addLayout(row2)
        layout.addLayout(row3)
        layout.addLayout(row4)
        layout.addLayout(btn_row)
        layout.addStretch()
        return widget

    # ── AI Tab ───────────────────────────────────────────────
    def _build_ai_tab(self):
        widget = QWidget()
        widget.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        lbl = QLabel("Describe the pivot table you need:")
        lbl.setStyleSheet("color:#888;font-size:12px;")

        self.ai_input = QPlainTextEdit()
        self.ai_input.setPlaceholderText(
            'e.g. "Show total Revenue by Region" or '
            '"Average salary grouped by Department"'
        )
        self.ai_input.setFixedHeight(100)
        self.ai_input.setStyleSheet("""
            QPlainTextEdit {
                background-color: #13151f;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 10px;
                color: #e0e0e0;
                font-size: 13px;
            }
            QPlainTextEdit:focus { border-color: #7c83ff; }
        """)

        # Examples
        ex_label = QLabel("💡  Quick examples:")
        ex_label.setStyleSheet("color:#444;font-size:11px;")

        examples_row = QHBoxLayout()
        examples_row.setSpacing(6)
        for ex in PIVOT_EXAMPLES[:3]:
            btn = QPushButton(ex)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda c, t=ex: self.ai_input.setPlainText(t))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #1a1d2e;
                    color: #555;
                    border: 1px solid #2a2d3e;
                    border-radius: 6px;
                    padding: 5px 10px;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #1e2035;
                    color: #c0c4ff;
                    border-color: #7c83ff;
                }
            """)
            examples_row.addWidget(btn)
        examples_row.addStretch()

        examples_row2 = QHBoxLayout()
        examples_row2.setSpacing(6)
        for ex in PIVOT_EXAMPLES[3:]:
            btn = QPushButton(ex)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda c, t=ex: self.ai_input.setPlainText(t))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #1a1d2e;
                    color: #555;
                    border: 1px solid #2a2d3e;
                    border-radius: 6px;
                    padding: 5px 10px;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #1e2035;
                    color: #c0c4ff;
                    border-color: #7c83ff;
                }
            """)
            examples_row2.addWidget(btn)
        examples_row2.addStretch()

        self.ai_status = QLabel("Ready")
        self.ai_status.setStyleSheet("color:#555;font-size:11px;")

        self.ai_btn = QPushButton("🤖  Generate with AI")
        self.ai_btn.setFixedWidth(190)
        self.ai_btn.setCursor(Qt.PointingHandCursor)
        self.ai_btn.clicked.connect(self._run_ai_pivot)
        self.ai_btn.setEnabled(False)
        self.ai_btn.setStyleSheet(self._btn_style("#4caf81", "#3d9e70"))

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.ai_status)
        btn_row.addStretch()
        btn_row.addWidget(self.ai_btn)

        layout.addWidget(lbl)
        layout.addWidget(self.ai_input)
        layout.addWidget(ex_label)
        layout.addLayout(examples_row)
        layout.addLayout(examples_row2)
        layout.addLayout(btn_row)
        layout.addStretch()
        return widget

    # ── Actions ──────────────────────────────────────────────
    def _run_manual_pivot(self):
        if not self.current_file:
            return
        index_col = self.index_combo.currentText()
        value_col = self.value_combo.currentText()
        agg_func  = self.agg_combo.currentText()
        cols_col  = self.cols_combo.currentText()
        if cols_col == "— None —":
            cols_col = None

        self.manual_btn.setEnabled(False)
        self.manual_btn.setText("⏳  Generating...")
        self.manual_status.setText("Building pivot table...")
        self.result_box.setPlainText("🔄  Processing...\n")

        from gui.workers.agent_worker import PivotWorker
        self.worker = PivotWorker(
            self.current_file, self.sheet_combo.currentText(),
            mode="manual",
            index_col=index_col, value_col=value_col,
            agg_func=agg_func, columns_col=cols_col
        )
        self.worker.status.connect(self._on_status)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _run_ai_pivot(self):
        if not self.current_file:
            return
        instruction = self.ai_input.toPlainText().strip()
        if not instruction:
            self.result_box.setPlainText("❌  Please enter a description.")
            return

        self.ai_btn.setEnabled(False)
        self.ai_btn.setText("⏳  Generating...")
        self.ai_status.setText("Asking LLaMA...")
        self.result_box.setPlainText("🔄  Processing...\n")

        from gui.workers.agent_worker import PivotWorker
        self.worker = PivotWorker(
            self.current_file, self.sheet_combo.currentText(),
            mode="ai", instruction=instruction
        )
        self.worker.status.connect(self._on_status)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_status(self, msg):
        self.manual_status.setText(msg)
        self.ai_status.setText(msg)
        self.main_window.set_status(msg)

    def _on_result(self, data):
        self.manual_btn.setEnabled(True)
        self.manual_btn.setText("📊  Generate Pivot Table")
        self.ai_btn.setEnabled(True)
        self.ai_btn.setText("🤖  Generate with AI")

        if data["status"] == "success":
            self.result_box.setPlainText(
                f"✅  {data.get('message', 'Done!')}\n\n"
                f"── Stats ───────────────────────────\n"
                f"  Rows    : {data.get('rows', '?')}\n"
                f"  Columns : {data.get('cols', '?')}\n"
                f"  Sheet   : {data.get('output_sheet', 'Pivot Table')}\n\n"
                f"── Preview ─────────────────────────\n"
                f"{data.get('preview', '')}"
            )
            self.manual_status.setText("Done ✓")
            self.ai_status.setText("Done ✓")
            self.main_window.set_status(
                f"Pivot table created ✓ — check '{data.get('output_sheet')}' sheet"
            )
            if hasattr(self.main_window, "spreadsheet_panel"):
                self.main_window.spreadsheet_panel._reload()
        else:
            self.result_box.setPlainText(
                f"❌  Failed: {data.get('message', 'Unknown error')}"
            )
            self.main_window.set_status("Pivot generation failed ✗")

    def _on_error(self, msg):
        self.manual_btn.setEnabled(True)
        self.manual_btn.setText("📊  Generate Pivot Table")
        self.ai_btn.setEnabled(True)
        self.ai_btn.setText("🤖  Generate with AI")
        self.result_box.setPlainText(f"❌  Error: {msg}")
        self.main_window.set_status(f"Error: {msg}")

    def _on_sheet_changed(self, sheet):
        if not self.current_file or not sheet:
            return
        self._load_columns(sheet)

    def _load_columns(self, sheet):
        try:
            import pandas as pd
            df = pd.read_excel(self.current_file, sheet_name=sheet)
            self.columns = list(df.columns)
            numeric = [
                c for c in self.columns
                if pd.api.types.is_numeric_dtype(df[c])
            ]

            for combo in [self.index_combo, self.value_combo, self.cols_combo]:
                combo.clear()

            if combo == self.cols_combo:
                self.cols_combo.addItem("— None —")

            self.cols_combo.addItem("— None —")
            for c in self.columns:
                self.index_combo.addItem(str(c))
                self.cols_combo.addItem(str(c))
            for c in numeric:
                self.value_combo.addItem(str(c))

            self.manual_btn.setEnabled(True)
            self.ai_btn.setEnabled(True)
        except Exception:
            pass

    def on_file_loaded(self, path):
        self.current_file = path
        filename = os.path.basename(path)
        self.file_status.setText(f"✅  File loaded: {filename}")
        self.file_status.setStyleSheet(
            "color:#4caf81;font-size:12px;background:#0e1e14;border-radius:6px;padding:8px 12px;"
        )
        from core.workbook_inspector import get_sheet_names
        sheets = get_sheet_names(path)
        self.sheet_combo.clear()
        for s in sheets:
            self.sheet_combo.addItem(s)
        self.sheet_names = sheets
        if sheets:
            self._load_columns(sheets[0])

    # ── Helpers ──────────────────────────────────────────────
    def _combo_style(self):
        return (
            "QComboBox{background-color:#13151f;border:1px solid #2a2d3e;"
            "border-radius:8px;padding:7px 12px;color:#e0e0e0;font-size:12px;}"
            "QComboBox:focus{border-color:#7c83ff;}"
            "QComboBox QAbstractItemView{background-color:#1a1d2e;"
            "border:1px solid #2a2d3e;color:#e0e0e0;}"
        )

    def _btn_style(self, color="#7c83ff", hover="#6b72ff"):
        return (
            f"QPushButton{{background-color:{color};color:white;border:none;"
            f"border-radius:8px;padding:9px 20px;font-size:13px;font-weight:bold;}}"
            f"QPushButton:hover{{background-color:{hover};}}"
            f"QPushButton:disabled{{background-color:#2a2d3e;color:#555;}}"
        )

    def _divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#2a2d3e;max-height:1px;border:none;")
        return line

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#444;font-size:10px;letter-spacing:1px;")
        return lbl
