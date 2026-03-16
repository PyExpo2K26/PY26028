"""
PyXcel — Spreadsheet Panel
Displays Excel file contents in a table view with sheet tabs.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTableWidget, QTableWidgetItem,
    QTabBar, QFrame, QSizePolicy, QHeaderView,
    QAbstractItemView, QComboBox, QFileDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor


class SpreadsheetPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.sheet_names  = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(16)

        # ── Header ──
        header_row = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(4)

        badge = QLabel("SPREADSHEET VIEWER")
        badge.setStyleSheet("""
            background-color: #1e2035; color: #7c83ff;
            border-radius: 12px; padding: 3px 12px;
            font-size: 10px; font-weight: bold; letter-spacing: 1px;
        """)
        badge.setFixedHeight(24)

        title = QLabel("Spreadsheet View")
        font  = QFont()
        font.setPointSize(18)
        font.setBold(True)
        title.setFont(font)

        self.subtitle = QLabel("Load an Excel file to view its contents")
        self.subtitle.setStyleSheet("color: #555; font-size: 12px;")

        left.addWidget(badge)
        left.addWidget(title)
        left.addWidget(self.subtitle)

        # Right buttons
        right = QHBoxLayout()
        right.setSpacing(8)
        right.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.reload_btn = QPushButton("🔄  Reload")
        self.reload_btn.setFixedWidth(110)
        self.reload_btn.setCursor(Qt.PointingHandCursor)
        self.reload_btn.clicked.connect(self._reload)
        self.reload_btn.setEnabled(False)
        self.reload_btn.setStyleSheet("""
            QPushButton {
                background-color: #1e2035; color: #7c83ff;
                border: 1px solid #2a2d3e; border-radius: 8px;
                padding: 8px 14px; font-size: 12px;
            }
            QPushButton:hover { background-color: #252840; border-color: #7c83ff; }
            QPushButton:disabled { color: #333; border-color: #1e2035; }
        """)

        self.export_btn = QPushButton("💾  Save Copy")
        self.export_btn.setFixedWidth(120)
        self.export_btn.setCursor(Qt.PointingHandCursor)
        self.export_btn.clicked.connect(self._save_copy)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #7c83ff; color: white;
                border: none; border-radius: 8px;
                padding: 8px 14px; font-size: 12px; font-weight: bold;
            }
            QPushButton:hover { background-color: #6b72ff; }
            QPushButton:disabled { background-color: #2a2d3e; color: #555; }
        """)

        right.addWidget(self.reload_btn)
        right.addWidget(self.export_btn)

        header_row.addLayout(left)
        header_row.addStretch()
        header_row.addLayout(right)
        layout.addLayout(header_row)

        # ── Info Bar ──
        self.info_bar = self._build_info_bar()
        layout.addWidget(self.info_bar)

        # ── Sheet Tabs ──
        self.tab_bar = QTabBar()
        self.tab_bar.setStyleSheet("""
            QTabBar::tab {
                background: #1a1d2e; color: #555;
                border: 1px solid #2a2d3e; border-radius: 6px;
                padding: 6px 16px; margin-right: 4px; font-size: 12px;
            }
            QTabBar::tab:selected {
                background: #1e2035; color: #7c83ff; border-color: #7c83ff;
            }
            QTabBar::tab:hover { color: #c0c4ff; background: #1e2035; }
        """)
        self.tab_bar.currentChanged.connect(self._on_tab_changed)
        layout.addWidget(self.tab_bar)

        # ── Table ──
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #1a1d2e;
                alternate-background-color: #16192a;
                border: 1px solid #2a2d3e; border-radius: 8px;
                gridline-color: #2a2d3e; color: #e0e0e0; font-size: 12px;
            }
            QTableWidget::item { padding: 4px 10px; }
            QTableWidget::item:selected {
                background-color: #2a2d3e; color: #7c83ff;
            }
            QHeaderView::section {
                background-color: #13151f; color: #7c83ff;
                padding: 8px 10px; border: none;
                border-bottom: 1px solid #2a2d3e;
                border-right: 1px solid #2a2d3e;
                font-size: 11px; font-weight: bold;
            }
        """)
        layout.addWidget(self.table)

        # ── Empty State ──
        self.empty_state = self._build_empty_state()
        layout.addWidget(self.empty_state)
        self.table.hide()

    def _build_info_bar(self):
        bar = QFrame()
        bar.setStyleSheet("""
            QFrame {
                background-color: #1a1d2e; border: 1px solid #2a2d3e;
                border-radius: 8px; padding: 4px;
            }
        """)
        bar.setFixedHeight(42)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(14, 0, 14, 0)
        layout.setSpacing(24)

        self.rows_label  = self._info_chip("Rows",  "—")
        self.cols_label  = self._info_chip("Cols",  "—")
        self.sheet_label = self._info_chip("Sheet", "—")
        self.file_label  = self._info_chip("File",  "None")

        for chip in [self.rows_label, self.cols_label,
                     self.sheet_label, self.file_label]:
            layout.addWidget(chip)
        layout.addStretch()
        return bar

    def _info_chip(self, label, value):
        w = QLabel(
            f"<span style='color:#555'>{label}: </span>"
            f"<span style='color:#7c83ff'>{value}</span>"
        )
        w.setStyleSheet("background: transparent; font-size: 12px;")
        return w

    def _update_chip(self, chip, label, value):
        chip.setText(
            f"<span style='color:#555'>{label}: </span>"
            f"<span style='color:#7c83ff'>{value}</span>"
        )

    def _build_empty_state(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(12)

        icon = QLabel("📊")
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("font-size: 48px;")

        msg = QLabel("No file loaded")
        msg.setAlignment(Qt.AlignCenter)
        msg.setStyleSheet("color: #555; font-size: 15px;")

        hint = QLabel("Load an Excel file from the sidebar or Home panel")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color: #3a3d5e; font-size: 12px;")

        btn = QPushButton("📂  Browse Files")
        btn.setFixedWidth(160)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._browse_file)

        layout.addWidget(icon)
        layout.addWidget(msg)
        layout.addWidget(hint)
        layout.addWidget(btn, alignment=Qt.AlignCenter)
        return widget

    def on_file_loaded(self, path: str):
        self.current_file = path
        self._load_workbook()

    def _reload(self):
        if self.current_file:
            self._load_workbook()

    def _load_workbook(self):
        from gui.workers.agent_worker import InspectorWorker
        self.main_window.set_status("Loading workbook...")
        self.worker = InspectorWorker(self.current_file)
        self.worker.result.connect(self._on_workbook_loaded)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_workbook_loaded(self, data: dict):
        if data["status"] != "success":
            return

        self.workbook_data = data["data"]
        self.sheet_names   = list(data["data"].keys())

        filename = os.path.basename(self.current_file)
        self.subtitle.setText(
            f"{filename}  ·  {len(self.sheet_names)} sheet(s)"
        )

        self.tab_bar.blockSignals(True)
        while self.tab_bar.count():
            self.tab_bar.removeTab(0)
        for name in self.sheet_names:
            self.tab_bar.addTab(name)
        self.tab_bar.blockSignals(False)

        self._show_sheet(self.sheet_names[0])
        self.reload_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.main_window.set_status(
            f"Loaded: {filename}  |  {len(self.sheet_names)} sheet(s)"
        )

    def _on_tab_changed(self, index: int):
        if 0 <= index < len(self.sheet_names):
            self._show_sheet(self.sheet_names[index])

    def _show_sheet(self, sheet_name: str):
        try:
            import pandas as pd
            df = pd.read_excel(self.current_file, sheet_name=sheet_name)

            self.empty_state.hide()
            self.table.show()

            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(
                [str(c) for c in df.columns]
            )

            for row_idx, row in df.iterrows():
                for col_idx, value in enumerate(row):
                    display = "" if value != value else str(value)
                    item = QTableWidgetItem(display)
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                    self.table.setItem(row_idx, col_idx, item)

            self.table.resizeColumnsToContents()
            self._update_chip(self.rows_label,  "Rows",  str(len(df)))
            self._update_chip(self.cols_label,  "Cols",  str(len(df.columns)))
            self._update_chip(self.sheet_label, "Sheet", sheet_name)
            self._update_chip(
                self.file_label, "File",
                os.path.basename(self.current_file)[:30]
            )

        except Exception as e:
            self.main_window.set_status(f"Error loading sheet: {str(e)}")

    def _save_copy(self):
        if not self.current_file:
            return
        import shutil
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Copy",
            os.path.basename(self.current_file),
            "Excel Files (*.xlsx)"
        )
        if path:
            shutil.copy2(self.current_file, path)
            self.main_window.set_status(f"Saved copy to: {path}")

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.main_window.current_file = path
            self.main_window._notify_panels_file_loaded(path)

    def _on_error(self, msg: str):
        self.main_window.set_status(f"Error: {msg}")
        self.empty_state.show()
        self.table.hide()