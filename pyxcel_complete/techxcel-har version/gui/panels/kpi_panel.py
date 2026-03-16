"""
PyXcel — KPI Cards Panel
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QScrollArea, QComboBox,
    QGridLayout, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

TREND_ICON  = {"up": "↑", "down": "↓", "neutral": "→"}
TREND_COLOR = {"up": "#4caf81", "down": "#f44336", "neutral": "#ff9800"}


class KpiCard(QFrame):
    def __init__(self, title, value, description, trend="neutral", parent=None):
        super().__init__(parent)
        self.setMinimumWidth(200)
        self.setMinimumHeight(130)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        color = TREND_COLOR.get(trend, "#ff9800")
        icon  = TREND_ICON.get(trend,  "→")
        self.setStyleSheet(f"QFrame{{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:12px;border-top:3px solid {color};}}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(6)

        title_row = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setStyleSheet("color:#888;font-size:11px;font-weight:bold;background:transparent;")
        title_label.setWordWrap(True)
        trend_label = QLabel(icon)
        trend_label.setStyleSheet(f"color:{color};font-size:16px;font-weight:bold;background:transparent;")
        trend_label.setFixedWidth(24)
        trend_label.setAlignment(Qt.AlignRight)
        title_row.addWidget(title_label)
        title_row.addStretch()
        title_row.addWidget(trend_label)

        value_label = QLabel(str(value))
        value_label.setStyleSheet(f"color:{color};font-size:26px;font-weight:bold;background:transparent;")
        value_label.setWordWrap(True)

        desc_label = QLabel(description)
        desc_label.setStyleSheet("color:#444;font-size:11px;background:transparent;")
        desc_label.setWordWrap(True)

        layout.addLayout(title_row)
        layout.addWidget(value_label)
        layout.addStretch()
        layout.addWidget(desc_label)


class KpiPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.sheet_names  = []
        self.kpi_cards    = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        header_row = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(4)

        badge = QLabel("AUTO KPI DETECTION")
        badge.setStyleSheet("background-color:#1e2035;color:#7c83ff;border-radius:12px;padding:3px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;")
        badge.setFixedHeight(24)

        title = QLabel("KPI Cards")
        font  = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel("LLaMA automatically identifies business KPIs from any spreadsheet and computes their values.")
        subtitle.setStyleSheet("color:#555;font-size:12px;")
        subtitle.setWordWrap(True)

        left.addWidget(badge)
        left.addWidget(title)
        left.addWidget(subtitle)

        self.refresh_btn = QPushButton("🔄  Regenerate")
        self.refresh_btn.setFixedWidth(140)
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.clicked.connect(self._generate_kpis)
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setStyleSheet("QPushButton{background-color:#1e2035;color:#7c83ff;border:1px solid #2a2d3e;border-radius:8px;padding:9px 16px;font-size:12px;}QPushButton:hover{background-color:#252840;border-color:#7c83ff;}QPushButton:disabled{color:#333;border-color:#1e2035;}")

        header_row.addLayout(left)
        header_row.addStretch()
        header_row.addWidget(self.refresh_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)
        layout.addLayout(header_row)
        layout.addWidget(self._divider())

        self.file_status = QLabel("⚠️  No file loaded — load an Excel file first")
        self.file_status.setStyleSheet("color:#ff9800;font-size:12px;background:#1e1a0e;border-radius:6px;padding:8px 12px;")
        layout.addWidget(self.file_status)

        controls_row = QHBoxLayout()
        controls_row.setSpacing(12)

        sheet_label = QLabel("Sheet:")
        sheet_label.setStyleSheet("color:#888;font-size:12px;")
        sheet_label.setFixedWidth(50)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setFixedWidth(180)
        self.sheet_combo.setStyleSheet("QComboBox{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:8px;padding:7px 12px;color:#e0e0e0;font-size:12px;}QComboBox:focus{border-color:#7c83ff;}")
        self.sheet_combo.addItem("Sheet1")

        self.generate_btn = QPushButton("📈  Generate KPI Cards")
        self.generate_btn.setFixedWidth(190)
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.clicked.connect(self._generate_kpis)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setStyleSheet("QPushButton{background-color:#7c83ff;color:white;border:none;border-radius:8px;padding:9px 20px;font-size:13px;font-weight:bold;}QPushButton:hover{background-color:#6b72ff;}QPushButton:disabled{background-color:#2a2d3e;color:#555;}")

        self.status_label = QLabel("Load a file to get started")
        self.status_label.setStyleSheet("color:#555;font-size:12px;")

        controls_row.addWidget(sheet_label)
        controls_row.addWidget(self.sheet_combo)
        controls_row.addWidget(self.generate_btn)
        controls_row.addStretch()
        controls_row.addWidget(self.status_label)
        layout.addLayout(controls_row)

        self.empty_state = self._build_empty_state()
        layout.addWidget(self.empty_state)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setStyleSheet("background:transparent;")
        self.scroll_area.hide()

        self.grid_container = QWidget()
        self.grid_container.setStyleSheet("background:transparent;")
        self.grid_layout = QGridLayout(self.grid_container)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)
        self.grid_layout.setSpacing(16)

        self.scroll_area.setWidget(self.grid_container)
        layout.addWidget(self.scroll_area, stretch=1)

        self.summary_frame = self._build_summary_box()
        self.summary_frame.hide()
        layout.addWidget(self.summary_frame)

    def _build_empty_state(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(12)
        icon = QLabel("📈"); icon.setAlignment(Qt.AlignCenter); icon.setStyleSheet("font-size:52px;")
        msg  = QLabel("No KPIs generated yet"); msg.setAlignment(Qt.AlignCenter); msg.setStyleSheet("color:#555;font-size:15px;")
        hint = QLabel("Load an Excel file and click 'Generate KPI Cards'\nLLaMA will automatically identify relevant business metrics")
        hint.setAlignment(Qt.AlignCenter); hint.setStyleSheet("color:#3a3d5e;font-size:12px;"); hint.setWordWrap(True)
        layout.addWidget(icon); layout.addWidget(msg); layout.addWidget(hint)
        return widget

    def _build_summary_box(self):
        frame = QFrame()
        frame.setStyleSheet("QFrame{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:10px;}")
        frame.setFixedHeight(48)
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(16, 0, 16, 0)
        layout.setSpacing(24)
        self.total_kpi_label = self._info_chip("KPIs Found",     "—")
        self.up_kpi_label    = self._info_chip("Trending Up",    "—")
        self.down_kpi_label  = self._info_chip("Trending Down",  "—")
        self.sheet_kpi_label = self._info_chip("Sheet",          "—")
        for chip in [self.total_kpi_label, self.up_kpi_label, self.down_kpi_label, self.sheet_kpi_label]:
            layout.addWidget(chip)
        layout.addStretch()
        return frame

    def _info_chip(self, label, value):
        w = QLabel(f"<span style='color:#444'>{label}: </span><span style='color:#7c83ff;font-weight:bold'>{value}</span>")
        w.setStyleSheet("background:transparent;font-size:12px;")
        return w

    def _update_chip(self, chip, label, value, color="#7c83ff"):
        chip.setText(f"<span style='color:#444'>{label}: </span><span style='color:{color};font-weight:bold'>{value}</span>")

    def _generate_kpis(self):
        if not self.current_file:
            return
        sheet = self.sheet_combo.currentText().strip() or "Sheet1"
        self.generate_btn.setEnabled(False)
        self.refresh_btn.setEnabled(False)
        self.generate_btn.setText("⏳  Analyzing...")
        self.status_label.setText("LLaMA is analyzing your data...")
        self._clear_cards()
        self.empty_state.hide()
        self.scroll_area.hide()
        self.summary_frame.hide()
        self.main_window.set_status("Generating KPIs with LLaMA...")

        from gui.workers.agent_worker import KpiWorker
        self.worker = KpiWorker(self.current_file, sheet)
        self.worker.status.connect(self._on_status)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_status(self, msg):
        self.status_label.setText(msg)
        self.main_window.set_status(msg)

    def _on_result(self, data):
        self.generate_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.generate_btn.setText("📈  Generate KPI Cards")
        if data["status"] == "success":
            kpis = data.get("kpis", [])
            self._clear_cards()
            self._render_cards(kpis)
            self.status_label.setText(f"{len(kpis)} KPIs generated ✓")
            self.main_window.set_status(f"{len(kpis)} KPI cards generated ✓")
        else:
            self._clear_cards()
            self.empty_state.show()
            self.scroll_area.hide()
            self.status_label.setText("Generation failed ✗")

    def _on_error(self, msg):
        self.generate_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.generate_btn.setText("📈  Generate KPI Cards")
        self._clear_cards()
        self.empty_state.show()
        self.scroll_area.hide()
        self.status_label.setText("Error — is Ollama running?")
        self.main_window.set_status(f"Error: {msg}")

    def _render_cards(self, kpis):
        if not kpis:
            self.empty_state.show()
            self.scroll_area.hide()
            return
        self.scroll_area.show()
        self.summary_frame.show()
        up_count   = sum(1 for k in kpis if k.get("trend") == "up")
        down_count = sum(1 for k in kpis if k.get("trend") == "down")
        sheet      = self.sheet_combo.currentText()
        self._update_chip(self.total_kpi_label, "KPIs Found",    str(len(kpis)))
        self._update_chip(self.up_kpi_label,    "Trending Up",   str(up_count),   "#4caf81")
        self._update_chip(self.down_kpi_label,  "Trending Down", str(down_count), "#f44336")
        self._update_chip(self.sheet_kpi_label, "Sheet",         sheet)
        cols = 3
        for i, kpi in enumerate(kpis):
            card = KpiCard(
                title       = kpi.get("title",       "KPI"),
                value       = kpi.get("value",       "—"),
                description = kpi.get("description", ""),
                trend       = kpi.get("trend",       "neutral")
            )
            self.grid_layout.addWidget(card, i // cols, i % cols)
            self.kpi_cards.append(card)

    def _clear_cards(self):
        for card in self.kpi_cards:
            self.grid_layout.removeWidget(card)
            card.deleteLater()
        self.kpi_cards.clear()
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def on_file_loaded(self, path):
        self.current_file = path
        filename = os.path.basename(path)
        self.file_status.setText(f"✅  File loaded: {filename}")
        self.file_status.setStyleSheet("color:#4caf81;font-size:12px;background:#0e1e14;border-radius:6px;padding:8px 12px;")
        from core.workbook_inspector import get_sheet_names
        sheets = get_sheet_names(path)
        self.sheet_combo.clear()
        for s in sheets:
            self.sheet_combo.addItem(s)
        self.sheet_names = sheets
        self.generate_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.status_label.setText("Ready — click Generate KPI Cards")
        self._clear_cards()
        self.empty_state.show()
        self.scroll_area.hide()
        self.summary_frame.hide()

    def _divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#2a2d3e;max-height:1px;border:none;")
        return line
