"""
PyXcel — PDF Export Panel
Export Excel sheets to styled PDF reports.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QFrame,
    QScrollArea, QCheckBox, QFileDialog, QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class PdfPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.sheet_names  = []
        self.last_output  = None
        self._build_ui()

    # ── Build UI ────────────────────────────────────────────
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
        badge = QLabel("PDF EXPORT")
        badge.setStyleSheet("background-color:#1e2035;color:#7c83ff;border-radius:12px;padding:3px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;")
        badge.setFixedHeight(24)

        title = QLabel("PDF Export")
        font  = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel(
            "Export your Excel sheets to professional styled PDF reports "
            "with summary statistics, data tables, and PyXcel branding."
        )
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

        # ── Main Content ──
        content_row = QHBoxLayout()
        content_row.setSpacing(20)
        content_row.addWidget(self._build_settings_panel(), stretch=2)
        content_row.addWidget(self._build_info_panel(),     stretch=3)
        layout.addLayout(content_row)

        # ── Result Box ──
        layout.addWidget(self._section_label("EXPORT RESULT"))
        self.result_box = QLabel("Export result will appear here...")
        self.result_box.setWordWrap(True)
        self.result_box.setStyleSheet("""
            QLabel {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-left: 3px solid #7c83ff;
                border-radius: 8px;
                padding: 16px;
                color: #c0c4ff;
                font-size: 12px;
                min-height: 80px;
            }
        """)
        layout.addWidget(self.result_box)

        # Open PDF button
        self.open_btn = QPushButton("📂  Open PDF Location")
        self.open_btn.setFixedWidth(200)
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.clicked.connect(self._open_output_folder)
        self.open_btn.setEnabled(False)
        self.open_btn.setStyleSheet("""
            QPushButton {
                background-color: #1e2035;
                color: #7c83ff;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 9px 16px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #252840;
                border-color: #7c83ff;
            }
            QPushButton:disabled { color: #333; border-color: #1e2035; }
        """)
        layout.addWidget(self.open_btn)
        layout.addStretch()

        scroll.setWidget(inner)
        outer.addWidget(scroll)

    # ── Settings Panel ───────────────────────────────────────
    def _build_settings_panel(self):
        frame = QFrame()
        frame.setStyleSheet("QFrame{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:12px;}")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        section = QLabel("Export Settings")
        section.setStyleSheet("color:#7c83ff;font-size:13px;font-weight:bold;background:transparent;")
        layout.addWidget(section)

        # Export mode
        mode_label = QLabel("Export mode:")
        mode_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")
        self.mode_combo = QComboBox()
        self.mode_combo.setStyleSheet(self._combo_style())
        self.mode_combo.addItem("📄  Single Sheet")
        self.mode_combo.addItem("📚  All Sheets")
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)

        # Sheet selector
        self.sheet_label = QLabel("Select sheet:")
        self.sheet_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")
        self.sheet_combo = QComboBox()
        self.sheet_combo.setStyleSheet(self._combo_style())
        self.sheet_combo.addItem("Sheet1")

        # Report title
        title_label = QLabel("Report title (optional):")
        title_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("Auto-generated if left blank")
        self.title_input.setStyleSheet("""
            QLineEdit {
                background-color: #13151f;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 8px 12px;
                color: #e0e0e0;
                font-size: 12px;
            }
            QLineEdit:focus { border-color: #7c83ff; }
        """)

        # Options
        opts_label = QLabel("Options:")
        opts_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")

        self.summary_check = QCheckBox("Include summary statistics")
        self.summary_check.setChecked(True)
        self.summary_check.setStyleSheet("""
            QCheckBox {
                color: #888;
                font-size: 12px;
                background: transparent;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #2a2d3e;
                border-radius: 4px;
                background: #13151f;
            }
            QCheckBox::indicator:checked {
                background: #7c83ff;
                border-color: #7c83ff;
            }
        """)

        # Output path
        path_label = QLabel("Save location:")
        path_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")

        path_row = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("Click Browse to choose location...")
        self.path_input.setReadOnly(True)
        self.path_input.setStyleSheet("""
            QLineEdit {
                background-color: #13151f;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 8px 12px;
                color: #666;
                font-size: 11px;
            }
        """)
        browse_btn = QPushButton("Browse")
        browse_btn.setFixedWidth(70)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse_output)
        browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #1e2035;
                color: #7c83ff;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
                padding: 8px;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #252840; }
        """)
        path_row.addWidget(self.path_input)
        path_row.addWidget(browse_btn)

        # Status + Export button
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color:#555;font-size:11px;background:transparent;")

        self.export_btn = QPushButton("📄  Export to PDF")
        self.export_btn.setFixedHeight(42)
        self.export_btn.setCursor(Qt.PointingHandCursor)
        self.export_btn.clicked.connect(self._export_pdf)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #d32f2f; }
            QPushButton:disabled { background-color: #2a2d3e; color: #555; }
        """)

        layout.addWidget(mode_label)
        layout.addWidget(self.mode_combo)
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.sheet_combo)
        layout.addWidget(title_label)
        layout.addWidget(self.title_input)
        layout.addWidget(opts_label)
        layout.addWidget(self.summary_check)
        layout.addWidget(path_label)
        layout.addLayout(path_row)
        layout.addWidget(self._mini_divider())
        layout.addWidget(self.status_label)
        layout.addWidget(self.export_btn)
        layout.addStretch()

        return frame

    # ── Info Panel ───────────────────────────────────────────
    def _build_info_panel(self):
        frame = QFrame()
        frame.setStyleSheet("QFrame{background-color:#13151f;border:1px solid #2a2d3e;border-radius:12px;}")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        title = QLabel("What's included in the PDF")
        title.setStyleSheet("color:#7c83ff;font-size:13px;font-weight:bold;background:transparent;")
        layout.addWidget(title)

        features = [
            ("📋", "Report Header",
             "Title, generation date, sheet name, row and column counts"),
            ("📊", "Summary Statistics",
             "Count, sum, average, min and max for all numeric columns"),
            ("📄", "Full Data Table",
             "All rows formatted in a styled table (up to 500 rows per sheet)"),
            ("🎨", "PyXcel Dark Theme",
             "Professional dark-themed PDF matching the PyXcel UI style"),
            ("📐", "Landscape Layout",
             "A4 landscape format for maximum data visibility"),
            ("🔖", "Page Headers",
             "Sheet name and page numbers on every page"),
        ]

        for icon, feat_title, desc in features:
            row = QHBoxLayout()
            row.setSpacing(12)

            icon_lbl = QLabel(icon)
            icon_lbl.setFixedWidth(28)
            icon_lbl.setStyleSheet("font-size:18px;background:transparent;")

            text_col = QVBoxLayout()
            text_col.setSpacing(2)
            feat_lbl = QLabel(feat_title)
            feat_lbl.setStyleSheet("color:#c0c4ff;font-size:12px;font-weight:bold;background:transparent;")
            desc_lbl = QLabel(desc)
            desc_lbl.setStyleSheet("color:#444;font-size:11px;background:transparent;")
            desc_lbl.setWordWrap(True)
            text_col.addWidget(feat_lbl)
            text_col.addWidget(desc_lbl)

            row.addWidget(icon_lbl)
            row.addLayout(text_col)
            layout.addLayout(row)

        layout.addStretch()

        # Preview card
        preview_card = QFrame()
        preview_card.setStyleSheet("""
            QFrame {
                background-color: #1a1d2e;
                border: 1px solid #2a2d3e;
                border-radius: 8px;
            }
        """)
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(16, 12, 16, 12)
        preview_layout.setSpacing(4)

        tip_title = QLabel("💡  Tips")
        tip_title.setStyleSheet("color:#ff9800;font-size:12px;font-weight:bold;background:transparent;")
        tip1 = QLabel("• Use 'All Sheets' to export the entire workbook as one PDF")
        tip1.setStyleSheet("color:#555;font-size:11px;background:transparent;")
        tip2 = QLabel("• Report title defaults to the filename if left blank")
        tip2.setStyleSheet("color:#555;font-size:11px;background:transparent;")
        tip3 = QLabel("• PDF is saved to outputs/ folder by default")
        tip3.setStyleSheet("color:#555;font-size:11px;background:transparent;")

        preview_layout.addWidget(tip_title)
        preview_layout.addWidget(tip1)
        preview_layout.addWidget(tip2)
        preview_layout.addWidget(tip3)

        layout.addWidget(preview_card)
        return frame

    # ── Actions ──────────────────────────────────────────────
    def _export_pdf(self):
        if not self.current_file:
            return

        output_path = self.path_input.text().strip()
        if not output_path:
            # Default to outputs folder
            base = os.path.splitext(
                os.path.basename(self.current_file)
            )[0]
            output_dir = os.path.join(
                os.path.dirname(self.current_file), "outputs"
            )
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, f"{base}_report.pdf")
            self.path_input.setText(output_path)

        export_all = self.mode_combo.currentIndex() == 1
        sheet      = self.sheet_combo.currentText()
        title      = self.title_input.text().strip()
        summary    = self.summary_check.isChecked()

        self.export_btn.setEnabled(False)
        self.export_btn.setText("⏳  Exporting...")
        self.status_label.setText("Generating PDF...")
        self.result_box.setText("🔄  Creating PDF report...")

        from gui.workers.agent_worker import PdfWorker
        self.worker = PdfWorker(
            self.current_file, sheet,
            output_path, title,
            export_all, summary
        )
        self.worker.status.connect(self._on_status)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_status(self, msg):
        self.status_label.setText(msg)
        self.main_window.set_status(msg)

    def _on_result(self, data):
        self.export_btn.setEnabled(True)
        self.export_btn.setText("📄  Export to PDF")

        if data["status"] == "success":
            self.last_output = data.get("output_path", "")
            rows    = data.get("rows",    "?")
            cols    = data.get("cols",    "?")
            sheets  = data.get("sheets",  1)
            trunc   = data.get("truncated", False)

            detail = (
                f"  Sheets exported : {sheets}\n"
                if sheets > 1 else
                f"  Rows  : {rows}\n"
                f"  Cols  : {cols}\n"
            )

            self.result_box.setText(
                f"✅  PDF exported successfully!\n\n"
                f"── Details ─────────────────────────\n"
                f"{detail}"
                f"  Saved : {os.path.basename(self.last_output)}\n"
                f"  Path  : {self.last_output}\n"
                + (f"\n⚠️  Large file: showing first 500 rows only" if trunc else "")
            )
            self.status_label.setText("Export complete ✓")
            self.open_btn.setEnabled(True)
            self.main_window.set_status(
                f"PDF exported: {os.path.basename(self.last_output)} ✓"
            )
        else:
            self.result_box.setText(
                f"❌  Export failed.\n\n"
                f"Error: {data.get('message', 'Unknown error')}\n\n"
                f"Make sure reportlab is installed:\n"
                f"  pip install reportlab"
            )
            self.status_label.setText("Export failed ✗")
            self.main_window.set_status("PDF export failed ✗")

    def _on_error(self, msg):
        self.export_btn.setEnabled(True)
        self.export_btn.setText("📄  Export to PDF")
        self.status_label.setText("Error ✗")
        self.result_box.setText(
            f"❌  Error: {msg}\n\n"
            f"Make sure reportlab is installed:\n"
            f"  pip install reportlab"
        )
        self.main_window.set_status(f"Error: {msg}")

    def _browse_output(self):
        base = os.path.splitext(
            os.path.basename(self.current_file or "report")
        )[0]
        path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF As",
            f"{base}_report.pdf",
            "PDF Files (*.pdf)"
        )
        if path:
            self.path_input.setText(path)

    def _open_output_folder(self):
        if self.last_output and os.path.exists(self.last_output):
            import subprocess
            folder = os.path.dirname(self.last_output)
            subprocess.Popen(f'explorer "{folder}"')

    def _on_mode_changed(self, index):
        # Hide sheet selector when exporting all
        self.sheet_label.setVisible(index == 0)
        self.sheet_combo.setVisible(index == 0)
        self.summary_check.setVisible(index == 0)

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
        self.export_btn.setEnabled(True)
        self.status_label.setText("Ready — configure settings and export")

        # Set default output path
        base       = os.path.splitext(filename)[0]
        output_dir = os.path.join(os.path.dirname(path), "outputs")
        os.makedirs(output_dir, exist_ok=True)
        self.path_input.setText(
            os.path.join(output_dir, f"{base}_report.pdf")
        )

    # ── Helpers ──────────────────────────────────────────────
    def _combo_style(self):
        return (
            "QComboBox{background-color:#13151f;border:1px solid #2a2d3e;"
            "border-radius:8px;padding:7px 12px;color:#e0e0e0;font-size:12px;}"
            "QComboBox:focus{border-color:#7c83ff;}"
            "QComboBox QAbstractItemView{background-color:#1a1d2e;"
            "border:1px solid #2a2d3e;color:#e0e0e0;}"
        )

    def _divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#2a2d3e;max-height:1px;border:none;")
        return line

    def _mini_divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#1e2035;max-height:1px;border:none;")
        return line

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#444;font-size:10px;letter-spacing:1px;")
        return lbl
