"""
PyXcel — Multi-file Merger Panel
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QFrame, QScrollArea,
    QListWidget, QListWidgetItem, QFileDialog,
    QAbstractItemView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class MergerPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.file_list    = []
        self.last_output  = None
        self._output_path = None
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

        badge = QLabel("MULTI-FILE MERGER")
        badge.setStyleSheet("background-color:#1e2035;color:#7c83ff;border-radius:12px;padding:3px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;")
        badge.setFixedHeight(24)

        title = QLabel("File Merger")
        font  = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel("Merge multiple Excel files into a single workbook. Choose to merge as separate sheets or stack all rows together.")
        subtitle.setStyleSheet("color:#555;font-size:12px;")
        subtitle.setWordWrap(True)

        layout.addWidget(badge)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self._divider())

        content_row = QHBoxLayout()
        content_row.setSpacing(20)
        content_row.addWidget(self._build_files_panel(),    stretch=3)
        content_row.addWidget(self._build_settings_panel(), stretch=2)
        layout.addLayout(content_row)

        layout.addWidget(self._section_label("RESULT"))
        self.result_box = QLabel("Merge result will appear here...")
        self.result_box.setWordWrap(True)
        self.result_box.setMinimumHeight(80)
        self.result_box.setStyleSheet("QLabel{background-color:#1a1d2e;border:1px solid #2a2d3e;border-left:3px solid #7c83ff;border-radius:8px;padding:16px;color:#c0c4ff;font-size:12px;}")
        layout.addWidget(self.result_box)
        layout.addStretch()

        scroll.setWidget(inner)
        outer.addWidget(scroll)

    def _build_files_panel(self):
        frame = QFrame()
        frame.setStyleSheet("QFrame{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:12px;}")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        header = QHBoxLayout()
        title  = QLabel("Files to Merge")
        title.setStyleSheet("color:#7c83ff;font-size:13px;font-weight:bold;background:transparent;")
        self.count_label = QLabel("0 files")
        self.count_label.setStyleSheet("color:#555;font-size:11px;background:transparent;")
        header.addWidget(title); header.addStretch(); header.addWidget(self.count_label)
        layout.addLayout(header)

        hint = QLabel("Add .xlsx or .xls files to merge")
        hint.setStyleSheet("color:#3a3d5e;font-size:11px;background:transparent;")
        layout.addWidget(hint)

        self.file_listbox = QListWidget()
        self.file_listbox.setMinimumHeight(220)
        self.file_listbox.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_listbox.setStyleSheet("""
            QListWidget{background-color:#13151f;border:1px solid #2a2d3e;border-radius:8px;color:#e0e0e0;font-size:12px;padding:4px;}
            QListWidget::item{padding:8px 12px;border-radius:6px;margin:2px;}
            QListWidget::item:selected{background-color:#1e2035;color:#7c83ff;}
        """)
        layout.addWidget(self.file_listbox)

        btn_row = QHBoxLayout(); btn_row.setSpacing(8)

        add_btn = QPushButton("➕  Add Files")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.clicked.connect(self._add_files)
        add_btn.setStyleSheet("QPushButton{background-color:#7c83ff;color:white;border:none;border-radius:8px;padding:8px 16px;font-size:12px;font-weight:bold;}QPushButton:hover{background-color:#6b72ff;}")

        remove_btn = QPushButton("➖  Remove")
        remove_btn.setCursor(Qt.PointingHandCursor)
        remove_btn.clicked.connect(self._remove_selected)
        remove_btn.setStyleSheet("QPushButton{background-color:#1e2035;color:#f44336;border:1px solid #3d1a1a;border-radius:8px;padding:8px 16px;font-size:12px;}QPushButton:hover{background-color:#3d1a1a;}")

        clear_btn = QPushButton("🗑  Clear All")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self._clear_files)
        clear_btn.setStyleSheet("QPushButton{background-color:#1e2035;color:#888;border:1px solid #2a2d3e;border-radius:8px;padding:8px 16px;font-size:12px;}QPushButton:hover{background-color:#252840;}")

        btn_row.addWidget(add_btn); btn_row.addWidget(remove_btn); btn_row.addWidget(clear_btn); btn_row.addStretch()
        layout.addLayout(btn_row)
        return frame

    def _build_settings_panel(self):
        frame = QFrame()
        frame.setStyleSheet("QFrame{background-color:#13151f;border:1px solid #2a2d3e;border-radius:12px;}")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        title = QLabel("Merge Settings")
        title.setStyleSheet("color:#7c83ff;font-size:13px;font-weight:bold;background:transparent;")
        layout.addWidget(title)

        mode_label = QLabel("Merge mode:")
        mode_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")

        self.mode_combo = QComboBox()
        self.mode_combo.setStyleSheet("QComboBox{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:8px;padding:7px 12px;color:#e0e0e0;font-size:12px;}QComboBox QAbstractItemView{background-color:#1a1d2e;border:1px solid #2a2d3e;color:#e0e0e0;}")
        self.mode_combo.addItem("📑  Separate Sheets")
        self.mode_combo.addItem("📋  Stack Rows")
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)

        self.mode_desc = QLabel("Each file becomes a separate sheet\nin the merged workbook.")
        self.mode_desc.setStyleSheet("color:#3a3d5e;font-size:11px;background:transparent;")
        self.mode_desc.setWordWrap(True)

        layout.addWidget(mode_label)
        layout.addWidget(self.mode_combo)
        layout.addWidget(self.mode_desc)
        layout.addWidget(self._mini_divider())

        path_label = QLabel("Output file:")
        path_label.setStyleSheet("color:#888;font-size:11px;background:transparent;")
        path_row = QHBoxLayout()
        self.path_display = QLabel("Auto — saved to outputs/")
        self.path_display.setStyleSheet("color:#444;font-size:11px;background:transparent;")
        browse_btn = QPushButton("Browse")
        browse_btn.setFixedWidth(70)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse_output)
        browse_btn.setStyleSheet("QPushButton{background-color:#1e2035;color:#7c83ff;border:1px solid #2a2d3e;border-radius:8px;padding:6px;font-size:11px;}QPushButton:hover{background-color:#252840;}")
        path_row.addWidget(self.path_display, stretch=1)
        path_row.addWidget(browse_btn)
        layout.addWidget(path_label)
        layout.addLayout(path_row)
        layout.addWidget(self._mini_divider())
        layout.addStretch()

        self.status_label = QLabel("Add files to start")
        self.status_label.setStyleSheet("color:#555;font-size:11px;background:transparent;")

        self.merge_btn = QPushButton("🔀  Merge Files")
        self.merge_btn.setFixedHeight(44)
        self.merge_btn.setCursor(Qt.PointingHandCursor)
        self.merge_btn.clicked.connect(self._merge_files)
        self.merge_btn.setEnabled(False)
        self.merge_btn.setStyleSheet("QPushButton{background-color:#4caf81;color:white;border:none;border-radius:8px;font-size:13px;font-weight:bold;}QPushButton:hover{background-color:#3d9e70;}QPushButton:disabled{background-color:#2a2d3e;color:#555;}")

        layout.addWidget(self.status_label)
        layout.addWidget(self.merge_btn)
        return frame

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)")
        for path in paths:
            if path not in self.file_list:
                self.file_list.append(path)
                item = QListWidgetItem(f"📄  {os.path.basename(path)}")
                item.setToolTip(path)
                self.file_listbox.addItem(item)
        self._update_count()

    def _remove_selected(self):
        for item in self.file_listbox.selectedItems():
            row = self.file_listbox.row(item)
            self.file_list.pop(row)
            self.file_listbox.takeItem(row)
        self._update_count()

    def _clear_files(self):
        self.file_list.clear()
        self.file_listbox.clear()
        self._update_count()

    def _update_count(self):
        count = len(self.file_list)
        self.count_label.setText(f"{count} file{'s' if count != 1 else ''}")
        self.merge_btn.setEnabled(count >= 2)
        self.status_label.setText(
            f"Ready — {count} files selected" if count >= 2 else "Add at least 2 files to merge"
        )

    def _on_mode_changed(self, index):
        self.mode_desc.setText(
            "Each file becomes a separate sheet\nin the merged workbook." if index == 0
            else "All rows stacked vertically.\nFiles must have same columns."
        )

    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Merged File As", "merged_output.xlsx", "Excel Files (*.xlsx)")
        if path:
            self._output_path = path
            self.path_display.setText(os.path.basename(path))

    def _get_output_path(self):
        if self._output_path:
            return self._output_path
        base_dir   = os.path.dirname(self.file_list[0]) if self.file_list else "."
        output_dir = os.path.join(base_dir, "outputs")
        os.makedirs(output_dir, exist_ok=True)
        mode = "sheets" if self.mode_combo.currentIndex() == 0 else "rows"
        return os.path.join(output_dir, f"merged_{mode}.xlsx")

    def _merge_files(self):
        if len(self.file_list) < 2:
            self.result_box.setText("❌  Add at least 2 files to merge.")
            return
        mode        = "sheets" if self.mode_combo.currentIndex() == 0 else "rows"
        output_path = self._get_output_path()
        self.merge_btn.setEnabled(False)
        self.merge_btn.setText("⏳  Merging...")
        self.status_label.setText("Processing files...")
        self.result_box.setText("🔄  Merging files...\n")

        from gui.workers.agent_worker import MergerWorker
        self.worker = MergerWorker(self.file_list.copy(), output_path, mode)
        self.worker.status.connect(self._on_status)
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_status(self, msg):
        self.status_label.setText(msg)
        self.main_window.set_status(msg)

    def _on_result(self, data):
        self.merge_btn.setEnabled(True)
        self.merge_btn.setText("🔀  Merge Files")
        if data["status"] == "success":
            outpath = data.get("output_path", "")
            self.last_output = outpath
            files   = data.get("files", "?")
            mode    = data.get("mode", "sheets")
            detail  = f"  Total rows : {data.get('total_rows','?')}\n" if mode == "rows" else "  Each file as separate sheet\n"
            self.result_box.setText(
                f"✅  {data.get('message','Done!')}\n\n"
                f"── Details ─────────────────────────\n"
                f"  Files merged : {files}\n"
                f"  Mode         : {mode}\n"
                f"{detail}"
                f"  Saved to     : {os.path.basename(outpath)}\n"
                f"  Full path    : {outpath}"
            )
            self.status_label.setText("Merge complete ✓")
            self.main_window.set_status(f"Merged {files} files → {os.path.basename(outpath)} ✓")
        else:
            self.result_box.setText(f"❌  Merge failed.\n\nError: {data.get('message','Unknown error')}")
            self.status_label.setText("Merge failed ✗")

    def _on_error(self, msg):
        self.merge_btn.setEnabled(True)
        self.merge_btn.setText("🔀  Merge Files")
        self.status_label.setText("Error ✗")
        self.result_box.setText(f"❌  Error: {msg}")
        self.main_window.set_status(f"Error: {msg}")

    def on_file_loaded(self, path):
        self.current_file = path
        if path and path not in self.file_list:
            self.file_list.append(path)
            item = QListWidgetItem(f"📄  {os.path.basename(path)} (current)")
            item.setToolTip(path)
            self.file_listbox.addItem(item)
            self._update_count()

    def _divider(self):
        line = QFrame(); line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#2a2d3e;max-height:1px;border:none;")
        return line

    def _mini_divider(self):
        line = QFrame(); line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color:#1e2035;max-height:1px;border:none;")
        return line

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#444;font-size:10px;letter-spacing:1px;")
        return lbl
