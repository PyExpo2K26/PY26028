"""
PyXcel — Chat Panel
Multi-turn conversational analysis of spreadsheet data via LLaMA.
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QFrame, QScrollArea, QSizePolicy
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont

STARTER_QUESTIONS = [
    "What is the total of all numeric columns?",
    "Which row has the highest value?",
    "Are there any missing values in this data?",
    "What is the average of the numeric columns?",
    "How many unique values are in each column?",
    "What trends do you see in this data?",
    "Which column has the most duplicate values?",
    "Give me a summary of this spreadsheet",
]


class ChatBubble(QFrame):
    def __init__(self, text, is_user, parent=None):
        super().__init__(parent)
        self.is_user = is_user
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 4)

        bubble = QLabel(text)
        bubble.setWordWrap(True)
        bubble.setTextInteractionFlags(Qt.TextSelectableByMouse)
        bubble.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        bubble.setMaximumWidth(620)

        if is_user:
            bubble.setStyleSheet("QLabel{background-color:#1e2035;color:#c0c4ff;border-radius:12px;border-bottom-right-radius:3px;padding:10px 14px;font-size:13px;}")
            avatar = QLabel("You")
            avatar.setFixedWidth(36)
            avatar.setAlignment(Qt.AlignTop | Qt.AlignCenter)
            avatar.setStyleSheet("QLabel{background-color:#7c83ff;color:white;border-radius:10px;padding:4px;font-size:10px;font-weight:bold;margin-left:8px;}")
            layout.addStretch()
            layout.addWidget(bubble)
            layout.addWidget(avatar)
        else:
            bubble.setStyleSheet("QLabel{background-color:#162820;color:#a0d4b4;border-radius:12px;border-bottom-left-radius:3px;padding:10px 14px;font-size:13px;}")
            avatar = QLabel("AI")
            avatar.setFixedWidth(36)
            avatar.setAlignment(Qt.AlignTop | Qt.AlignCenter)
            avatar.setStyleSheet("QLabel{background-color:#4caf81;color:white;border-radius:10px;padding:4px;font-size:10px;font-weight:bold;margin-right:8px;}")
            layout.addWidget(avatar)
            layout.addWidget(bubble)
            layout.addStretch()


class TypingIndicator(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 4)
        avatar = QLabel("AI")
        avatar.setFixedWidth(36)
        avatar.setAlignment(Qt.AlignCenter)
        avatar.setStyleSheet("QLabel{background-color:#4caf81;color:white;border-radius:10px;padding:4px;font-size:10px;font-weight:bold;margin-right:8px;}")
        self.dots_label = QLabel("Thinking .")
        self.dots_label.setStyleSheet("QLabel{background-color:#162820;color:#4caf81;border-radius:12px;border-bottom-left-radius:3px;padding:10px 14px;font-size:13px;}")
        layout.addWidget(avatar)
        layout.addWidget(self.dots_label)
        layout.addStretch()
        self._dot_state = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._animate)

    def start(self):
        self._timer.start(400)

    def stop(self):
        self._timer.stop()

    def _animate(self):
        dots = ["Thinking .", "Thinking ..", "Thinking ..."]
        self.dots_label.setText(dots[self._dot_state % 3])
        self._dot_state += 1


class ChatPanel(QWidget):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window  = main_window
        self.current_file = None
        self.history      = []
        self.bubbles      = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 20)
        layout.setSpacing(16)

        header_row = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(4)

        badge = QLabel("CONVERSATIONAL DATA ANALYSIS")
        badge.setStyleSheet("background-color:#1e2035;color:#7c83ff;border-radius:12px;padding:3px 12px;font-size:10px;font-weight:bold;letter-spacing:1px;")
        badge.setFixedHeight(24)

        title = QLabel("Chat with Data")
        font  = QFont(); font.setPointSize(18); font.setBold(True)
        title.setFont(font)

        subtitle = QLabel("Ask questions about your spreadsheet in plain English. LLaMA has full awareness of your data structure.")
        subtitle.setStyleSheet("color:#555;font-size:12px;")
        subtitle.setWordWrap(True)

        left.addWidget(badge)
        left.addWidget(title)
        left.addWidget(subtitle)

        self.clear_chat_btn = QPushButton("🗑  Clear Chat")
        self.clear_chat_btn.setFixedWidth(120)
        self.clear_chat_btn.setCursor(Qt.PointingHandCursor)
        self.clear_chat_btn.clicked.connect(self._clear_chat)
        self.clear_chat_btn.setStyleSheet("QPushButton{background-color:#1e2035;color:#888;border:1px solid #2a2d3e;border-radius:8px;padding:8px 14px;font-size:12px;}QPushButton:hover{background-color:#3d1a1a;color:#f44336;border-color:#5a2020;}")

        header_row.addLayout(left)
        header_row.addStretch()
        header_row.addWidget(self.clear_chat_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)
        layout.addLayout(header_row)

        self.file_status = QLabel("⚠️  No file loaded — load an Excel file to start chatting")
        self.file_status.setStyleSheet("color:#ff9800;font-size:12px;background:#1e1a0e;border-radius:6px;padding:8px 12px;")
        layout.addWidget(self.file_status)

        self.starters_widget = self._build_starters()
        layout.addWidget(self.starters_widget)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet("QScrollArea{background-color:#13151f;border:1px solid #2a2d3e;border-radius:12px;}")

        self.chat_container = QWidget()
        self.chat_container.setStyleSheet("background-color:#13151f;")
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setContentsMargins(16, 16, 16, 16)
        self.chat_layout.setSpacing(8)
        self.chat_layout.addStretch()

        self.scroll_area.setWidget(self.chat_container)
        layout.addWidget(self.scroll_area, stretch=1)

        self.typing_indicator = TypingIndicator()
        self.typing_indicator.hide()
        layout.addWidget(self.typing_indicator)

        layout.addWidget(self._build_input_bar())

        self.history_label = QLabel("0 messages  ·  context: 0 exchanges")
        self.history_label.setStyleSheet("color:#2a2d3e;font-size:11px;padding:2px 0;")
        self.history_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.history_label)

    def _build_starters(self):
        widget = QWidget()
        widget.setStyleSheet("background:transparent;")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        label = QLabel("💡  Try asking:")
        label.setStyleSheet("color:#444;font-size:11px;")
        layout.addWidget(label)

        row1 = QHBoxLayout(); row1.setSpacing(8)
        row2 = QHBoxLayout(); row2.setSpacing(8)

        for i, q in enumerate(STARTER_QUESTIONS):
            btn = QPushButton(q)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, text=q: self._use_starter(text))
            btn.setStyleSheet("QPushButton{background-color:#1a1d2e;color:#555;border:1px solid #2a2d3e;border-radius:20px;padding:6px 14px;font-size:11px;}QPushButton:hover{background-color:#1e2035;color:#c0c4ff;border-color:#7c83ff;}")
            if i < 4:
                row1.addWidget(btn)
            else:
                row2.addWidget(btn)

        layout.addLayout(row1)
        layout.addLayout(row2)
        return widget

    def _build_input_bar(self):
        bar = QFrame()
        bar.setStyleSheet("QFrame{background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:12px;}")
        bar.setFixedHeight(60)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(10)

        self.msg_input = QLineEdit()
        self.msg_input.setPlaceholderText("Ask anything about your spreadsheet data...")
        self.msg_input.setStyleSheet("QLineEdit{background:transparent;border:none;color:#e0e0e0;font-size:13px;padding:4px 0;}")
        self.msg_input.returnPressed.connect(self._send_message)

        self.send_btn = QPushButton("Send  ➤")
        self.send_btn.setFixedWidth(100)
        self.send_btn.setCursor(Qt.PointingHandCursor)
        self.send_btn.clicked.connect(self._send_message)
        self.send_btn.setStyleSheet("QPushButton{background-color:#7c83ff;color:white;border:none;border-radius:8px;padding:8px 16px;font-size:13px;font-weight:bold;}QPushButton:hover{background-color:#6b72ff;}QPushButton:disabled{background-color:#2a2d3e;color:#555;}")

        layout.addWidget(self.msg_input)
        layout.addWidget(self.send_btn)
        return bar

    def _send_message(self):
        if not self.current_file:
            self._add_system_message("⚠️  Please load an Excel file first before chatting.")
            return
        message = self.msg_input.text().strip()
        if not message:
            return

        self.starters_widget.hide()
        self._add_bubble(message, is_user=True)
        self.msg_input.clear()
        self.send_btn.setEnabled(False)
        self.msg_input.setEnabled(False)
        self.send_btn.setText("...")
        self.typing_indicator.show()
        self.typing_indicator.start()
        self.main_window.set_status("LLaMA is thinking...")

        from gui.workers.agent_worker import ChatWorker
        self.worker = ChatWorker(self.current_file, message, self.history.copy())
        self.worker.result.connect(self._on_result)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_result(self, data):
        self.typing_indicator.stop()
        self.typing_indicator.hide()
        self.send_btn.setEnabled(True)
        self.msg_input.setEnabled(True)
        self.send_btn.setText("Send  ➤")
        self.msg_input.setFocus()

        if data["status"] == "success":
            response = data.get("response", "")
            message  = data.get("message",  "")
            self._add_bubble(response, is_user=False)
            self.history.append({"role": "user",      "content": message})
            self.history.append({"role": "assistant",  "content": response})
            if len(self.history) > 20:
                self.history = self.history[-20:]
            exchanges = len(self.history) // 2
            self.history_label.setText(f"{len(self.history)} messages  ·  context: {exchanges} exchanges")
            self.main_window.set_status("Response received ✓")
        else:
            self._add_system_message("❌  Error getting response from LLaMA.")

    def _on_error(self, msg):
        self.typing_indicator.stop()
        self.typing_indicator.hide()
        self.send_btn.setEnabled(True)
        self.msg_input.setEnabled(True)
        self.send_btn.setText("Send  ➤")
        self._add_system_message(f"❌  Error: {msg}\n\nMake sure Ollama is running: ollama serve")
        self.main_window.set_status(f"Error: {msg}")

    def _add_bubble(self, text, is_user):
        bubble = ChatBubble(text, is_user)
        count  = self.chat_layout.count()
        self.chat_layout.insertWidget(count - 1, bubble)
        self.bubbles.append(bubble)
        self._scroll_to_bottom()

    def _add_system_message(self, text):
        label = QLabel(text)
        label.setAlignment(Qt.AlignCenter)
        label.setWordWrap(True)
        label.setStyleSheet("QLabel{color:#555;font-size:12px;background-color:#1a1d2e;border:1px solid #2a2d3e;border-radius:8px;padding:10px 16px;margin:4px 40px;}")
        count = self.chat_layout.count()
        self.chat_layout.insertWidget(count - 1, label)
        self._scroll_to_bottom()

    def _scroll_to_bottom(self):
        QTimer.singleShot(50, lambda: self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        ))

    def _use_starter(self, text):
        self.msg_input.setText(text)
        self.msg_input.setFocus()

    def _clear_chat(self):
        for bubble in self.bubbles:
            self.chat_layout.removeWidget(bubble)
            bubble.deleteLater()
        self.bubbles.clear()
        self.history.clear()
        while self.chat_layout.count() > 1:
            item = self.chat_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.history_label.setText("0 messages  ·  context: 0 exchanges")
        self.starters_widget.show()
        self.main_window.set_status("Chat cleared")

    def on_file_loaded(self, path):
        self.current_file = path
        filename = os.path.basename(path)
        self.file_status.setText(f"✅  Chatting about: {filename}")
        self.file_status.setStyleSheet("color:#4caf81;font-size:12px;background:#0e1e14;border-radius:6px;padding:8px 12px;")
        self._clear_chat()
        self._add_system_message(f"📄  {filename} loaded — Ask me anything about your data!")
