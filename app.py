# app.py - The Main GUI Application

import sys
import pyperclip
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QTextEdit
from PyQt6.QtCore import QTimer, Qt

# Import our conversion engine from main.py
from main import convert_html_to_docx

class AIFormatterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("المنسق الذكي (AI Formatter)")
        self.setGeometry(100, 100, 500, 400) # x, y, width, height

        # --- UI Elements ---
        self.status_label = QLabel("الحالة: جاهز للنسخ...", self)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.status_label.font()
        font.setPointSize(12)
        self.status_label.setFont(font)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True) # Make it non-editable
        self.log_box.append("مرحباً بك في المنسق الذكي!")
        
        # --- Layout ---
        layout = QVBoxLayout()
        layout.addWidget(self.status_label)
        layout.addWidget(self.log_box)
        
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # --- Clipboard Monitoring Logic ---
        self.recent_content = ""
        
        # We use QTimer to check the clipboard periodically without freezing the app
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_clipboard)
        self.timer.start(1000) # Check every 1000 milliseconds = 1 second

    def is_valid_html(self, content):
        """
        A simple check to see if the clipboard content is likely the HTML we want.
        """
        content = content.strip()
        return content.startswith('<!DOCTYPE html>') or (content.startswith('<') and content.endswith('>'))

    def check_clipboard(self):
        """
        This function is called by the QTimer every second.
        """
        try:
            clipboard_content = pyperclip.paste()

            if clipboard_content != self.recent_content and self.is_valid_html(clipboard_content):
                self.status_label.setText("الحالة: تم اكتشاف محتوى جديد، جاري المعالجة...")
                self.log_box.append("تم اكتشاف HTML جديد في الحافظة...")
                QApplication.processEvents() # Update the UI immediately

                self.recent_content = clipboard_content
                
                # --- Call our conversion engine ---
                output_filename = 'formatted_output.docx'
                convert_html_to_docx(clipboard_content, output_filename)
                
                self.log_box.append(f"-> نجاح! تم حفظ الملف باسم '{output_filename}'.")
                self.status_label.setText("الحالة: جاهز للنسخ...")
        
        except Exception as e:
            self.log_box.append(f"حدث خطأ: {e}")
            self.status_label.setText("الحالة: حدث خطأ، يرجى المحاولة مرة أخرى.")
            # Optional: reset recent_content on error
            self.recent_content = ""

# --- Application Entry Point ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AIFormatterApp()
    window.show()
    sys.exit(app.exec())