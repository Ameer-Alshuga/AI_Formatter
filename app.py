# app.py - The Final User-Friendly GUI Application

import sys
import pyperclip
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QTextEdit, QMessageBox, QPushButton, QDialog
from PyQt6.QtCore import QTimer, Qt
from main import convert_html_to_docx

### --- NEW: The Setup Dialog Window --- ###
class SetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("إعداد المتصفح لأول مرة")
        self.setMinimumSize(600, 450)

        self.bookmarklet_code = """javascript:void(function(){try{let selection=window.getSelection();if(selection.rangeCount>0&&selection.toString().trim()!==''){let range=selection.getRangeAt(0);let div=document.createElement("div");div.appendChild(range.cloneContents());let htmlContent='<!DOCTYPE html><html lang="ar"><head><meta charset="UTF--8"></head><body>'+div.innerHTML+'</body></html>';let textarea=document.createElement("textarea");textarea.value=htmlContent;document.body.appendChild(textarea);textarea.select();document.execCommand("copy");document.body.removeChild(textarea);alert("AI Formatter:\\nتم نسخ بنية HTML بنجاح!")}else{alert("AI Formatter:\\nيرجى تحديد نص يحتوي على محتوى أولاً!")}}catch(e){alert("AI Formatter:\\nحدث خطأ: "+e.message)}})();"""

        layout = QVBoxLayout(self)

        # Instructions
        instructions_label = QLabel(
            "<h3>لربط البرنامج بالمتصفح، اتبع هذه الخطوات السهلة:</h3>"
            "<ol>"
            "<li><b>انسخ الكود التالي بالكامل</b> بالضغط على الزر أدناه.</li>"
            "<li>في متصفحك (Chrome, Firefox, etc)، <b>أنشئ إشارة مرجعية جديدة</b> (New Bookmark).</li>"
            "<li>في حقل 'الاسم' (Name)، اكتب اسماً من اختيارك (مثلاً: <b>✨ المنسق الذكي</b>).</li>"
            "<li>في حقل 'العنوان' أو 'URL'، <b>الصق الكود المنسوخ</b>.</li>"
            "</ol>"
            "هذا كل شيء! الآن يمكنك استخدام هذه الإشارة المرجعية لنسخ المحتوى المنسق."
        )
        instructions_label.setWordWrap(True)

        # Code Box
        self.code_box = QTextEdit()
        self.code_box.setPlainText(self.bookmarklet_code)
        self.code_box.setReadOnly(True)
        self.code_box.setFixedHeight(120)

        # Copy Button
        self.copy_button = QPushButton("نسخ الكود")
        self.copy_button.clicked.connect(self.copy_code)

        layout.addWidget(instructions_label)
        layout.addWidget(self.code_box)
        layout.addWidget(self.copy_button)
        self.setLayout(layout)

    def copy_code(self):
        pyperclip.copy(self.bookmarklet_code)
        self.copy_button.setText("✓ تم النسخ بنجاح!")
        # Revert button text after 2 seconds
        QTimer.singleShot(2000, lambda: self.copy_button.setText("نسخ الكود"))


class AIFormatterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("المنسق الذكي (AI Formatter)")
        self.setGeometry(100, 100, 500, 400)
        
        self.status_label = QLabel("الحالة: جاهز للنسخ...", self)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.status_label.font(); font.setPointSize(12); self.status_label.setFont(font)
        
        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.append("مرحباً بك في المنسق الذكي!")
        
        ### --- NEW: The Setup Button --- ###
        self.setup_button = QPushButton("إعداد المتصفح / عرض التعليمات")
        self.setup_button.clicked.connect(self.open_setup_dialog)
        
        layout = QVBoxLayout()
        layout.addWidget(self.status_label)
        layout.addWidget(self.log_box)
        layout.addWidget(self.setup_button) # Add button to layout
        
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        
        self.recent_content = ""
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_clipboard)
        self.timer.start(1000)

    def is_valid_html(self, content):
        content = content.strip()
        return content.startswith('<!DOCTYPE html>') or (content.startswith('<') and content.endswith('>'))

    def check_clipboard(self):
        # This function remains the same as the last version
        try:
            clipboard_content = pyperclip.paste()
            if clipboard_content != self.recent_content and self.is_valid_html(clipboard_content):
                self.status_label.setText("الحالة: تم اكتشاف محتوى جديد، جاري المعالجة...")
                self.log_box.append("تم اكتشاف HTML جديد في الحافظة...")
                QApplication.processEvents()
                
                self.recent_content = clipboard_content
                output_filename = 'formatted_output.docx'
                
                success = convert_html_to_docx(clipboard_content, output_filename)
                
                if success:
                    self.log_box.append(f"-> نجاح! تم تحديث الملف '{output_filename}'.")
                else:
                    self.log_box.append(f"-> فشل! لا يمكن الكتابة على الملف '{output_filename}'.")
                    self.show_error_popup("خطأ في الحفظ", "لم يتمكن البرنامج من حفظ الملف.\n\nالرجاء التأكد من أن ملف 'formatted_output.docx' مغلق ثم حاول مرة أخرى.")
                
                self.status_label.setText("الحالة: جاهز للنسخ...")
        
        except Exception as e:
            self.log_box.append(f"حدث خطأ غير متوقع: {e}")
            self.show_error_popup("خطأ غير متوقع", f"حدث خطأ:\n{e}")

    ### --- NEW: Function to open the setup dialog --- ###
    def open_setup_dialog(self):
        dialog = SetupDialog(self)
        dialog.exec()

    def show_error_popup(self, title, message):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AIFormatterApp()
    window.show()
    sys.exit(app.exec())