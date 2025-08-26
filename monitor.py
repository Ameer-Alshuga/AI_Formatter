# monitor.py - The Clipboard Watcher

import pyperclip
import time
from main import convert_html_to_docx  # استيراد دالة التحويل الخاصة بنا

def is_valid_html(content):
    """
    فحص بسيط لمعرفة ما إذا كان محتوى الحافظة هو على الأرجح HTML الذي نريده.
    """
    content = content.strip()
    return content.startswith('<!DOCTYPE html>') or (content.startswith('<') and content.endswith('>'))

def main():
    """
    الدالة الرئيسية لتشغيل حلقة مراقبة الحافظة.
    """
    print("Starting AI Formatter...")
    print("Monitoring clipboard. Copy HTML from your AI chat to begin.")
    print("Press Ctrl+C to stop.")

    recent_content = ""
    while True:
        try:
            # الحصول على المحتوى من الحافظة
            clipboard_content = pyperclip.paste()

            # التحقق مما إذا كان جديداً وصالحاً
            if clipboard_content != recent_content and is_valid_html(clipboard_content):
                print("\nNew HTML detected on clipboard! Processing...")
                
                # تحديث المحتوى الأخير لتجنب إعادة المعالجة
                recent_content = clipboard_content
                
                # استدعاء محرك التحويل الخاص بنا
                output_filename = 'formatted_output.docx'
                convert_html_to_docx(clipboard_content, output_filename)
                
                print(f"Success! Check for the file '{output_filename}'.")
                print("Monitoring clipboard again...")

            # الانتظار قليلاً قبل التحقق مرة أخرى لتقليل استهلاك المعالج
            time.sleep(1)

        except KeyboardInterrupt:
            print("\nStopping AI Formatter. Goodbye!")
            break
        except Exception as e:
            print(f"An error occurred: {e}")
            recent_content = "" 
            time.sleep(3) # الانتظار لفترة أطول بعد حدوث خطأ

if __name__ == '__main__':
    main()