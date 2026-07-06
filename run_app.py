import os
import sys
import streamlit.web.cli as stcli
import webbrowser
import threading
import time

def open_browser():
    # ۵ ثانیه تنفس برای استقرار کامل سرور بر روی ریشه رم
    time.sleep(5)
    webbrowser.open("http://localhost:8501")

if __name__ == "__main__":
    # ردیابی دایرکتوری فیزیکی فایل اجرایی در سیستم مقصد
    if getattr(sys, 'frozen', False):
        exe_dir = sys._MEIPASS
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
        
    os.chdir(exe_dir)
    script_path = os.path.join(exe_dir, "app.py")
    
    # 📌 غیرفعال کردن اجباری developmentMode جهت رفع قطعی خطای کانفلیت پورت
    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.port=8501"
    ]
    
    print("⚡ در حال استارت پورتال هوشمند پرتابل توانیر...")
    
    # فعال‌سازی ترد بازکننده مرورگر
    threading.Thread(target=open_browser, daemon=True).start()
    
    # شلیک مستقیم هسته استریم‌لیت بدون واسطه ساب‌پراسس (لوپ و کرش غیرممکن می‌شود)
    sys.exit(stcli.main())
