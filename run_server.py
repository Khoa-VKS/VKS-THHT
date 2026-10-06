# -*- coding: utf-8 -*-
import sys
import os
import webbrowser
import threading
import time
import uvicorn

# Thêm thư mục hiện tại vào sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def open_browser():
    time.sleep(1.5)
    url = "http://127.0.0.1:8000"
    print(f"\n[INFO] Đang mở trình duyệt tại: {url}\n")
    webbrowser.open(url)

if __name__ == "__main__":
    print("=" * 65)
    print("   HỆ THỐNG BÓC TÁCH BÁO CÁO QLAHS & XUẤT EXCEL TỔNG HỢP")
    print("   KIẾN TRÚC: BACKEND (FastAPI + RapidOCR) & FRONTEND (HTML/JS)")
    print("=" * 65)

    # Chạy thread mở trình duyệt
    threading.Thread(target=open_browser, daemon=True).start()

    # Chạy uvicorn server
    uvicorn.run("backend.main:app", host="127.0.0.1", port=8000, reload=False)
