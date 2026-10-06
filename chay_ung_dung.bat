@echo off
chcp 65001 > nul
title HỆ THỐNG BÓC TÁCH BÁO CÁO QLAHS & XUẤT EXCEL
echo ===================================================================
echo     HỆ THỐNG BÓC TÁCH BÁO CÁO QLAHS & XUẤT EXCEL TỔNG HỢP
echo     (Tách riêng Backend AI RapidOCR và Frontend Giao Diện Web)
echo ===================================================================
echo.
echo [1/2] Đang khởi động Backend FastAPI & RapidOCR...
echo [2/2] Tự động mở giao diện Frontend trên trình duyệt Web...
echo.

cd /d "%~dp0"
python run_server.py

pause
