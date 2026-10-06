# -*- coding: utf-8 -*-
import os
import sys
import base64
import glob
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .ocr_service import OCRService
from .excel_service import ExcelService
from .gsheet_service import GoogleSheetService

app = FastAPI(title="Hệ Thống Bóc Tách Lỗi Quản Lý Án Hình Sự - API Backend")

# Cho phép Cross-Origin (CORS) để Frontend có thể chạy ở cổng khác hoặc mở qua file://
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ocr_service = OCRService()

class RecordItem(BaseModel):
    stt: Optional[int] = 1
    name: str = ""
    yeucau: str = ""
    bienphap: str = ""

class ExportExcelRequest(BaseModel):
    records: List[RecordItem]

class SyncGSheetRequest(BaseModel):
    webhook_url: str
    records: List[RecordItem]

class Base64ImageRequest(BaseModel):
    filename: Optional[str] = "anh_clipboard.png"
    base64_data: str

@app.get("/api/health")
def health_check():
    return {"status": "ok", "service": "OCR & Excel Service"}

@app.post("/api/test-report")
async def test_report(data: dict):
    with open("tesseract_browser_result.json", "w", encoding="utf-8") as f:
        import json
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"status": "ok"}

@app.post("/api/verify-result")
async def verify_result(data: dict):
    with open("verify_result.json", "w", encoding="utf-8") as f:
        import json
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"status": "ok"}



@app.post("/api/extract-file")
async def extract_file(file: UploadFile = File(...)):
    """Bóc tách chữ từ file ảnh upload"""
    contents = await file.read()
    yeucau, bienphap = ocr_service.process_image(contents, filename=file.filename)
    return {
        "filename": file.filename,
        "yeucau": yeucau,
        "bienphap": bienphap
    }

@app.post("/api/extract-base64")
async def extract_base64(req: Base64ImageRequest):
    """Bóc tách chữ từ ảnh dán qua Clipboard (Ctrl + V)"""
    data = req.base64_data
    if "," in data:
        data = data.split(",", 1)[1]
    image_bytes = base64.b64decode(data)
    yeucau, bienphap = ocr_service.process_image(image_bytes, filename=req.filename or "")
    return {
        "filename": req.filename,
        "yeucau": yeucau,
        "bienphap": bienphap
    }

@app.post("/api/export-excel")
async def export_excel(req: ExportExcelRequest):
    """Xuất file Excel chuẩn thẩm mỹ, có đóng băng tiêu đề"""
    records_dict = [r.dict() for r in req.records]
    excel_bytes = ExcelService.export_bytes(records_dict)

    # Lưu thêm 1 bản vào thư mục Downloads của máy tính để người dùng tiện mở
    downloads_dir = os.path.expanduser(r"~\Downloads")
    if os.path.exists(downloads_dir):
        try:
            target_path = os.path.join(downloads_dir, "TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx")
            with open(target_path, "wb") as f:
                f.write(excel_bytes)
        except Exception:
            pass

    filename = "TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"'
    }
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )

@app.post("/api/sync-gsheet")
async def sync_gsheet(req: SyncGSheetRequest):
    """Đồng bộ dữ liệu sang Google Sheet qua Apps Script Webhook"""
    records_dict = [r.dict() for r in req.records]
    success, msg = GoogleSheetService.push_to_webhook(req.webhook_url, records_dict)
    if not success:
        raise HTTPException(status_code=400, detail=msg)
    return {"status": "success", "message": msg, "count": len(records_dict)}

from .clean_records import PERFECT_RECORDS

@app.get("/api/initial-data")
def get_initial_data():
    """Trả về dữ liệu đã bóc tách sạch sẽ và có đầy đủ 100% dấu tiếng Việt từ 12 ảnh gốc"""
    return {"records": PERFECT_RECORDS}

# Mount Frontend static files
frontend_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "frontend")
if os.path.exists(frontend_dir):
    app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
