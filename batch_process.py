# -*- coding: utf-8 -*-
"""
Script tự động bóc tách hàng loạt ảnh trong thư mục input_images/
và tổng hợp vào file Excel Ket_Qua_Tong_Hop.xlsx
(Đã tích hợp nhận diện màu xanh Biện pháp xử lý & Định dạng Excel thẩm mỹ, đóng băng tiêu đề)
"""

import os
import re
import glob
import json
import base64
import openpyxl
from PIL import Image
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fix_excel import unwrap_text
from style_excel import format_excel_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input_images")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "Ket_Qua_Tong_Hop.xlsx")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

def load_api_key():
    key = os.environ.get("GEMINI_API_KEY")
    if key:
        return key
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                return cfg.get("gemini_api_key", "").strip()
        except Exception:
            pass
    return ""

def process_image_with_gemini(image_path, api_key):
    import requests
    
    with open(image_path, "rb") as f:
        img_bytes = f.read()
    b64_data = base64.b64encode(img_bytes).decode("utf-8")
    
    ext = os.path.splitext(image_path)[1].lower()
    mime = "image/png" if ext == ".png" else "image/jpeg"

    prompt = """Bạn là chuyên gia bóc tách dữ liệu từ ảnh chụp màn hình báo cáo pháp luật/vụ án.

HƯỚNG DẪN ĐỊNH VỊ CỰC KỲ QUAN TRỌNG:
1. Dòng chữ "Biện pháp xử lý" LUÔN CÓ MÀU XANH LÁ CÂY (GREEN) nằm ở giữa ảnh.
2. PHẦN 1 - "NỘI DUNG YÊU CẦU":
   - Toàn bộ văn bản nằm phía trên dòng chữ màu xanh lá (dưới tiêu đề NỘI DUNG YÊU CẦU).
   - Nối liền các câu thành 1 đoạn duy nhất, KHÔNG ĐƯỢC TỰ Ý XUỐNG DÒNG trong phần này.
3. PHẦN 2 - "BIỆN PHÁP XỬ LÝ":
   - Toàn bộ văn bản nằm từ dòng chữ màu xanh lá trở xuống (các mục được đánh số 1., 2...).
   - Chỉ xuống dòng khi chuyển sang mục mới (mục 1 và mục 2 nằm ở 2 dòng riêng).
   - TUYỆT ĐỐI KHÔNG BỎ QUÊN MỤC NÀO VÀ KHÔNG ĐƯỢC GỘP VÀO NỘI DUNG YÊU CẦU.

Giữ nguyên vẹn 100% tiếng Việt có dấu, đúng ngữ pháp, số điều luật, thời hạn.
Trả về DUY NHẤT một JSON hợp lệ:
{
  "yeu_cau": "Nội dung phần NỘI DUNG YÊU CẦU",
  "bien_phap": "Nội dung phần Biện pháp xử lý (gồm mục 1, mục 2...)"
}"""

    models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime, "data": b64_data}}
                ]
            }],
            "generationConfig": {
                "response_mime_type": "application/json"
            }
        }
        resp = requests.post(url, json=payload, timeout=30)
        if resp.status_code == 200:
            res_data = resp.json()
            text = res_data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
            clean = text.replace("```json", "").replace("```", "").strip()
            parsed = json.loads(clean)
            yc = unwrap_text(parsed.get("yeu_cau", ""), is_numbered=False)
            bp = unwrap_text(parsed.get("bien_phap", ""), is_numbered=True)
            return yc, bp
    
    raise RuntimeError("Không thể gọi API Gemini qua các model.")

def process_image_with_rapidocr(image_path):
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR()

    img = Image.open(image_path)
    arr = np.array(img.convert("RGB"))
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]

    # Quét pixel xanh lá chữ "Biện pháp xử lý"
    green_mask = (g > 65) & (g > r.astype(int) + 16) & (g > b.astype(int) + 16)
    green_mask[:, int(img.width * 0.6):] = False
    green_mask[:int(img.height * 0.15), :] = False
    green_mask[int(img.height * 0.85):, :] = False
    y_idx, _ = np.where(green_mask)

    if len(y_idx) >= 15:
        min_y = int(np.min(y_idx))
        max_y = int(np.max(y_idx))

        top_crop = img.crop((0, 0, img.width, max(20, min_y - 4)))
        bot_crop = img.crop((0, min(img.height - 20, max_y + 4), img.width, img.height))

        res_top, _ = engine(np.array(top_crop))
        res_bot, _ = engine(np.array(bot_crop))

        yc_lines = []
        if res_top:
            for l in res_top:
                txt = l[1]
                lower = txt.lower()
                if "dien bien" not in lower and "noi dung yeu cau" not in lower and "nội dung yêu cầu" not in lower:
                    yc_lines.append(txt)

        bp_lines = []
        if res_bot:
            for l in res_bot:
                txt = l[1]
                lower = txt.lower()
                if "bien phap xu ly" not in lower and "biện pháp xử lý" not in lower:
                    bp_lines.append(txt)

        return unwrap_text(" ".join(yc_lines), is_numbered=False), unwrap_text("\n".join(bp_lines), is_numbered=True)
    else:
        result, _ = engine(arr)
        if not result:
            return "", ""
        lines = [item[1] for item in result]

        yeu_cau_lines = []
        bien_phap_lines = []
        is_bien_phap = False
        bullet_regex = re.compile(r"^(\d+[\.\/\)]|\b1\.\b|1\b)")

        for line in lines:
            lower = line.lower()
            if "bien phap xu ly" in lower or "biện pháp xử lý" in lower or "bien phap" in lower:
                is_bien_phap = True
                continue
            if bullet_regex.match(line) and yeu_cau_lines:
                is_bien_phap = True

            if is_bien_phap:
                bien_phap_lines.append(line)
            else:
                if "dien bien" not in lower and "noi dung yeu cau" not in lower:
                    yeu_cau_lines.append(line)

        return unwrap_text(" ".join(yeu_cau_lines), is_numbered=False), unwrap_text("\n".join(bien_phap_lines), is_numbered=True)

def export_to_excel(records, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tổng Hợp Xử Lý"

    headers = ["STT", "TÊN FILE ẢNH", "NỘI DUNG YÊU CẦU", "BIỆN PHÁP XỬ LÝ"]
    ws.append(headers)

    for idx, r in enumerate(records, 1):
        yc_clean = unwrap_text(r["yeu_cau"], is_numbered=False)
        bp_clean = unwrap_text(r["bien_phap"], is_numbered=True)
        row_data = [idx, r["file"], yc_clean, bp_clean]
        ws.append(row_data)

    wb.save(output_path)
    # Áp dụng bộ định dạng thẩm mỹ & đóng băng tiêu đề
    format_excel_file(output_path)
    print(f"\n-> Đã lưu và định dạng hoàn hảo file Excel tại:\n   {output_path}")

def main():
    print("=" * 60)
    print(" CÔNG CỤ TỰ ĐỘNG BÓC TÁCH ẢNH BÁO CÁO & XUẤT EXCEL")
    print(" (Tự động đóng băng tiêu đề, căn lề và định dạng thẩm mỹ)")
    print("=" * 60)
    
    os.makedirs(INPUT_DIR, exist_ok=True)
    api_key = load_api_key()

    image_exts = ("*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp")
    image_files = []
    for ext in image_exts:
        image_files.extend(glob.glob(os.path.join(INPUT_DIR, ext)))
        image_files.extend(glob.glob(os.path.join(INPUT_DIR, ext.upper())))
    image_files = sorted(list(set(image_files)))

    if not image_files:
        print(f"\n[!] Chưa có file ảnh nào trong thư mục:\n    {INPUT_DIR}")
        print("    Vui lòng copy các file ảnh cần bóc tách vào thư mục trên rồi chạy lại!")
        return

    print(f"\nTìm thấy {len(image_files)} ảnh cần xử lý.")
    if api_key:
        print("[*] Đang sử dụng chế độ: Gemini AI Vision (Độ chính xác tiếng Việt 100%)")
    else:
        print("[*] Đang sử dụng chế độ: Quét pixel màu xanh lá 'Biện pháp xử lý' + RapidOCR nội bộ")

    records = []
    for i, img_path in enumerate(image_files, 1):
        filename = os.path.basename(img_path)
        print(f"[{i}/{len(image_files)}] Đang xử lý: {filename} ...", end="", flush=True)
        try:
            if api_key:
                yeu_cau, bien_phap = process_image_with_gemini(img_path, api_key)
            else:
                yeu_cau, bien_phap = process_image_with_rapidocr(img_path)
            
            records.append({
                "file": filename,
                "yeu_cau": yeu_cau,
                "bien_phap": bien_phap
            })
            print(" XONG")
        except Exception as e:
            print(f" LỖI: {e}")

    if records:
        export_to_excel(records, OUTPUT_EXCEL)

if __name__ == "__main__":
    main()
