# -*- coding: utf-8 -*-
import os
import glob
import re
import openpyxl
from PIL import Image
import numpy as np
from rapidocr_onnxruntime import RapidOCR
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

engine = RapidOCR()

input_dir = r"C:\Users\VKS\Downloads\LỖI NỀN TẢNG QLAHS\Lý Phương Nam [04-10-2026 10_35]"
out_file = r"C:\Users\VKS\.gemini\antigravity\scratch\tach_noi_dung_excel\TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx"

def clean_yeu_cau(raw_text):
    if not raw_text:
        return ""
    text = raw_text.strip()
    
    # 1. Tìm NỘI DUNG YÊU CẦU và cắt bỏ toàn bộ phần đứng trước nó
    m = re.search(r'(?:.*)(?:n[ộoô]i\s*dung\s*)?y[êe]u\s*c[ầa]u\s*[:\-\.]?\s*', text, re.IGNORECASE | re.DOTALL)
    if m:
        text = text[m.end():]
    else:
        text = re.sub(r'^[^\n]*(?:diễn biến xử lý|dien bien xu ly)[^\n]*\s*', '', text, flags=re.IGNORECASE)

    # 2. Xóa các ký tự rác đầu câu
    text = re.sub(r'^[:\-\.\s\t“"\'\(\)0-9A-Za-z\/\,\_\+\=]+\s*(?:NA|NG|Cự|TM)\b\s*', '', text)
    text = text.lstrip(': -.\t\n“"\'()').strip()

    # 3. Cắt bỏ nếu dính chữ Biện pháp xử lý
    bp_idx = text.lower().find("biện pháp xử lý")
    if bp_idx != -1:
        text = text[:bp_idx].strip()
    bp_idx_nodau = text.lower().find("bien phap xu ly")
    if bp_idx_nodau != -1:
        text = text[:bp_idx_nodau].strip()

    # Nối câu liền mạch
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    return " ".join(lines)

def clean_bien_phap(raw_text):
    if not raw_text:
        return ""
    text = raw_text.strip()

    # 1. Cắt bỏ chữ Biện pháp xử lý
    text = re.sub(r'^[^\n]*(?:biện pháp xử lý|bien phap xu ly)\s*[:\-\.]?\s*', '', text, flags=re.IGNORECASE).strip()

    # 2. Xóa rác đuôi: icon bình luận, .SEaS, CỔ L4...
    text = re.sub(r'[\s\.\,\_\-]+(?:\.SEaS|SEaS|CỔ L4|mkl|ki mkl)[\s\S]*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(\.)\s*[\.\-\_][\w\s\.\,\/À-ỹ]+$', r'\1', text)

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    while lines and (len(lines[-1]) <= 4 or lines[-1] in ['---', '___', '===', '💬', '[]', 'Q', 'O']):
        lines.pop()

    # Xử lý đánh số và ngắt dòng
    bullet_regex = re.compile(r'^(\d+[\.\/\)]|\-|\+|\*|•|[a-zA-Z]\))\s*')
    items = []
    current = ""
    for line in lines:
        if bullet_regex.match(line):
            if current:
                items.append(current.strip())
            current = line
        else:
            if not current:
                current = line
            else:
                current += " " + line
    if current:
        items.append(current.strip())

    return "\n".join(items) if len(items) > 1 else " ".join(lines)

def process_image(img_path):
    img = Image.open(img_path)
    arr = np.array(img.convert('RGB'))
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

        # Cắt bỏ 8% đỉnh ảnh (bỏ thanh trạng thái TM..., pin)
        top_start = int(img.height * 0.08)
        top_crop = img.crop((0, top_start, img.width, max(20, min_y - 4)))

        # Cắt bỏ 12% đáy ảnh (bỏ icon bình luận 💬)
        bot_end = int(img.height * 0.88)
        bot_crop = img.crop((0, max_y + 4, img.width, max(20, bot_end)))

        res_top, _ = engine(np.array(top_crop))
        res_bot, _ = engine(np.array(bot_crop))

        top_text = " ".join([l[1] for l in res_top]) if res_top else ""
        bot_text = "\n".join([l[1] for l in res_bot]) if res_bot else ""

        return clean_yeu_cau(top_text), clean_bien_phap(bot_text)
    else:
        res, _ = engine(arr)
        if not res:
            return "", ""
        full_text = "\n".join([l[1] for l in res])
        bp_idx = full_text.lower().find("biện pháp xử lý")
        if bp_idx != -1:
            return clean_yeu_cau(full_text[:bp_idx]), clean_bien_phap(full_text[bp_idx:])
        return clean_yeu_cau(full_text), ""

# Quét toàn bộ ảnh trong thư mục
img_files = sorted(glob.glob(os.path.join(input_dir, "*.jpg")))
print(f"Tim thay {len(img_files)} anh.")

records = []
for idx, f in enumerate(img_files, 1):
    fname = os.path.basename(f)
    print(f"[{idx}/{len(img_files)}] Xu ly: {fname}...", end="", flush=True)
    yc, bp = process_image(f)
    records.append({
        "stt": idx,
        "file": fname,
        "yeucau": yc,
        "bienphap": bp
    })
    print(" XONG")

# Xuất ra file Excel thẩm mỹ cao
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Lỗi Thường Gặp QLAHS"
ws.freeze_panes = "A2"

headers = ["STT", "TÊN / NGUỒN ẢNH", "NỘI DUNG YÊU CẦU", "BIỆN PHÁP XỬ LÝ"]
ws.append(headers)

for r in records:
    ws.append([r["stt"], r["file"], r["yeucau"], r["bienphap"]])

header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
data_font = Font(name="Segoe UI", size=10, color="1E293B")

thin = Side(border_style="thin", color="CBD5E1")
border = Border(top=thin, left=thin, right=thin, bottom=thin)

ws.row_dimensions[1].height = 34
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

for r_idx in range(2, len(records) + 2):
    is_odd = (r_idx % 2 == 1)
    row_fill = zebra_fill if is_odd else white_fill
    row_cells = ws[r_idx]
    
    yc_len = len(str(row_cells[2].value or ""))
    bp_len = len(str(row_cells[3].value or ""))
    max_lines = max(yc_len // 48 + 1, bp_len // 48 + 1, str(row_cells[3].value or "").count("\n") + 1)
    ws.row_dimensions[r_idx].height = max(50, max_lines * 19 + 10)

    for c_idx, cell in enumerate(row_cells):
        cell.font = data_font
        cell.fill = row_fill
        cell.border = border
        if c_idx == 0:
            cell.alignment = Alignment(horizontal="center", vertical="top")
        elif c_idx == 1:
            cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

ws.column_dimensions["A"].width = 8
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 60
ws.column_dimensions["D"].width = 60

wb.save(out_file)
# Copy sang Ket_Qua_Mau.xlsx
wb.save(r"C:\Users\VKS\.gemini\antigravity\scratch\tach_noi_dung_excel\Ket_Qua_Mau.xlsx")
# Copy sang ca thu muc Downloads cua nguoi dung de mo truc tiep
wb.save(os.path.join(input_dir, "TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx"))
print("Da xuat thanh cong toan bo 12 anh ra file Excel!")
