# -*- coding: utf-8 -*-
import os
import sys
import json
import base64
import webbrowser
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageGrab
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fix_excel import unwrap_text, clean_excel_file
from style_excel import format_excel_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
DEFAULT_EXCEL = os.path.join(BASE_DIR, "Ket_Qua_Tong_Hop.xlsx")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"gemini_api_key": ""}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Lỗi lưu config:", e)

class ToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bóc Tách Báo Cáo Xử Lý & Xuất Excel (Đã Định Dạng Thẩm Mỹ & Đóng Băng Tiêu Đề)")
        self.geometry("1150x740")
        self.minsize(950, 620)

        self.cfg = load_config()
        self.items_data = []
        
        self.init_ui()
        self.bind_events()

    def init_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Top Bar: API Key & Modes
        top_frame = ttk.LabelFrame(self, text=" Cấu Hình Nhận Diện ", padding=10)
        top_frame.pack(fill="x", padx=15, pady=(10, 5))

        lbl_key = ttk.Label(top_frame, text="Gemini API Key (Tùy chọn):")
        lbl_key.pack(side="left", padx=(0, 5))

        self.ent_key = ttk.Entry(top_frame, width=40, show="*")
        self.ent_key.insert(0, self.cfg.get("gemini_api_key", ""))
        self.ent_key.pack(side="left", padx=5)

        btn_save_key = ttk.Button(top_frame, text="Lưu Key", command=self.on_save_key)
        btn_save_key.pack(side="left", padx=5)

        btn_fix_excel = ttk.Button(top_frame, text="✨ Làm Đẹp & Đóng Băng File Excel", command=self.on_fix_existing_excel)
        btn_fix_excel.pack(side="right", padx=5)

        btn_web = ttk.Button(top_frame, text="🌐 Mở Bản Web App", command=self.open_web_app)
        btn_web.pack(side="right", padx=5)

        # Action Toolbar
        action_frame = ttk.Frame(self, padding=5)
        action_frame.pack(fill="x", padx=15, pady=5)

        btn_add_files = ttk.Button(action_frame, text="🖼️ Chọn File Ảnh...", command=self.on_add_files)
        btn_add_files.pack(side="left", padx=(0, 6))

        btn_add_dir = ttk.Button(action_frame, text="📁 Chọn Thư Mục Ảnh...", command=self.on_add_dir)
        btn_add_dir.pack(side="left", padx=6)

        btn_paste = ttk.Button(action_frame, text="📋 Dán từ Clipboard (Ctrl+V)", command=self.on_paste_clipboard)
        btn_paste.pack(side="left", padx=6)

        btn_clear = ttk.Button(action_frame, text="🗑️ Xóa Danh Sách", command=self.on_clear_list)
        btn_clear.pack(side="left", padx=6)

        self.btn_export = ttk.Button(action_frame, text="📥 Xuất File Excel Đẹp (.xlsx)", command=self.on_export_excel)
        self.btn_export.pack(side="right", padx=(6, 0))

        self.btn_run = ttk.Button(action_frame, text="⚡ Bắt Đầu Bóc Tách", command=self.on_start_process)
        self.btn_run.pack(side="right", padx=6)

        # Status & Progress
        self.status_var = tk.StringVar(value="Sẵn sàng. Tự động nhận diện chữ màu xanh Biện pháp xử lý, đóng băng tiêu đề Excel.")
        status_bar = ttk.Frame(self, padding=(15, 2))
        status_bar.pack(fill="x")
        
        self.lbl_status = ttk.Label(status_bar, textvariable=self.status_var, font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack(side="left")

        self.prog_bar = ttk.Progressbar(status_bar, mode="determinate", length=200)
        self.prog_bar.pack(side="right")

        # Main Split Area
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=15, pady=5)

        # Treeview Table
        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3)

        columns = ("stt", "file", "yeu_cau", "bien_phap")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("stt", text="STT")
        self.tree.heading("file", text="Tên Ảnh / Nguồn")
        self.tree.heading("yeu_cau", text="NỘI DUNG YÊU CẦU")
        self.tree.heading("bien_phap", text="BIỆN PHÁP XỬ LÝ (Màu xanh)")

        self.tree.column("stt", width=50, anchor="center")
        self.tree.column("file", width=180, anchor="w")
        self.tree.column("yeu_cau", width=440, anchor="w")
        self.tree.column("bien_phap", width=440, anchor="w")

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

        # Detail Editor Frame
        edit_frame = ttk.LabelFrame(paned, text=" Xem & Chỉnh Sửa Dòng Đang Chọn ", padding=8)
        paned.add(edit_frame, weight=2)

        edit_grid = ttk.Frame(edit_frame)
        edit_grid.pack(fill="both", expand=True)

        lbl_yc = ttk.Label(edit_grid, text="Nội Dung Yêu Cầu (Liền mạch):", font=("Segoe UI", 9, "bold"))
        lbl_yc.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 2))
        self.txt_yeu_cau = tk.Text(edit_grid, height=5, wrap="word", font=("Segoe UI", 10))
        self.txt_yeu_cau.grid(row=1, column=0, sticky="nsew", padx=5)

        lbl_bp = ttk.Label(edit_grid, text="Biện Pháp Xử Lý (Chữ màu xanh):", font=("Segoe UI", 9, "bold"))
        lbl_bp.grid(row=0, column=1, sticky="w", padx=5, pady=(0, 2))
        self.txt_bien_phap = tk.Text(edit_grid, height=5, wrap="word", font=("Segoe UI", 10))
        self.txt_bien_phap.grid(row=1, column=1, sticky="nsew", padx=5)

        edit_grid.columnconfigure(0, weight=1)
        edit_grid.columnconfigure(1, weight=1)
        edit_grid.rowconfigure(1, weight=1)

        btn_update_cell = ttk.Button(edit_frame, text="💾 Lưu Chỉnh Sửa Dòng Này", command=self.on_save_row_edit)
        btn_update_cell.pack(anchor="e", pady=(6, 0))

        # Thêm 2 vụ án mẫu
        self.add_record(
            "Vụ án 1 (Điều 251)",
            "Vụ án và bị can khởi tố theo khoản 2 Điều 251 là tội rất nghiêm trọng, thời hạn truy tố là 29 ngày nhưng sau khi chuyển sang giai đoạn truy tố chỉ có thời hạn 19 ngày.",
            "1. Không kiểm tra, điều chỉnh loại tội trước khi chuyển giai đoạn truy tố. Hiện vụ án vẫn loại tội NT.\n2. Chỉ liên hệ khi vụ án trên nền tảng đã quá hạn."
        )
        self.add_record(
            "Vụ án 2 (Nguyễn Xuân Thủy)",
            "Vụ án đã chuyển đến VKS Thành phố giải quyết theo thẩm quyền nhưng Viện kiểm sát thành phố nhận được vụ án nhưng không nhận được biện pháp ngăn chặn của bị can. Đề nghị chuyển biện pháp ngăn chặn của bị can Nguyễn Xuân Thủy cho VKS Thành phố để ra VKS Thành phố ban hành lệnh giam.",
            "1. Việc xác định tình trạng yêu cầu \"Rất khẩn\" chưa phù hợp với tính chất của hỗ trợ khi kiểm tra (bị can có thời hạn tạm giam đến ngày 19/8/2026).\n2. Vụ án sau khi Phòng 2 tiếp nhận đã tự \"Thêm mới\" 01 trường hợp tạm giam khác. Ghi nhận để nhắc nhở tại Hội nghị toàn quốc."
        )

    def bind_events(self):
        self.bind("<Control-v>", lambda e: self.on_paste_clipboard())

    def on_save_key(self):
        key = self.ent_key.get().strip()
        self.cfg["gemini_api_key"] = key
        save_config(self.cfg)
        messagebox.showinfo("Thông báo", "Đã lưu Gemini API Key thành công!")

    def open_web_app(self):
        web_path = os.path.join(BASE_DIR, "index.html")
        if os.path.exists(web_path):
            webbrowser.open(f"file:///{web_path.replace(os.sep, '/')}")
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file index.html")

    def on_fix_existing_excel(self):
        target = filedialog.askopenfilename(
            title="Chọn file Excel cần làm đẹp và đóng băng tiêu đề",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if not target:
            return
        
        try:
            base, ext = os.path.splitext(target)
            out = f"{base}_Dep_DongBangTieuDe{ext}"
            clean_excel_file(target, out)
            format_excel_file(out)
            if messagebox.askyesno("Thành công", f"Đã định dạng thẩm mỹ, đóng băng tiêu đề và lưu tại:\n{os.path.basename(out)}\n\nBạn có muốn mở file ngay không?"):
                os.startfile(out)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xử lý file Excel: {e}")

    def add_record(self, filename, yeu_cau, bien_phap, file_path=""):
        idx = len(self.items_data) + 1
        yc_clean = unwrap_text(yeu_cau, is_numbered=False)
        bp_clean = unwrap_text(bien_phap, is_numbered=True)

        item = {
            "idx": idx,
            "file": filename,
            "path": file_path,
            "yeu_cau": yc_clean,
            "bien_phap": bp_clean
        }
        self.items_data.append(item)
        self.tree.insert("", "end", iid=str(idx), values=(idx, filename, yc_clean, bp_clean.replace("\n", " | ")))

    def on_select_row(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0]) - 1
        if 0 <= idx < len(self.items_data):
            rec = self.items_data[idx]
            self.txt_yeu_cau.delete("1.0", "end")
            self.txt_yeu_cau.insert("1.0", rec["yeu_cau"])
            self.txt_bien_phap.delete("1.0", "end")
            self.txt_bien_phap.insert("1.0", rec["bien_phap"])

    def on_save_row_edit(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Thông báo", "Vui lòng chọn 1 dòng trên bảng để chỉnh sửa.")
            return
        idx = int(selected[0]) - 1
        if 0 <= idx < len(self.items_data):
            new_yc = unwrap_text(self.txt_yeu_cau.get("1.0", "end").strip(), is_numbered=False)
            new_bp = unwrap_text(self.txt_bien_phap.get("1.0", "end").strip(), is_numbered=True)
            self.items_data[idx]["yeu_cau"] = new_yc
            self.items_data[idx]["bien_phap"] = new_bp
            self.tree.item(selected[0], values=(
                self.items_data[idx]["idx"],
                self.items_data[idx]["file"],
                new_yc,
                new_bp.replace("\n", " | ")
            ))
            messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu dòng!")

    def on_clear_list(self):
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa tất cả danh sách hiện tại?"):
            self.items_data.clear()
            for row in self.tree.get_children():
                self.tree.delete(row)
            self.txt_yeu_cau.delete("1.0", "end")
            self.txt_bien_phap.delete("1.0", "end")
            self.status_var.set("Đã xóa danh sách.")

    def on_add_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn file ảnh",
            filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp;*.bmp"), ("All files", "*.*")]
        )
        if files:
            self.process_files_async(list(files))

    def on_add_dir(self):
        dir_path = filedialog.askdirectory(title="Chọn thư mục chứa ảnh")
        if dir_path:
            valid_exts = (".jpg", ".jpeg", ".png", ".webp", ".bmp")
            files = [os.path.join(dir_path, f) for f in os.listdir(dir_path) if f.lower().endswith(valid_exts)]
            if files:
                self.process_files_async(files)
            else:
                messagebox.showinfo("Thông báo", "Không tìm thấy file ảnh phù hợp trong thư mục đã chọn.")

    def on_paste_clipboard(self):
        try:
            im = ImageGrab.grabclipboard()
            if isinstance(im, Image.Image):
                temp_dir = os.path.join(BASE_DIR, "scratch_images")
                os.makedirs(temp_dir, exist_ok=True)
                import time
                filename = f"paste_{int(time.time())}.png"
                temp_path = os.path.join(temp_dir, filename)
                im.save(temp_path, "PNG")
                self.process_files_async([temp_path])
            else:
                messagebox.showinfo("Thông báo", "Không có dữ liệu ảnh trong clipboard. Hãy chụp ảnh màn hình trước (Windows + Shift + S)!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lấy ảnh từ Clipboard: {e}")

    def on_start_process(self):
        self.on_add_files()

    def process_files_async(self, file_paths):
        t = threading.Thread(target=self._worker_process, args=(file_paths,), daemon=True)
        t.start()

    def _worker_process(self, file_paths):
        total = len(file_paths)
        api_key = self.ent_key.get().strip()

        self.btn_run.config(state="disabled")
        self.prog_bar["maximum"] = total
        self.prog_bar["value"] = 0

        for i, path in enumerate(file_paths, 1):
            fname = os.path.basename(path)
            self.status_var.set(f"Đang xử lý [{i}/{total}]: {fname}...")
            
            try:
                if api_key:
                    from batch_process import process_image_with_gemini
                    yeu_cau, bien_phap = process_image_with_gemini(path, api_key)
                else:
                    from batch_process import process_image_with_rapidocr
                    yeu_cau, bien_phap = process_image_with_rapidocr(path)
                
                self.after(0, self.add_record, fname, yeu_cau, bien_phap, path)
            except Exception as e:
                print(f"Lỗi {fname}:", e)
                try:
                    from batch_process import process_image_with_rapidocr
                    yeu_cau, bien_phap = process_image_with_rapidocr(path)
                    self.after(0, self.add_record, fname, yeu_cau, bien_phap, path)
                except Exception as ex2:
                    self.status_var.set(f"Lỗi: {e}")

            self.prog_bar["value"] = i

        self.status_var.set(f"Hoàn tất xử lý {total} ảnh!")
        self.btn_run.config(state="normal")

    def on_export_excel(self):
        if not self.items_data:
            messagebox.showwarning("Thông báo", "Chưa có dữ liệu để xuất Excel!")
            return

        save_path = filedialog.asksaveasfilename(
            title="Lưu file Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=os.path.basename(DEFAULT_EXCEL)
        )
        if not save_path:
            return

        try:
            from batch_process import export_to_excel
            export_to_excel(self.items_data, save_path)
            format_excel_file(save_path)
            if messagebox.askyesno("Thành công", f"Đã xuất file Excel đẹp & đóng băng tiêu đề thành công!\nBạn có muốn mở file ngay không?"):
                os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Excel: {e}")

if __name__ == "__main__":
    app = ToolApp()
    app.mainloop()
