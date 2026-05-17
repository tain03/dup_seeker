import os
import sys
import ctypes
import zipfile
import hashlib
import json
import threading
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from PIL import Image
import io
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
from tkinterdnd2 import DND_FILES, TkinterDnD
import imagehash
import xlsxwriter
import sqlite3

# --- THIẾT LẬP ID TIẾN TRÌNH CHO WINDOWS TASKBAR ICON ---
try:
    # Buộc Windows nhận diện tiến trình độc lập và hiển thị icon tùy chỉnh trên Taskbar thay vì logo Python mặc định
    myappid = "tain03.dupseeker.excelchecker.v2"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

def get_resource_path(relative_path):
    """Đường dẫn tuyệt đối cho tài nguyên khi chạy bình thường và sau khi đóng gói PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# --- ELITE PREMIUM DESIGN SYSTEM ---
COLORS = {
    "bg_canvas": "#F8FAFC",      
    "sidebar_primary": "#0F172A", 
    "sidebar_accent": "#1E293B",
    "primary": "#0284C7",        
    "primary_glow": "#38BDF8",
    "secondary": "#6366F1",      
    "success": "#10B981",        
    "danger": "#EF4444",         
    "text_dark": "#0F172A",
    "text_muted": "#64748B",
    "white": "#FFFFFF",
    "border": "#E2E8F0"
}

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

NS = {
    'sh': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'dr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

def get_cell_address(col, row):
    string = ""
    col += 1
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        string = chr(65 + remainder) + string
    return f"{string}{row + 1}"

# --- DATABASE TỐI ƯU HÓA CẤP ĐỘ BIG DATA ---
def init_db():
    conn = sqlite3.connect('dup_seeker_cache.db')
    c = conn.cursor()
    # HashCache: Bộ nhớ Ký ức vĩnh viễn (md5 -> phash)
    c.execute('''CREATE TABLE IF NOT EXISTS hash_cache
                 (md5 TEXT PRIMARY KEY, phash TEXT)''')
    # CurrentScan: Bảng tạm thay thế cho mảng trong RAM
    c.execute('''CREATE TABLE IF NOT EXISTS current_scan
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  full_path TEXT,
                  file_name TEXT,
                  sheet TEXT,
                  cell TEXT,
                  m_path TEXT,
                  md5 TEXT,
                  phash TEXT)''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_md5 ON current_scan(md5)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_phash ON current_scan(phash)')
    conn.commit()
    conn.close()

class ExcelProcessor:
    @staticmethod
    def get_sheet_names(file_paths):
        all_sheets = set()
        for path in file_paths:
            try:
                with zipfile.ZipFile(path, 'r') as z:
                    workbook_xml = z.read('xl/workbook.xml')
                    root = ET.fromstring(workbook_xml)
                    for s in root.findall('.//sh:sheet', NS):
                        all_sheets.add(s.get('name'))
            except: pass
        return sorted(list(all_sheets))

    @staticmethod
    def extract_metadata(file_path, ignore_sheets=None):
        """Pass 1: Trích xuất tọa độ ảnh và băm MD5 thô (Siêu nhanh, không AI)"""
        ignore_sheets = ignore_sheets or []
        results = []
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                sheet_names = {}
                workbook_xml = z.read('xl/workbook.xml')
                root = ET.fromstring(workbook_xml)
                for s in root.findall('.//sh:sheet', NS):
                    sheet_names[f"Sheet{s.get('sheetId')}"] = s.get('name')

                sheet_to_drawing = {}
                for f in z.namelist():
                    if f.startswith('xl/worksheets/_rels/sheet') and f.endswith('.xml.rels'):
                        s_id = f.replace('xl/worksheets/_rels/sheet', '').replace('.xml.rels', '')
                        root = ET.fromstring(z.read(f))
                        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if 'drawing' in rel.get('Type'):
                                target = rel.get('Target').replace('../', 'xl/')
                                sheet_to_drawing[target] = f"Sheet{s_id}"

                for drawing_path, s_id in sheet_to_drawing.items():
                    s_real_name = sheet_names.get(s_id, s_id)
                    if s_real_name in ignore_sheets: continue
                    if drawing_path not in z.namelist(): continue
                    
                    rel_path = f"xl/drawings/_rels/{Path(drawing_path).name}.rels"
                    drawing_to_media = {}
                    if rel_path in z.namelist():
                        rel_root = ET.fromstring(z.read(rel_path))
                        for rel in rel_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            drawing_to_media[rel.get('Id')] = rel.get('Target').replace('../', 'xl/')

                    d_root = ET.fromstring(z.read(drawing_path))
                    anchors = d_root.findall('.//dr:twoCellAnchor', NS) + \
                              d_root.findall('.//dr:oneCellAnchor', NS) + \
                              d_root.findall('.//dr:absoluteAnchor', NS)

                    for anchor in anchors:
                        cell_addr = "N/A"
                        from_tag = anchor.find('dr:from', NS)
                        if from_tag is not None:
                            col_tag = from_tag.find('dr:col', NS)
                            row_tag = from_tag.find('dr:row', NS)
                            if col_tag is not None and row_tag is not None:
                                cell_addr = get_cell_address(int(col_tag.text), int(row_tag.text))
                        elif anchor.tag.endswith('absoluteAnchor'):
                            pos = anchor.find('dr:pos', NS)
                            cell_addr = f"Float({pos.get('x')},{pos.get('y')})"

                        blip = anchor.find('.//a:blip', NS)
                        if blip is not None:
                            r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            m_path = drawing_to_media.get(r_id)
                            if m_path in z.namelist():
                                data = z.read(m_path)
                                md5_hash = hashlib.md5(data).hexdigest()
                                
                                results.append({
                                    "file": Path(file_path).name,
                                    "full_path": str(file_path),
                                    "sheet": s_real_name,
                                    "cell": cell_addr,
                                    "m_path": m_path,
                                    "md5": md5_hash
                                })
        except Exception:
            pass
        return results

    @staticmethod
    def compute_phash_for_unique(task):
        """Pass 2: AI Vision (pHash) chạy độc quyền cho các MD5 CHƯA TỪNG XUẤT HIỆN"""
        md5_val, full_path, m_path = task
        try:
            with zipfile.ZipFile(full_path, 'r') as z:
                data = z.read(m_path)
            pil_img = Image.open(io.BytesIO(data))
            return md5_val, str(imagehash.phash(pil_img))
        except Exception:
            return md5_val, None


class DuplicateApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("DUP-SEEKER - Elite Integrity Suite")
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(1200, int(screen_width * 0.8))
        height = min(750, int(screen_height * 0.8))
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(1300, 750)
        

        
        try:
            icon_path = get_resource_path("assets/app.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except:
            pass
        
        self.configure(fg_color=COLORS["bg_canvas"])

        self.selected_files = []
        self.ignore_sheets_str = tk.StringVar(value="Sheet1, Cosmetic, Critical Part")
        self.is_scanning = False
        self.duplicates_cache = []
        self.node_data = {}
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._draw_sidebar()
        self._draw_main()
        
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.handle_drop)

        # Phóng to cửa sổ toàn màn hình sau khi luồng xử lý giao diện (DPI scaling) của CustomTkinter hoàn thành
        self.after(150, lambda: self.state("zoomed"))

    def _draw_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=300, corner_radius=0, fg_color=COLORS["sidebar_primary"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.pack_propagate(False)
        
        branding = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        branding.pack(pady=(30, 30), padx=25, fill="x")
        
        ctk.CTkLabel(branding, text="DUP-", font=ctk.CTkFont(size=24, weight="bold"), text_color=COLORS["white"]).pack(side="left")
        ctk.CTkLabel(branding, text="SEEKER", font=ctk.CTkFont(size=24, weight="bold"), text_color=COLORS["primary"]).pack(side="left")

        self.btn_guide = ctk.CTkButton(self.sidebar, text="HELP / HƯỚNG DẪN 💡", fg_color="transparent", hover_color=COLORS["sidebar_accent"],
                                       border_width=1, border_color=COLORS["sidebar_accent"], font=ctk.CTkFont(size=12, weight="bold"),
                                       height=40, corner_radius=10, command=self.show_guide)
        self.btn_guide.pack(pady=(0, 30), padx=30, fill="x")

        group_lbl = ctk.CTkLabel(self.sidebar, text="MAIN ACTIONS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"])
        group_lbl.pack(anchor="w", padx=35, pady=(0, 10))

        self.btn_select = ctk.CTkButton(self.sidebar, text="SELECT DATA SOURCE", fg_color=COLORS["sidebar_accent"], hover_color=COLORS["primary"],
                                       border_width=1, border_color=COLORS["primary"], font=ctk.CTkFont(weight="bold"),
                                       height=50, corner_radius=12, command=self.select_files)
        self.btn_select.pack(pady=5, padx=30, fill="x")

        self.file_count_lbl = ctk.CTkLabel(self.sidebar, text="System standby", text_color=COLORS["text_muted"], font=ctk.CTkFont(size=12))
        self.file_count_lbl.pack(pady=5)

        config_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        config_box.pack(pady=25, padx=25, fill="x")
        
        ctk.CTkLabel(config_box, text="EXCLUSION SETS", font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", pady=(0,8))
        
        entry_f = ctk.CTkFrame(config_box, fg_color="transparent")
        entry_f.pack(fill="x")
        
        self.entry_ignore = ctk.CTkEntry(entry_f, textvariable=self.ignore_sheets_str, fg_color=COLORS["sidebar_accent"], border_color=COLORS["sidebar_accent"],
                                        height=35, corner_radius=10, text_color=COLORS["white"])
        self.entry_ignore.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.btn_browse_sheets = ctk.CTkButton(entry_f, text="🔍", width=35, height=35, fg_color=COLORS["sidebar_accent"], hover_color=COLORS["primary"],
                                              command=self.show_sheet_selector)
        self.btn_browse_sheets.pack(side="right")

        self.btn_export = ctk.CTkButton(self.sidebar, text="EXPORT REPORT (XLSX)", command=self.export_report, height=45, fg_color=COLORS["secondary"], hover_color="#4F46E5",
                                       font=ctk.CTkFont(size=13, weight="bold"), corner_radius=12, state="disabled")
        self.btn_export.pack(pady=(20, 0), padx=30, fill="x")

        self.btn_scan = ctk.CTkButton(self.sidebar, text="EXECUTE AUDIT", command=self.start_scan, height=50, fg_color=COLORS["success"], hover_color="#047857",
                                     font=ctk.CTkFont(size=15, weight="bold"), corner_radius=12)
        self.btn_scan.pack(side="bottom", pady=30, padx=25, fill="x")

        self.progress = ctk.CTkProgressBar(self.sidebar, progress_color=COLORS["primary"], height=8, corner_radius=4)
        self.progress.pack(side="bottom", pady=(0, 20), padx=30, fill="x")
        self.progress.set(0)
        
        ctk.CTkLabel(self.sidebar, text="Developed by ductai.nguyen", font=ctk.CTkFont(size=10), text_color=COLORS["text_muted"]).pack(side="bottom", pady=(0, 10))

    def _draw_main(self):
        self.main_view = ctk.CTkFrame(self, fg_color="transparent")
        self.main_view.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_view.grid_columnconfigure(0, weight=1)
        self.main_view.grid_rowconfigure(1, weight=1)

        # Dashboard Header
        metrics_grid = ctk.CTkFrame(self.main_view, fg_color="transparent", height=90)
        metrics_grid.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        metrics_grid.grid_propagate(False)
        metrics_grid.grid_columnconfigure((0,1,2), weight=1)

        self.m_files = self._metric_card(metrics_grid, "0", "TOTAL RECORDS", 0)
        self.m_dups = self._metric_card(metrics_grid, "0", "INTEGRITY ISSUES", 1, COLORS["danger"])
        self.m_total = self._metric_card(metrics_grid, "0", "PARSED OBJECTS", 2)

        # 2 Panels Layout
        self.content_frame = ctk.CTkFrame(self.main_view, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=4)
        self.content_frame.grid_columnconfigure(1, weight=3)
        self.content_frame.grid_rowconfigure(0, weight=1)

        # Left Panel (Treeview)
        self.list_frame = ctk.CTkFrame(self.content_frame, fg_color=COLORS["white"], corner_radius=10, border_width=1, border_color=COLORS["border"])
        self.list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        self.tree_scroll_y = ttk.Scrollbar(self.list_frame, orient="vertical")
        self.tree_scroll_y.pack(side="right", fill="y", pady=10, padx=(0, 5))
        
        self.tree_scroll_x = ttk.Scrollbar(self.list_frame, orient="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x", padx=10, pady=(0, 5))
        
        style = ttk.Style()
        style.theme_use("alt")
        
        # Thiết kế style cho bảng Treeview mang hơi hướng SaaS hiện đại có viền lưới đầy đủ
        style.configure("Treeview", 
                        background=COLORS["white"], 
                        foreground="#334155", 
                        rowheight=38, 
                        fieldbackground=COLORS["white"], 
                        borderwidth=1, 
                        cellborderwidth=1,
                        gridcolor="#E2E8F0",  # Vẽ các đường kẻ lưới (gridlines) phân tách cột và hàng sắc nét
                        font=("Segoe UI", 10))
                        
        style.configure("Treeview.Heading", 
                        font=("Segoe UI", 10, "bold"), 
                        background="#F1F5F9", 
                        foreground="#0F172A", 
                        bordercolor="#CBD5E1",  # Viền cột tiêu đề sẫm hơn để tạo điểm nhấn
                        relief="solid", 
                        borderwidth=1)
                        
        style.map("Treeview.Heading", background=[('active', '#E2E8F0')])
        
        # Lựa chọn màu Highlight nhẹ nhàng thanh lịch (Soft Blue) khi click chọn dòng
        style.map('Treeview', 
                  background=[('selected', '#E0F2FE')], 
                  foreground=[('selected', '#0369A1')])
        
        # Làm đẹp cho thanh cuộn (Scrollbar) đồng bộ với giao diện
        style.configure("Vertical.TScrollbar", 
                        background="#E2E8F0", 
                        troughcolor=COLORS["white"], 
                        bordercolor="transparent", 
                        arrowcolor="#64748B", 
                        relief="flat", 
                        borderwidth=0)
                        
        style.configure("Horizontal.TScrollbar", 
                        background="#E2E8F0", 
                        troughcolor=COLORS["white"], 
                        bordercolor="transparent", 
                        arrowcolor="#64748B", 
                        relief="flat", 
                        borderwidth=0)
        
        self.tree = ttk.Treeview(self.list_frame, columns=("Sheet", "Cell", "Matches"), show="tree headings", 
                                 yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        
        # Cấu hình các thẻ tag phân biệt dòng cha, dòng con chẵn và lẻ để tạo hiệu ứng Zebra Stripe (kẻ sọc bảng tính)
        self.tree.tag_configure("group_row", background="#EBF5FF", font=("Segoe UI", 11, "bold"), foreground="#1E3A8A") # Dòng cha xanh nhạt nổi bật
        self.tree.tag_configure("even_row", background="#F8FAFC", foreground="#334155")                              # Dòng con chẵn xám nhẹ
        self.tree.tag_configure("odd_row", background="#FFFFFF", foreground="#334155")                               # Dòng con lẻ trắng tinh
        
        self.tree.heading("#0", text="Duplicate Group / Files")
        self.tree.heading("Sheet", text="Sheet")
        self.tree.heading("Cell", text="Cell Address")
        self.tree.heading("Matches", text="Occurrences")
        
        # We start with minimum widths; they will be dynamically auto-fitted in render_results
        self.tree.column("#0", width=300, minwidth=200, stretch=True)
        self.tree.column("Sheet", width=150, minwidth=100, stretch=True, anchor="w")
        self.tree.column("Cell", width=100, minwidth=80, stretch=True, anchor="center")
        self.tree.column("Matches", width=120, minwidth=100, stretch=True, anchor="center")
        
        self.tree.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        self.tree_scroll_y.configure(command=self.tree.yview)
        self.tree_scroll_x.configure(command=self.tree.xview)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Right Panel (Preview)
        self.preview_frame = ctk.CTkFrame(self.content_frame, fg_color=COLORS["white"], corner_radius=10, border_width=1, border_color=COLORS["border"])
        self.preview_frame.grid(row=0, column=1, sticky="nsew")
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)
        
        self.preview_img_lbl = ctk.CTkLabel(self.preview_frame, text="Select a group to preview", text_color=COLORS["text_muted"], font=ctk.CTkFont(size=14, slant="italic"))
        self.preview_img_lbl.grid(row=0, column=0, pady=20)
        
        self.preview_details_scroll = ctk.CTkScrollableFrame(self.preview_frame, fg_color="transparent")
        self.preview_details_scroll.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

    def _metric_card(self, parent, val, label, col, color=None):
        card = ctk.CTkFrame(parent, fg_color=COLORS["white"], corner_radius=16, border_width=1, border_color=COLORS["border"])
        card.grid(row=0, column=col, sticky="nsew", padx=10)
        v_lbl = ctk.CTkLabel(card, text=val, font=ctk.CTkFont(size=28, weight="bold"), text_color=color or COLORS["text_dark"])
        v_lbl.pack(pady=(15, 0))
        ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"]).pack(pady=(0,15))
        return v_lbl

    def select_files(self):
        files = filedialog.askopenfilenames(title="System Selection", filetypes=[("Excel files", "*.xlsx")])
        if files:
            self.selected_files = list(files)
            self.file_count_lbl.configure(text=f"{len(self.selected_files)} resources active", text_color=COLORS["primary"])
            self.m_files.configure(text=str(len(self.selected_files)))

    def handle_drop(self, event):
        import re
        matches = [f[0] if f[0] else f[1] for f in re.findall(r'\{(.*?)\}|(\S+)', event.data)]
        xlsx_files = [f for f in matches if f.lower().endswith('.xlsx')]
        if xlsx_files:
            self.selected_files = list(set(self.selected_files + xlsx_files))
            self.file_count_lbl.configure(text=f"{len(self.selected_files)} resources active", text_color=COLORS["primary"])
            self.m_files.configure(text=str(len(self.selected_files)))

    def start_scan(self):
        if self.is_scanning or not self.selected_files: return
        self.is_scanning = True
        self.btn_scan.configure(state="disabled", text="INITIALIZING SQLITE...")
        self.btn_export.configure(state="disabled")
        threading.Thread(target=self.scan_logic, daemon=True).start()

    def scan_logic(self):
        try:
            self.tree.delete(*self.tree.get_children())
            self.preview_img_lbl.configure(image="", text="Processing Data...")
            for child in self.preview_details_scroll.winfo_children(): child.destroy()
            
            ignore_list = [s.strip() for s in self.ignore_sheets_str.get().split(",") if s.strip()]
            
            # Khởi tạo lại bảng tạm cho lượt quét mới
            conn = sqlite3.connect('dup_seeker_cache.db')
            c = conn.cursor()
            c.execute("DELETE FROM current_scan")
            conn.commit()
            
            # --- PASS 1: EXTRACT METADATA ---
            self.after(0, lambda: self.btn_scan.configure(text="[1/3] INDEXING..."))
            processed_files = 0
            
            # Chunking Data: Đọc luồng dữ liệu song song và GHI TRỰC TIẾP VÀO Ổ CỨNG thay vì RAM
            with ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count() - 1)) as executor:
                future_to_file = {executor.submit(ExcelProcessor.extract_metadata, f, ignore_list): f for f in self.selected_files}
                for future in as_completed(future_to_file):
                    metadata_list = future.result()
                    if metadata_list:
                        c.executemany('''INSERT INTO current_scan (full_path, file_name, sheet, cell, m_path, md5)
                                         VALUES (:full_path, :file, :sheet, :cell, :m_path, :md5)''', metadata_list)
                        conn.commit()
                    processed_files += 1
                    self.after(0, lambda p=processed_files: self.progress.set(p / len(self.selected_files) * 0.4))

            # Tìm những hình ảnh (MD5) MỚI HOÀN TOÀN chưa từng được phân tích AI
            c.execute('''SELECT DISTINCT current_scan.md5, current_scan.full_path, current_scan.m_path 
                         FROM current_scan 
                         LEFT JOIN hash_cache ON current_scan.md5 = hash_cache.md5 
                         WHERE hash_cache.md5 IS NULL''')
            uncached_tasks = c.fetchall()
            
            # --- PASS 2: COMPUTE AI VISION HASH ---
            if uncached_tasks:
                self.after(0, lambda: self.btn_scan.configure(text=f"[2/3] AI VISION ({len(uncached_tasks)})..."))
                processed_hashes = 0
                
                with ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count() - 1)) as executor:
                    future_to_hash = {executor.submit(ExcelProcessor.compute_phash_for_unique, t): t for t in uncached_tasks}
                    for future in as_completed(future_to_hash):
                        md5_val, phash_val = future.result()
                        if phash_val:
                            c.execute("INSERT OR IGNORE INTO hash_cache (md5, phash) VALUES (?, ?)", (md5_val, phash_val))
                        processed_hashes += 1
                        if processed_hashes % 50 == 0:
                            conn.commit() # Lưu định kỳ để bảo vệ dữ liệu nếu crash
                        self.after(0, lambda p=processed_hashes: self.progress.set(0.4 + (p / len(uncached_tasks) * 0.5)))
                conn.commit()

            # --- GROUPING ---
            self.after(0, lambda: self.btn_scan.configure(text="[3/3] AGGREGATING..."))
            self.after(0, lambda: self.progress.set(0.95))
            
            # Gắn p_hash cho mọi bức ảnh trong phiên quét hiện tại
            c.execute('''UPDATE current_scan 
                         SET phash = (SELECT phash FROM hash_cache WHERE hash_cache.md5 = current_scan.md5)''')
            conn.commit()
            
            # Tìm các mã p_hash có số lần xuất hiện > 1
            c.execute('''SELECT phash FROM current_scan 
                         WHERE phash IS NOT NULL 
                         GROUP BY phash HAVING COUNT(id) > 1''')
            duplicate_phashes = [row[0] for row in c.fetchall()]
            
            duplicates = []
            for p_hash in duplicate_phashes:
                c.execute('''SELECT file_name, full_path, sheet, cell, m_path 
                             FROM current_scan WHERE phash = ?''', (p_hash,))
                locs = []
                for row in c.fetchall():
                    locs.append({
                        'file': row[0],
                        'full_path': row[1],
                        'sheet': row[2],
                        'cell': row[3],
                        'm_path': row[4]
                    })
                duplicates.append(locs)
                
            c.execute("SELECT COUNT(id) FROM current_scan")
            total_images = c.fetchone()[0]
            conn.close()
            
            self.duplicates_cache = duplicates
            self.after(0, lambda: self.render_results(duplicates, total_images))
        except Exception as e:
            self.after(0, lambda e=e: messagebox.showerror("System Error", str(e)))
            self.after(0, self._reset_state)

    def render_results(self, duplicates, total_count):
        self.m_dups.configure(text=str(len(duplicates)))
        self.m_total.configure(text=str(total_count))
        
        self.tree.delete(*self.tree.get_children())
        self.node_data = {}
        
        if not duplicates:
            self.btn_export.configure(state="disabled")
            self.preview_img_lbl.configure(text="SYSTEM INTEGRITY VERIFIED - NO DUPLICATES", text_color=COLORS["success"])
        else:
            self.btn_export.configure(state="normal")
            self.preview_img_lbl.configure(text="Select a group or item on the left to preview", text_color=COLORS["text_muted"])
            
            # Khởi tạo độ dài cột mặc định bằng tiêu đề
            max_tree_len = len("Duplicate Group / Files")
            max_sheet_len = len("Sheet")
            max_cell_len = len("Cell Address")
            max_matches_len = len("Occurrences")
            
            for i, group in enumerate(duplicates):
                parent_iid = f"set_{i}"
                self.tree.insert("", "end", iid=parent_iid, text=f"SET #{i+1}", values=("", "", f"{len(group)} matches"), tags=("group_row",))
                self.node_data[parent_iid] = {"type": "group", "group": group, "index": i}
                
                max_tree_len = max(max_tree_len, len(f"SET #{i+1}"))
                max_matches_len = max(max_matches_len, len(f"{len(group)} matches"))
                
                for j, loc in enumerate(group):
                    child_iid = f"child_{i}_{j}"
                    row_tag = "even_row" if j % 2 == 0 else "odd_row"
                    self.tree.insert(parent_iid, "end", iid=child_iid, text=loc['file'], values=(loc['sheet'], loc['cell'], ""), tags=(row_tag,))
                    self.node_data[child_iid] = {"type": "loc", "loc": loc}
                    
                    max_tree_len = max(max_tree_len, len(loc['file']) + 4)
                    max_sheet_len = max(max_sheet_len, len(loc['sheet']))
                    max_cell_len = max(max_cell_len, len(loc['cell']))
            
            # Giới hạn độ rộng tối đa tự động (Cap width) để tránh việc file có tên siêu dài đẩy các cột khác ra rìa màn hình
            tree_width = min(max(max_tree_len * 8 + 40, 250), 380)
            sheet_width = min(max(max_sheet_len * 8 + 25, 120), 200)
            cell_width = min(max(max_cell_len * 8 + 25, 100), 130)
            matches_width = min(max(max_matches_len * 8 + 25, 110), 150)
            
            # Cấu hình các cột co giãn thông minh, bảo vệ không gian hiển thị của các cột thông tin phụ
            self.tree.column("#0", width=tree_width, minwidth=180, stretch=True)
            self.tree.column("Sheet", width=sheet_width, minwidth=100, stretch=True, anchor="w")
            self.tree.column("Cell", width=cell_width, minwidth=80, anchor="center", stretch=True)
            self.tree.column("Matches", width=matches_width, minwidth=100, anchor="center", stretch=True)
                
        self._reset_state()

    def on_tree_select(self, event):
        selected_item = self.tree.selection()
        if not selected_item: return
        
        iid = selected_item[0]
        data = self.node_data.get(iid)
        if not data: return
        
        # 1. Tải ảnh Preview
        loc_for_preview = None
        if data["type"] == "group":
            loc_for_preview = data["group"][0]
        else:
            loc_for_preview = data["loc"]
            
        try:
            with zipfile.ZipFile(loc_for_preview['full_path'], 'r') as z:
                img_data = z.read(loc_for_preview['m_path'])
            pil_img = Image.open(io.BytesIO(img_data))
            pil_img.thumbnail((320, 320))
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
            self.preview_img_lbl.configure(image=ctk_img, text="")
        except Exception:
            self.preview_img_lbl.configure(image="", text="Error loading image")
            
        # 2. Cập nhật thông tin chi tiết
        for child in self.preview_details_scroll.winfo_children(): child.destroy()
        
        if data["type"] == "group":
            group = data["group"]
            
            # Khung tổng quan của nhóm trùng lặp
            group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
            group_frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            ctk.CTkLabel(group_frame, text=f"SET #{data['index'] + 1} OVERVIEW", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
            ctk.CTkLabel(group_frame, text=f"Total Occurrences: {len(group)} duplicates", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["primary"]).pack(anchor="w", padx=15, pady=(0, 12))
            
            # Thiết lập bảng viền lưới (Spreadsheet Grid)
            grid_container = ctk.CTkFrame(group_frame, fg_color=COLORS["border"], corner_radius=8) # Background làm màu viền lưới
            grid_container.pack(fill="x", padx=15, pady=(0, 15))
            
            # Cấu hình tỉ lệ co giãn cột cho bảng lưới
            grid_container.grid_columnconfigure(0, weight=3) # File name
            grid_container.grid_columnconfigure(1, weight=2) # Sheet
            grid_container.grid_columnconfigure(2, weight=1) # Cell
            grid_container.grid_columnconfigure(3, weight=1) # Action
            
            # Tạo Tiêu đề Cột của bảng lưới (Header Row)
            headers = ["File Name", "Sheet", "Cell", "Action"]
            for col_idx, h_text in enumerate(headers):
                h_cell = ctk.CTkFrame(grid_container, fg_color="#F1F5F9", corner_radius=0)
                h_cell.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                ctk.CTkLabel(h_cell, text=h_text, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_dark"]).pack(padx=5, pady=6)
                
            # Đổ dữ liệu các file trùng lặp vào lưới (Data Rows)
            for row_idx, loc in enumerate(group):
                r = row_idx + 1
                bg_color = "#F8FAFC" if row_idx % 2 == 0 else "#FFFFFF"
                
                # Cột 0: Tên File
                c0_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                c0_frame.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
                ctk.CTkLabel(c0_frame, text=loc['file'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=130, justify="left").pack(anchor="w", padx=6, pady=6)
                
                # Cột 1: Sheet
                c1_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                c1_frame.grid(row=r, column=1, sticky="nsew", padx=1, pady=1)
                ctk.CTkLabel(c1_frame, text=loc['sheet'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=80).pack(padx=5, pady=6)
                
                # Cột 2: Ô Cell
                c2_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                c2_frame.grid(row=r, column=2, sticky="nsew", padx=1, pady=1)
                ctk.CTkLabel(c2_frame, text=loc['cell'], font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["danger"]).pack(padx=5, pady=6)
                
                # Cột 3: Nút bấm mở file trực tiếp
                c3_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                c3_frame.grid(row=r, column=3, sticky="nsew", padx=1, pady=1)
                btn = ctk.CTkButton(c3_frame, text="OPEN", width=50, height=22, fg_color=COLORS["secondary"], hover_color=COLORS["primary"], corner_radius=4,
                                   font=ctk.CTkFont(size=9, weight="bold"), command=lambda p=loc['full_path'], s=loc['sheet'], c=loc['cell']: self.open_excel_at_location(p, s, c))
                btn.pack(padx=5, pady=4)
        else:
            # Giao diện khi chọn một Vị trí trùng cụ thể (Child Node)
            loc = data["loc"]
            loc_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
            loc_frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            ctk.CTkLabel(loc_frame, text="OCCURRENCE DETAILS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 10))
            
            # File
            f_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
            f_box.pack(fill="x", padx=15, pady=4)
            ctk.CTkLabel(f_box, text="File Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
            ctk.CTkLabel(f_box, text=loc['file'], font=ctk.CTkFont(size=12), text_color=COLORS["text_dark"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
            
            # Sheet
            s_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
            s_box.pack(fill="x", padx=15, pady=4)
            ctk.CTkLabel(s_box, text="Sheet:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
            ctk.CTkLabel(s_box, text=loc['sheet'], font=ctk.CTkFont(size=12), text_color=COLORS["primary"], anchor="w").pack(side="left")
            
            # Cell
            c_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
            c_box.pack(fill="x", padx=15, pady=4)
            ctk.CTkLabel(c_box, text="Cell Address:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
            ctk.CTkLabel(c_box, text=loc['cell'], font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["danger"], anchor="w").pack(side="left")
            
            # Path
            p_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
            p_box.pack(fill="x", padx=15, pady=4)
            ctk.CTkLabel(p_box, text="Full Path:", font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"], width=80, anchor="w").pack(side="left")
            ctk.CTkLabel(p_box, text=loc['full_path'], font=ctk.CTkFont(size=9), text_color=COLORS["text_muted"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
            
            # Nút bấm mở file to rõ ràng, không lo bị đè/mất text
            btn = ctk.CTkButton(loc_frame, text="📂 OPEN EXCEL FILE", height=45, fg_color=COLORS["primary"], hover_color=COLORS["primary_glow"], corner_radius=10,
                               font=ctk.CTkFont(size=13, weight="bold"), 
                               command=lambda p=loc['full_path'], s=loc['sheet'], c=loc['cell']: self.open_excel_at_location(p, s, c))
            btn.pack(fill="x", padx=15, pady=(20, 15))

    def open_excel_at_location(self, file_path, sheet_name, cell_address):
        """Mở tệp Excel và tự động chọn đúng Sheet + Cell bằng win32com (COM Automation).
           Nếu máy không hỗ trợ win32com, hệ thống sẽ tự động hạ cấp xuống os.startfile (mở file cơ bản)."""
        try:
            import win32com.client
            abs_path = os.path.abspath(file_path)
            
            # Kết nối tới tiến trình Excel đang chạy hoặc khởi chạy mới
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                excel = win32com.client.Dispatch("Excel.Application")
                
            excel.Visible = True
            
            # Tìm xem Workbook đã được mở từ trước chưa
            wb = None
            try:
                for opened_wb in excel.Workbooks:
                    if os.path.abspath(opened_wb.FullName).lower() == abs_path.lower():
                        wb = opened_wb
                        break
            except Exception:
                pass
                
            if wb is None:
                wb = excel.Workbooks.Open(abs_path)
                
            # Kích hoạt đúng Sheet chứa ảnh trùng
            try:
                ws = wb.Sheets(sheet_name)
                ws.Activate()
            except Exception:
                pass
                
            # Chọn (Focus) vào đúng ô Cell chứa ảnh
            try:
                if cell_address and cell_address != "N/A" and not cell_address.startswith("Float"):
                    ws.Range(cell_address).Select()
            except Exception:
                pass
                
            # Đưa cửa sổ Active lên trước
            try:
                excel.ActiveWindow.Activate()
            except Exception:
                pass
        except Exception:
            # Fallback về mở file mặc định nếu có bất cứ lỗi gì xảy ra (ví dụ: máy không cài Excel hoặc thiếu thư viện)
            try:
                os.startfile(file_path)
            except Exception:
                pass

    def export_report(self):
        if not self.duplicates_cache: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save Audit Report")
        if not file_path: return

        try:
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet("Audit Results")
            
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0F172A', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            group_fmt = workbook.add_format({'bold': True, 'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1})

            headers = ["PREVIEW", "GROUP ID", "FILE NAME", "SHEET NAME", "CELL ADDRESS"]
            for col, text in enumerate(headers):
                worksheet.write(0, col, text, header_fmt)
            
            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:E', 25)

            current_row = 1
            for i, group in enumerate(self.duplicates_cache):
                group_id = f"SET #{i+1}"
                
                with zipfile.ZipFile(group[0]['full_path'], 'r') as z:
                    img_data = z.read(group[0]['m_path'])
                img_io = io.BytesIO(img_data)
                
                worksheet.set_row(current_row, 120)
                worksheet.insert_image(current_row, 0, f"img_{i}.png", {'image_data': img_io, 'x_scale': 0.2, 'y_scale': 0.2, 'x_offset': 5, 'y_offset': 5})
                
                for j, loc in enumerate(group):
                    worksheet.write(current_row, 1, group_id, group_fmt if j==0 else cell_fmt)
                    worksheet.write(current_row, 2, loc['file'], cell_fmt)
                    worksheet.write(current_row, 3, loc['sheet'], cell_fmt)
                    worksheet.write(current_row, 4, loc['cell'], cell_fmt)
                    current_row += 1

            workbook.close()
            messagebox.showinfo("Success", f"Report exported successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export report: {e}")

    def _reset_state(self):
        self.is_scanning = False
        self.btn_scan.configure(state="normal", text="EXECUTE AUDIT")
        self.progress.set(0)

    def show_guide(self):
        guide_win = ctk.CTkToplevel(self)
        guide_win.title("DUP-SEEKER USER GUIDE")
        guide_win.geometry("600x550")
        guide_win.configure(fg_color=COLORS["white"])
        guide_win.after(100, lambda: guide_win.focus())
        
        scroll = ctk.CTkScrollableFrame(guide_win, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(scroll, text="🇻🇳 HƯỚNG DẪN SỬ DỤNG", font=ctk.CTkFont(size=18, weight="bold"), text_color=COLORS["primary"]).pack(anchor="w", pady=(0,10))
        vn_text = (
            "1. CHỌN DỮ LIỆU: Nhấn 'SELECT DATA SOURCE' hoặc kéo thả các file Excel (.xlsx) trực tiếp vào ứng dụng.\n\n"
            "2. LOẠI TRỪ SHEET: Nhấn biểu tượng 🔍 cạnh ô nhập liệu để hiện danh sách toàn bộ Sheet và tích chọn những Sheet mẫu cần bỏ qua.\n\n"
            "3. THỰC HIỆN: Nhấn 'EXECUTE AUDIT'. Hệ thống sử dụng AI Vision (pHash) kết hợp SQLite Database để quét hàng nghìn file siêu tốc.\n\n"
            "4. KẾT QUẢ: Xem danh sách ảnh lỗi. Nhấn 'VIEW FILE' để mở tệp Excel gốc hoặc 'EXPORT REPORT' để lưu báo cáo tổng hợp có kèm ảnh xem trước."
        )
        ctk.CTkLabel(scroll, text=vn_text, font=ctk.CTkFont(size=13), justify="left", wraplength=520, text_color=COLORS["text_dark"]).pack(anchor="w", pady=(0,30))

        ctk.CTkLabel(scroll, text="🇺🇸 USER GUIDE", font=ctk.CTkFont(size=18, weight="bold"), text_color=COLORS["secondary"]).pack(anchor="w", pady=(0,10))
        en_text = (
            "1. SELECT DATA: Click 'SELECT DATA SOURCE' or drag and drop Excel files (.xlsx) into the app.\n\n"
            "2. EXCLUDE SHEETS: Click the 🔍 icon to browse all sheet names and check the ones you want to skip.\n\n"
            "3. EXECUTE: Click 'EXECUTE AUDIT'. The system uses AI Vision with SQLite cache for unlimited scalability.\n\n"
            "4. RESULTS: View issues. Click 'VIEW FILE' to open source Excel or 'EXPORT REPORT' to save a summary report with image previews."
        )
        ctk.CTkLabel(scroll, text=en_text, font=ctk.CTkFont(size=13), justify="left", wraplength=520, text_color=COLORS["text_dark"]).pack(anchor="w")
        ctk.CTkButton(guide_win, text="CLOSE / ĐÓNG", command=guide_win.destroy, fg_color=COLORS["sidebar_primary"], corner_radius=10).pack(pady=20)

    def show_sheet_selector(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Vui lòng chọn file Excel trước / Please select Excel files first")
            return
            
        sheets = ExcelProcessor.get_sheet_names(self.selected_files)
        if not sheets: return

        selector = ctk.CTkToplevel(self)
        selector.title("SELECT SHEETS TO IGNORE")
        selector.geometry("400x500")
        selector.configure(fg_color=COLORS["white"])
        selector.after(100, lambda: selector.focus())
        
        ctk.CTkLabel(selector, text="TÍCH CHỌN SHEET BỎ QUA", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["primary"]).pack(pady=15)
        
        scroll = ctk.CTkScrollableFrame(selector, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)
        
        current_ignores = [s.strip() for s in self.ignore_sheets_str.get().split(",") if s.strip()]
        checkboxes = {}
        
        for s in sheets:
            var = tk.BooleanVar(value=s in current_ignores)
            cb = ctk.CTkCheckBox(scroll, text=s, variable=var, font=ctk.CTkFont(size=12), 
                                 fg_color=COLORS["primary"], hover_color=COLORS["primary_glow"])
            cb.pack(anchor="w", pady=5)
            checkboxes[s] = var
            
        def apply_selection():
            selected = [s for s, v in checkboxes.items() if v.get()]
            self.ignore_sheets_str.set(", ".join(selected))
            selector.destroy()
            
        ctk.CTkButton(selector, text="APPLY / XÁC NHẬN", command=apply_selection, fg_color=COLORS["primary"], corner_radius=10).pack(pady=20)

if __name__ == "__main__":
    # Rất quan trọng khi sử dụng ProcessPoolExecutor trong ứng dụng PyInstaller trên Windows
    multiprocessing.freeze_support()
    init_db()
    app = DuplicateApp()
    app.mainloop()
