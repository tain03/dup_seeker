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
import re

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
    c.execute('DROP TABLE IF EXISTS current_scan')
    c.execute('''CREATE TABLE current_scan
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  full_path TEXT,
                  file_name TEXT,
                  sheet TEXT,
                  cell TEXT,
                  m_path TEXT,
                  md5 TEXT,
                  phash TEXT,
                  pic_name TEXT)''')
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

                        # Trích xuất tên hình ảnh (pic_name)
                        pic_name = ""
                        pic = anchor.find('.//dr:pic', NS)
                        if pic is not None:
                            nvPicPr = pic.find('dr:nvPicPr', NS)
                            if nvPicPr is not None:
                                cNvPr = nvPicPr.find('dr:cNvPr', NS)
                                if cNvPr is not None:
                                    pic_name = cNvPr.get('name') or ""

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
                                    "md5": md5_hash,
                                    "pic_name": pic_name
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

    @staticmethod
    def col_to_num(col_str):
        num = 0
        for char in col_str:
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num

    @staticmethod
    def num_to_col(num):
        col_str = ""
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            col_str = chr(remainder + ord('A')) + col_str
        return col_str

    @staticmethod
    def cell_to_coord(cell_addr):
        match = re.match(r'^([A-Z]+)([0-9]+)$', cell_addr)
        if match:
            col_str, row_str = match.groups()
            return int(row_str), ExcelProcessor.col_to_num(col_str)
        return None

    @staticmethod
    def translate_formula(formula, row_offset, col_offset):
        if not formula:
            return ""
        
        def replace_match(match):
            col_abs, col_letter, row_abs, row_num_str = match.groups()
            
            if not row_abs:
                new_row = int(row_num_str) + row_offset
                row_str = str(max(1, new_row))
            else:
                row_str = row_num_str
                
            if not col_abs:
                new_col = ExcelProcessor.col_to_num(col_letter) + col_offset
                col_str = ExcelProcessor.num_to_col(max(1, new_col))
            else:
                col_str = col_letter
                
            return f"{col_abs}{col_str}{row_abs}{row_str}"
        
        pattern = r'(\$?)([A-Z]{1,3})(\$?)([0-9]+)\b(?!\s*\()'
        return re.sub(pattern, replace_match, formula)

    @staticmethod
    def extract_formula_integrity(file_path, ignore_sheets=None):
        """Quét các ô trong tệp Excel để tìm ô chứa giá trị cứng Pass/Fail và ghi lại các công thức"""
        ignore_sheets = ignore_sheets or []
        results = []
        formulas = {}
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                # 1. Đọc tên sheet thực tế
                sheet_names = {}
                workbook_xml = z.read('xl/workbook.xml')
                root = ET.fromstring(workbook_xml)
                sheet_info = []
                for s in root.findall('.//sh:sheet', NS):
                    s_name = s.get('name')
                    s_id = s.get('sheetId')
                    r_id = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    sheet_info.append((s_name, s_id, r_id))
                    sheet_names[f"Sheet{s_id}"] = s_name

                # 2. Đọc sharedStrings.xml (nếu có)
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    ss_xml = z.read('xl/sharedStrings.xml')
                    ss_root = ET.fromstring(ss_xml)
                    for si in ss_root.findall('.//sh:t', NS):
                        shared_strings.append(si.text or "")
                
                # 3. Đọc workbook relationships
                wb_rels = {}
                if 'xl/_rels/workbook.xml.rels' in z.namelist():
                    rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                    for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        wb_rels[rel.get('Id')] = rel.get('Target').replace('worksheets/', 'xl/worksheets/')

                # Duyệt từng sheet
                for s_name, s_id, r_id in sheet_info:
                    if s_name in ignore_sheets:
                        continue
                    
                    sheet_path = wb_rels.get(r_id, f"xl/worksheets/sheet{s_id}.xml")
                    if sheet_path not in z.namelist():
                        sheet_path = f"xl/{sheet_path}" if not sheet_path.startswith('xl/') else sheet_path
                        if sheet_path not in z.namelist():
                            continue
                    
                    sheet_xml = z.read(sheet_path)
                    s_root = ET.fromstring(sheet_xml)
                    
                    shared_formulas = {} # Lưu công thức mẫu theo si trong từng sheet
                    
                    for c_tag in s_root.findall('.//sh:c', NS):
                        cell_addr = c_tag.get('r')
                        cell_type = c_tag.get('t')
                        
                        f_tag = c_tag.find('sh:f', NS)
                        has_formula = f_tag is not None
                        if has_formula:
                            f_type = f_tag.get('t')
                            f_si = f_tag.get('si')
                            formula_text = f_tag.text or ""
                            
                            if f_type == 'shared' and f_si is not None:
                                if formula_text:
                                    shared_formulas[f_si] = (cell_addr, formula_text)
                                else:
                                    master_info = shared_formulas.get(f_si)
                                    if master_info:
                                        master_cell, master_formula = master_info
                                        m_coord = ExcelProcessor.cell_to_coord(master_cell)
                                        c_coord = ExcelProcessor.cell_to_coord(cell_addr)
                                        if m_coord and c_coord:
                                            row_off = c_coord[0] - m_coord[0]
                                            col_off = c_coord[1] - m_coord[1]
                                            formula_text = ExcelProcessor.translate_formula(master_formula, row_off, col_off)
                                        else:
                                            formula_text = master_formula
                            
                            # Chuẩn hóa công thức có dấu bằng ở đầu
                            if formula_text and not formula_text.startswith("="):
                                formula_text = f"={formula_text}"
                            formulas[f"{s_name}!{cell_addr}"] = formula_text
                            continue
                        
                        v_tag = c_tag.find('sh:v', NS)
                        val = ""
                        if v_tag is not None:
                            raw_val = v_tag.text or ""
                            if cell_type == 's':
                                try:
                                    idx = int(raw_val)
                                    if 0 <= idx < len(shared_strings):
                                        val = shared_strings[idx]
                                except:
                                    pass
                            elif cell_type == 'inlineStr' or cell_type == 'str':
                                val = raw_val
                            else:
                                val = raw_val
                        
                        is_tag = c_tag.find('.//sh:t', NS)
                        if is_tag is not None and not val:
                            val = is_tag.text or ""
                        
                        clean_val = val.strip().upper()
                        if clean_val in ["PASS", "FAIL"]:
                            results.append({
                                "file": Path(file_path).name,
                                "full_path": str(file_path),
                                "sheet": s_name,
                                "cell": cell_addr,
                                "value": val,
                                "type": "Hardcoded " + val.capitalize()
                            })
        except Exception as e:
            print(f"Error scanning formula in file {Path(file_path).name}: {e}")
        return {"violations": results, "formulas": formulas}


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

        self.audit_mode = ctk.CTkSegmentedButton(self.sidebar, values=["IMAGE AUDIT", "FORMULA AUDIT"],
                                                 command=self.on_mode_change, fg_color=COLORS["sidebar_accent"],
                                                 selected_color=COLORS["primary"], selected_hover_color="#0284C7",
                                                 text_color=COLORS["white"], height=40, corner_radius=10)
        self.audit_mode.pack(pady=(0, 15), padx=25, fill="x")
        self.audit_mode.set("IMAGE AUDIT")
        self.current_mode = "IMAGE AUDIT"

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

    def on_mode_change(self, value):
        self.current_mode = value
        self._reset_state()
        if value == "IMAGE AUDIT":
            self.tree.heading("#0", text="Duplicate Group / Files")
            self.tree.heading("Sheet", text="Sheet")
            self.tree.heading("Cell", text="Cell Address")
            self.tree.heading("Matches", text="Occurrences")
            self.preview_img_lbl.configure(text="Select a group to preview", image="")
            self.m_files.configure(text=str(len(self.selected_files)))
            self.m_dups.configure(text="0")
            self.m_total.configure(text="0")
        else:
            self.tree.heading("#0", text="File Name")
            self.tree.heading("Sheet", text="Sheet Name")
            self.tree.heading("Cell", text="Cell Address")
            self.tree.heading("Matches", text="Hardcoded Value")
            self.preview_img_lbl.configure(text="Select a manual Pass/Fail cell to inspect", image="")
            self.m_files.configure(text=str(len(self.selected_files)))
            self.m_dups.configure(text="0")
            self.m_total.configure(text="0")

    def start_scan(self):
        if self.is_scanning or not self.selected_files: return
        self.is_scanning = True
        self.btn_scan.configure(state="disabled", text="INDEXING FILES..." if self.current_mode == "FORMULA AUDIT" else "INITIALIZING SQLITE...")
        self.btn_export.configure(state="disabled")
        threading.Thread(target=self.scan_logic, daemon=True).start()

    def scan_logic(self):
        try:
            self.tree.delete(*self.tree.get_children())
            self.preview_img_lbl.configure(image="", text="Processing Data...")
            for child in self.preview_details_scroll.winfo_children(): child.destroy()
            
            ignore_list = [s.strip() for s in self.ignore_sheets_str.get().split(",") if s.strip()]
            
            if self.current_mode == "IMAGE AUDIT":
                conn = sqlite3.connect('dup_seeker_cache.db')
                c = conn.cursor()
                c.execute("DELETE FROM current_scan")
                conn.commit()
                
                self.after(0, lambda: self.btn_scan.configure(text="[1/3] INDEXING..."))
                processed_files = 0
                
                with ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count() - 1)) as executor:
                    future_to_file = {executor.submit(ExcelProcessor.extract_metadata, f, ignore_list): f for f in self.selected_files}
                    for future in as_completed(future_to_file):
                        metadata_list = future.result()
                        if metadata_list:
                            c.executemany('''INSERT INTO current_scan (full_path, file_name, sheet, cell, m_path, md5, pic_name)
                                             VALUES (:full_path, :file, :sheet, :cell, :m_path, :md5, :pic_name)''', metadata_list)
                            conn.commit()
                        processed_files += 1
                        self.after(0, lambda p=processed_files: self.progress.set(p / len(self.selected_files) * 0.4))

                c.execute('''SELECT DISTINCT current_scan.md5, current_scan.full_path, current_scan.m_path 
                             FROM current_scan 
                             LEFT JOIN hash_cache ON current_scan.md5 = hash_cache.md5 
                             WHERE hash_cache.md5 IS NULL''')
                uncached_tasks = c.fetchall()
                
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
                                conn.commit()
                            self.after(0, lambda p=processed_hashes: self.progress.set(0.4 + (p / len(uncached_tasks) * 0.5)))
                    conn.commit()

                self.after(0, lambda: self.btn_scan.configure(text="[3/3] AGGREGATING..."))
                self.after(0, lambda: self.progress.set(0.95))
                
                c.execute('''UPDATE current_scan 
                             SET phash = (SELECT phash FROM hash_cache WHERE hash_cache.md5 = current_scan.md5)''')
                conn.commit()
                
                c.execute('''SELECT phash FROM current_scan 
                             WHERE phash IS NOT NULL 
                             GROUP BY phash HAVING COUNT(id) > 1''')
                duplicate_phashes = [row[0] for row in c.fetchall()]
                
                duplicates = []
                for p_hash in duplicate_phashes:
                    c.execute('''SELECT file_name, full_path, sheet, cell, m_path, pic_name 
                                 FROM current_scan WHERE phash = ?''', (p_hash,))
                    locs = []
                    for row in c.fetchall():
                        locs.append({
                            'file': row[0],
                            'full_path': row[1],
                            'sheet': row[2],
                            'cell': row[3],
                            'm_path': row[4],
                            'pic_name': row[5]
                        })
                    duplicates.append(locs)
                    
                c.execute("SELECT COUNT(id) FROM current_scan")
                total_images = c.fetchone()[0]
                conn.close()
                
                self.duplicates_cache = duplicates
                self.after(0, lambda: self.render_results(duplicates, total_images))
            else:
                self.after(0, lambda: self.btn_scan.configure(text="[1/2] AUDITING FORMULAS..."))
                processed_files = 0
                violations_list = []
                file_formulas = {}
                file_violations = {}
                
                with ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count() - 1)) as executor:
                    future_to_file = {executor.submit(ExcelProcessor.extract_formula_integrity, f, ignore_list): f for f in self.selected_files}
                    for future in as_completed(future_to_file):
                        file_path = future_to_file[future]
                        res = future.result()
                        if res:
                            v_list = res.get("violations", [])
                            violations_list.extend(v_list)
                            file_formulas[str(file_path)] = res.get("formulas", {})
                            file_violations[str(file_path)] = {
                                f"{v['sheet']}!{v['cell']}": v['value'] for v in v_list
                            }
                        processed_files += 1
                        self.after(0, lambda p=processed_files: self.progress.set(p / len(self.selected_files) * 0.9))
                        
                self.after(0, lambda: self.btn_scan.configure(text="[2/2] COMPILING RESULTS..."))
                self.after(0, lambda: self.progress.set(0.95))
                
                # Tìm các tọa độ ô có công thức ở ít nhất một file để so sánh chéo
                all_formula_coords = set()
                for f_path, f_map in file_formulas.items():
                    all_formula_coords.update(f_map.keys())
                
                inconsistencies = []
                for coord in all_formula_coords:
                    coord_details = {}
                    distinct_states = set()
                    
                    for f_path in self.selected_files:
                        f_path_str = str(f_path)
                        f_name = Path(f_path).name
                        
                        f_map = file_formulas.get(f_path_str, {})
                        v_map = file_violations.get(f_path_str, {})
                        
                        if coord in f_map:
                            formula_text = f_map[coord]
                            coord_details[f_path_str] = {
                                "file": f_name,
                                "full_path": f_path_str,
                                "state": "formula",
                                "value": formula_text
                            }
                            distinct_states.add(f"formula:{formula_text}")
                        elif coord in v_map:
                            val = v_map[coord]
                            coord_details[f_path_str] = {
                                "file": f_name,
                                "full_path": f_path_str,
                                "state": "hardcoded",
                                "value": val
                            }
                            distinct_states.add("hardcoded")
                        else:
                            coord_details[f_path_str] = {
                                "file": f_name,
                                "full_path": f_path_str,
                                "state": "missing",
                                "value": "(No formula / Empty)"
                            }
                            distinct_states.add("missing")
                            
                    # Nếu có nhiều hơn 1 trạng thái công thức/giá trị tại cùng 1 vị trí thì là không đồng nhất
                    if len(distinct_states) > 1:
                        sheet_name, cell_addr = coord.split("!", 1)
                        inconsistencies.append({
                            "sheet": sheet_name,
                            "cell": cell_addr,
                            "coordinate": coord,
                            "files": coord_details
                        })
                
                combined_results = {
                    "hardcoded": violations_list,
                    "inconsistencies": inconsistencies
                }
                
                self.duplicates_cache = combined_results
                total_formulas_count = sum(len(f_map) for f_map in file_formulas.values())
                total_scanned_count = total_formulas_count + len(violations_list)
                
                self.after(0, lambda: self.render_results(combined_results, total_scanned_count))
        except Exception as e:
            self.after(0, lambda e=e: messagebox.showerror("System Error", str(e)))
            self.after(0, self._reset_state)

    def render_results(self, duplicates, total_count):
        self.tree.delete(*self.tree.get_children())
        self.node_data = {}
        
        if self.current_mode == "IMAGE AUDIT":
            self.m_dups.configure(text=str(len(duplicates)))
            self.m_total.configure(text=str(total_count))
            
            if not duplicates:
                self.btn_export.configure(state="disabled")
                self.preview_img_lbl.configure(text="SYSTEM INTEGRITY VERIFIED - NO DUPLICATES", text_color=COLORS["success"])
            else:
                self.btn_export.configure(state="normal")
                self.preview_img_lbl.configure(text="Select a group or item on the left to preview", text_color=COLORS["text_muted"])
                
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
                
                tree_width = min(max(max_tree_len * 8 + 40, 250), 380)
                sheet_width = min(max(max_sheet_len * 8 + 25, 120), 200)
                cell_width = min(max(max_cell_len * 8 + 25, 100), 130)
                matches_width = min(max(max_matches_len * 8 + 25, 110), 150)
                
                self.tree.column("#0", width=tree_width, minwidth=180, stretch=True)
                self.tree.column("Sheet", width=sheet_width, minwidth=100, stretch=True, anchor="w")
                self.tree.column("Cell", width=cell_width, minwidth=80, anchor="center", stretch=True)
                self.tree.column("Matches", width=matches_width, minwidth=100, anchor="center", stretch=True)
        else:
            hardcoded_violations = duplicates.get("hardcoded", [])
            inconsistencies = duplicates.get("inconsistencies", [])
            
            total_issues = len(hardcoded_violations) + len(inconsistencies)
            self.m_dups.configure(text=str(total_issues))
            self.m_total.configure(text=str(total_count))
            
            if not hardcoded_violations and not inconsistencies:
                self.btn_export.configure(state="disabled")
                self.preview_img_lbl.configure(text="SYSTEM INTEGRITY VERIFIED - NO ISSUES DETECTED", text_color=COLORS["success"])
            else:
                self.btn_export.configure(state="normal")
                self.preview_img_lbl.configure(text="Select an item on the left to inspect", text_color=COLORS["text_muted"])
                
                max_tree_len = len("🔄 CROSS-FILE FORMULA INCONSISTENCIES")
                max_sheet_len = len("Sheet Name")
                max_cell_len = len("Cell Address")
                max_matches_len = len("Formula / Value")
                
                # 1. Báo cáo ô trị tĩnh Pass/Fail
                if hardcoded_violations:
                    parent_hc_iid = "parent_hardcoded"
                    self.tree.insert("", "end", iid=parent_hc_iid, text=f"⚠️ HARDCODED PASS/FAIL ({len(hardcoded_violations)} cells)", values=("", "", "Inspect"), tags=("group_row",))
                    self.node_data[parent_hc_iid] = {"type": "hardcoded_root", "items": hardcoded_violations}
                    
                    from collections import defaultdict
                    by_file = defaultdict(list)
                    for item in hardcoded_violations:
                        by_file[item['file']].append(item)
                        
                    for f_idx, (f_name, items) in enumerate(by_file.items()):
                        file_parent_iid = f"file_{f_name.replace(' ', '_')}"
                        self.tree.insert(parent_hc_iid, "end", iid=file_parent_iid, text=f_name, values=("", "", f"{len(items)} cells"), tags=("group_row",))
                        self.node_data[file_parent_iid] = {"type": "formula_file", "file": f_name, "items": items}
                        
                        for j, item in enumerate(items):
                            child_iid = f"cell_{f_name.replace(' ', '_')}_{j}"
                            row_tag = "even_row" if j % 2 == 0 else "odd_row"
                            self.tree.insert(file_parent_iid, "end", iid=child_iid, text=f"Cell: {item['cell']} [{item['value']}]",
                                             values=(item['sheet'], item['cell'], item['value']), tags=(row_tag,))
                            self.node_data[child_iid] = {"type": "formula_cell", "item": item}
                            
                            max_sheet_len = max(max_sheet_len, len(item['sheet']))
                            max_cell_len = max(max_cell_len, len(item['cell']))
                            max_matches_len = max(max_matches_len, len(item['value']))
                
                # 2. Báo cáo không đồng nhất công thức
                if inconsistencies:
                    parent_inc_iid = "parent_inconsistency"
                    self.tree.insert("", "end", iid=parent_inc_iid, text=f"🔄 FORMULA INCONSISTENCY ({len(inconsistencies)} coords)", values=("", "", "Inspect"), tags=("group_row",))
                    self.node_data[parent_inc_iid] = {"type": "inconsistency_root", "items": inconsistencies}
                    
                    for idx, inc in enumerate(inconsistencies):
                        coord_parent_iid = f"coord_{idx}"
                        self.tree.insert(parent_inc_iid, "end", iid=coord_parent_iid, text=f"Coord: {inc['coordinate']}",
                                         values=(inc['sheet'], inc['cell'], f"{len(inc['files'])} files"), tags=("group_row",))
                        self.node_data[coord_parent_iid] = {"type": "inconsistency_group", "item": inc}
                        
                        for f_idx, (f_path, f_detail) in enumerate(inc['files'].items()):
                            child_iid = f"inc_cell_{idx}_{f_idx}"
                            row_tag = "even_row" if f_idx % 2 == 0 else "odd_row"
                            state_desc = f"{f_detail['file']}: {f_detail['value']}"
                            
                            self.tree.insert(coord_parent_iid, "end", iid=child_iid, text=state_desc,
                                             values=(inc['sheet'], inc['cell'], f_detail['state'].upper()), tags=(row_tag,))
                            self.node_data[child_iid] = {"type": "inconsistency_cell", "coordinate": inc['coordinate'], "file_detail": f_detail, "coord_info": inc}
                            
                            max_sheet_len = max(max_sheet_len, len(inc['sheet']))
                            max_cell_len = max(max_cell_len, len(inc['cell']))
                            max_matches_len = max(max_matches_len, len(state_desc))
                
                tree_width = min(max(max_tree_len * 8 + 40, 250), 380)
                sheet_width = min(max(max_sheet_len * 8 + 25, 120), 200)
                cell_width = min(max(max_cell_len * 8 + 25, 100), 130)
                matches_width = min(max(max_matches_len * 8 + 25, 110), 150)
                
                self.tree.column("#0", width=tree_width, minwidth=180, stretch=True)
                self.tree.column("Sheet", width=sheet_width, minwidth=100, stretch=True, anchor="w")
                self.tree.column("Cell", width=cell_width, minwidth=80, anchor="center", stretch=True)
                self.tree.column("Matches", width=matches_width, minwidth=100, anchor="center", stretch=True)
                
                for child in self.tree.get_children():
                    self.tree.item(child, open=True)
                    
        self._reset_state()

    def on_tree_select(self, event):
        selected_item = self.tree.selection()
        if not selected_item: return
        
        iid = selected_item[0]
        data = self.node_data.get(iid)
        if not data: return
        
        if self.current_mode == "IMAGE AUDIT":
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
                
            for child in self.preview_details_scroll.winfo_children(): child.destroy()
            
            if data["type"] == "group":
                group = data["group"]
                
                group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                group_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(group_frame, text=f"SET #{data['index'] + 1} OVERVIEW", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
                ctk.CTkLabel(group_frame, text=f"Total Occurrences: {len(group)} duplicates", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["primary"]).pack(anchor="w", padx=15, pady=(0, 12))
                
                grid_container = ctk.CTkFrame(group_frame, fg_color=COLORS["border"], corner_radius=8)
                grid_container.pack(fill="x", padx=15, pady=(0, 15))
                
                grid_container.grid_columnconfigure(0, weight=3)
                grid_container.grid_columnconfigure(1, weight=2)
                grid_container.grid_columnconfigure(2, weight=1)
                grid_container.grid_columnconfigure(3, weight=1)
                
                headers = ["File Name", "Sheet", "Cell", "Action"]
                for col_idx, h_text in enumerate(headers):
                    h_cell = ctk.CTkFrame(grid_container, fg_color="#F1F5F9", corner_radius=0)
                    h_cell.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(h_cell, text=h_text, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_dark"]).pack(padx=5, pady=6)
                    
                for row_idx, loc in enumerate(group):
                    r = row_idx + 1
                    bg_color = "#F8FAFC" if row_idx % 2 == 0 else "#FFFFFF"
                    
                    c0_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c0_frame.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c0_frame, text=loc['file'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=130, justify="left").pack(anchor="w", padx=6, pady=6)
                    
                    c1_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c1_frame.grid(row=r, column=1, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c1_frame, text=loc['sheet'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=80).pack(padx=5, pady=6)
                    
                    c2_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c2_frame.grid(row=r, column=2, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c2_frame, text=loc['cell'], font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["danger"]).pack(padx=5, pady=6)
                    
                    c3_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c3_frame.grid(row=r, column=3, sticky="nsew", padx=1, pady=1)
                    btn = ctk.CTkButton(c3_frame, text="OPEN", width=50, height=22, fg_color=COLORS["secondary"], hover_color=COLORS["primary"], corner_radius=4,
                                       font=ctk.CTkFont(size=9, weight="bold"), command=lambda p=loc['full_path'], s=loc['sheet'], c=loc['cell'], pic=loc.get('pic_name'): self.open_excel_at_location(p, s, c, pic))
                    btn.pack(padx=5, pady=4)
            else:
                loc = data["loc"]
                loc_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                loc_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(loc_frame, text="OCCURRENCE DETAILS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 10))
                
                f_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                f_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(f_box, text="File Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
                ctk.CTkLabel(f_box, text=loc['file'], font=ctk.CTkFont(size=12), text_color=COLORS["text_dark"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
                
                s_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                s_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(s_box, text="Sheet:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
                ctk.CTkLabel(s_box, text=loc['sheet'], font=ctk.CTkFont(size=12), text_color=COLORS["primary"], anchor="w").pack(side="left")
                
                c_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                c_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(c_box, text="Cell Address:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=80, anchor="w").pack(side="left")
                ctk.CTkLabel(c_box, text=loc['cell'], font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["danger"], anchor="w").pack(side="left")
                
                p_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                p_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(p_box, text="Full Path:", font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"], width=80, anchor="w").pack(side="left")
                ctk.CTkLabel(p_box, text=loc['full_path'], font=ctk.CTkFont(size=9), text_color=COLORS["text_muted"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
                
                btn = ctk.CTkButton(loc_frame, text="📂 OPEN EXCEL FILE", height=45, fg_color=COLORS["primary"], hover_color=COLORS["primary_glow"], corner_radius=10,
                                   font=ctk.CTkFont(size=13, weight="bold"), 
                                   command=lambda p=loc['full_path'], s=loc['sheet'], c=loc['cell'], pic=loc.get('pic_name'): self.open_excel_at_location(p, s, c, pic))
                btn.pack(fill="x", padx=15, pady=(20, 15))
        else:
            self.preview_img_lbl.configure(image="", text="⚠️ FORMULA INTEGRITY ALERT", text_color=COLORS["danger"])
            for child in self.preview_details_scroll.winfo_children(): child.destroy()
            
            if data["type"] == "hardcoded_root":
                group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                group_frame.pack(fill="both", expand=True, padx=5, pady=5)
                ctk.CTkLabel(group_frame, text="HARDCODED PASS/FAIL", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
                ctk.CTkLabel(group_frame, text=f"Total violations: {len(data['items'])} cells", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["danger"]).pack(anchor="w", padx=15, pady=(0, 12))
                ctk.CTkLabel(group_frame, text="Select a file or cell coordinate on the left for details and deep Excel navigation.", font=ctk.CTkFont(size=11), text_color=COLORS["text_dark"], justify="left", wraplength=270).pack(anchor="w", padx=15, pady=10)
                
            elif data["type"] == "inconsistency_root":
                group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                group_frame.pack(fill="both", expand=True, padx=5, pady=5)
                ctk.CTkLabel(group_frame, text="FORMULA INCONSISTENCIES", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
                ctk.CTkLabel(group_frame, text=f"Total discrepancies: {len(data['items'])} coords", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["danger"]).pack(anchor="w", padx=15, pady=(0, 12))
                ctk.CTkLabel(group_frame, text="Select a coordinate on the left to see the side-by-side formula mismatch grid across all workbooks.", font=ctk.CTkFont(size=11), text_color=COLORS["text_dark"], justify="left", wraplength=270).pack(anchor="w", padx=15, pady=10)
                
            elif data["type"] == "formula_file":
                f_name = data["file"]
                items = data["items"]
                
                group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                group_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(group_frame, text="FILE INTEGRITY REPORT", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
                ctk.CTkLabel(group_frame, text=f"Found: {len(items)} hardcoded cells", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["danger"]).pack(anchor="w", padx=15, pady=(0, 12))
                
                grid_container = ctk.CTkFrame(group_frame, fg_color=COLORS["border"], corner_radius=8)
                grid_container.pack(fill="x", padx=15, pady=(0, 15))
                
                grid_container.grid_columnconfigure(0, weight=2)
                grid_container.grid_columnconfigure(1, weight=1)
                grid_container.grid_columnconfigure(2, weight=1)
                grid_container.grid_columnconfigure(3, weight=1)
                
                headers = ["Sheet Name", "Cell", "Static Value", "Action"]
                for col_idx, h_text in enumerate(headers):
                    h_cell = ctk.CTkFrame(grid_container, fg_color="#F1F5F9", corner_radius=0)
                    h_cell.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(h_cell, text=h_text, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_dark"]).pack(padx=5, pady=6)
                    
                for row_idx, item in enumerate(items):
                    r = row_idx + 1
                    bg_color = "#F8FAFC" if row_idx % 2 == 0 else "#FFFFFF"
                    
                    c0_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c0_frame.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c0_frame, text=item['sheet'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=120).pack(padx=5, pady=6)
                    
                    c1_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c1_frame.grid(row=r, column=1, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c1_frame, text=item['cell'], font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["danger"]).pack(padx=5, pady=6)
                    
                    c2_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c2_frame.grid(row=r, column=2, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c2_frame, text=item['value'], font=ctk.CTkFont(size=10), text_color=COLORS["text_muted"]).pack(padx=5, pady=6)
                    
                    c3_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c3_frame.grid(row=r, column=3, sticky="nsew", padx=1, pady=1)
                    btn = ctk.CTkButton(c3_frame, text="OPEN", width=50, height=22, fg_color=COLORS["secondary"], hover_color=COLORS["primary"], corner_radius=4,
                                       font=ctk.CTkFont(size=9, weight="bold"), command=lambda p=item['full_path'], s=item['sheet'], c=item['cell']: self.open_excel_at_location(p, s, c))
                    btn.pack(padx=5, pady=4)
                    
            elif data["type"] == "formula_cell":
                item = data["item"]
                
                loc_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                loc_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(loc_frame, text="FORMULA VIOLATION DETAILS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 10))
                
                f_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                f_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(f_box, text="File Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(f_box, text=item['file'], font=ctk.CTkFont(size=12), text_color=COLORS["text_dark"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
                
                s_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                s_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(s_box, text="Sheet Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(s_box, text=item['sheet'], font=ctk.CTkFont(size=12), text_color=COLORS["primary"], anchor="w").pack(side="left")
                
                c_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                c_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(c_box, text="Cell Address:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(c_box, text=item['cell'], font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["danger"], anchor="w").pack(side="left")
                
                v_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                v_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(v_box, text="Static Value:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(v_box, text=f"'{item['value']}' (Hardcoded Bypass)", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["danger"], anchor="w").pack(side="left")
                
                desc_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                desc_box.pack(fill="x", padx=15, pady=10)
                ctk.CTkLabel(desc_box, text="Explanation: This cell evaluates to Pass/Fail but does not contain any formula. The user typed the string manually, which compromises report reliability.", font=ctk.CTkFont(size=10, slant="italic"), text_color=COLORS["text_muted"], wraplength=270, justify="left").pack(anchor="w")
                
                btn = ctk.CTkButton(loc_frame, text="📂 OPEN & FOCUS EXCEL CELL", height=45, fg_color=COLORS["primary"], hover_color=COLORS["primary_glow"], corner_radius=10,
                                   font=ctk.CTkFont(size=13, weight="bold"), 
                                   command=lambda p=item['full_path'], s=item['sheet'], c=item['cell']: self.open_excel_at_location(p, s, c))
                btn.pack(fill="x", padx=15, pady=(10, 15))
                
            elif data["type"] == "inconsistency_group":
                item = data["item"]
                coord = item["coordinate"]
                
                group_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                group_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(group_frame, text="FORMULA MISMATCH REPORT", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 5))
                ctk.CTkLabel(group_frame, text=f"Coord: {coord}", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["danger"]).pack(anchor="w", padx=15, pady=(0, 12))
                
                grid_container = ctk.CTkFrame(group_frame, fg_color=COLORS["border"], corner_radius=8)
                grid_container.pack(fill="x", padx=15, pady=(0, 15))
                
                grid_container.grid_columnconfigure(0, weight=2)
                grid_container.grid_columnconfigure(1, weight=3)
                grid_container.grid_columnconfigure(2, weight=1)
                
                headers = ["File Name", "Formula / State", "Action"]
                for col_idx, h_text in enumerate(headers):
                    h_cell = ctk.CTkFrame(grid_container, fg_color="#F1F5F9", corner_radius=0)
                    h_cell.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(h_cell, text=h_text, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_dark"]).pack(padx=5, pady=6)
                    
                for row_idx, (f_path, f_detail) in enumerate(item["files"].items()):
                    r = row_idx + 1
                    bg_color = "#F8FAFC" if row_idx % 2 == 0 else "#FFFFFF"
                    
                    c0_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c0_frame.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)
                    ctk.CTkLabel(c0_frame, text=f_detail['file'], font=ctk.CTkFont(size=10), text_color=COLORS["text_dark"], wraplength=120, justify="left").pack(anchor="w", padx=5, pady=6)
                    
                    c1_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c1_frame.grid(row=r, column=1, sticky="nsew", padx=1, pady=1)
                    
                    t_color = COLORS["danger"] if f_detail['state'] != 'formula' else COLORS["primary"]
                    disp_val = f_detail['value']
                    ctk.CTkLabel(c1_frame, text=disp_val, font=ctk.CTkFont(size=9, weight="bold" if f_detail['state'] != 'missing' else "normal"), text_color=t_color, wraplength=140, justify="left").pack(anchor="w", padx=5, pady=6)
                    
                    c2_frame = ctk.CTkFrame(grid_container, fg_color=bg_color, corner_radius=0)
                    c2_frame.grid(row=r, column=2, sticky="nsew", padx=1, pady=1)
                    btn = ctk.CTkButton(c2_frame, text="OPEN", width=50, height=22, fg_color=COLORS["secondary"], hover_color=COLORS["primary"], corner_radius=4,
                                       font=ctk.CTkFont(size=9, weight="bold"), command=lambda p=f_detail['full_path'], s=item['sheet'], c=item['cell']: self.open_excel_at_location(p, s, c))
                    btn.pack(padx=5, pady=4)
                    
            elif data["type"] == "inconsistency_cell":
                f_detail = data["file_detail"]
                coord_info = data["coord_info"]
                
                loc_frame = ctk.CTkFrame(self.preview_details_scroll, fg_color=COLORS["bg_canvas"], corner_radius=12, border_width=1, border_color=COLORS["border"])
                loc_frame.pack(fill="both", expand=True, padx=5, pady=5)
                
                ctk.CTkLabel(loc_frame, text="INCONSISTENCY DETAILS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", padx=15, pady=(15, 10))
                
                f_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                f_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(f_box, text="File Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(f_box, text=f_detail['file'], font=ctk.CTkFont(size=12), text_color=COLORS["text_dark"], anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
                
                s_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                s_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(s_box, text="Sheet Name:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(s_box, text=coord_info['sheet'], font=ctk.CTkFont(size=12), text_color=COLORS["primary"], anchor="w").pack(side="left")
                
                c_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                c_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(c_box, text="Cell Address:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(c_box, text=coord_info['cell'], font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["danger"], anchor="w").pack(side="left")
                
                v_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                v_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(v_box, text="Cell Content:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                
                disp_val = f_detail['value']
                t_color = COLORS["danger"] if f_detail['state'] != 'formula' else COLORS["primary"]
                ctk.CTkLabel(v_box, text=disp_val, font=ctk.CTkFont(size=12, weight="bold"), text_color=t_color, anchor="w", wraplength=200, justify="left").pack(side="left", fill="x", expand=True)
                
                state_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                state_box.pack(fill="x", padx=15, pady=4)
                ctk.CTkLabel(state_box, text="Cell State:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text_dark"], width=90, anchor="w").pack(side="left")
                ctk.CTkLabel(state_box, text=f_detail['state'].upper(), font=ctk.CTkFont(size=12, weight="bold"), text_color=t_color, anchor="w").pack(side="left")
                
                desc_box = ctk.CTkFrame(loc_frame, fg_color="transparent")
                desc_box.pack(fill="x", padx=15, pady=10)
                ctk.CTkLabel(desc_box, text="Explanation: The content of this cell differs from other audited files at the same coordinate. To ensure compliance, every file must share the exact same formula expressions.", font=ctk.CTkFont(size=10, slant="italic"), text_color=COLORS["text_muted"], wraplength=270, justify="left").pack(anchor="w")
                
                btn = ctk.CTkButton(loc_frame, text="📂 OPEN & FOCUS EXCEL CELL", height=45, fg_color=COLORS["primary"], hover_color=COLORS["primary_glow"], corner_radius=10,
                                   font=ctk.CTkFont(size=13, weight="bold"), 
                                   command=lambda p=f_detail['full_path'], s=coord_info['sheet'], c=coord_info['cell']: self.open_excel_at_location(p, s, c))
                btn.pack(fill="x", padx=15, pady=(10, 15))

    def open_excel_at_location(self, file_path, sheet_name, cell_address, pic_name=None):
        """Mở tệp Excel và tự động chọn đúng Sheet + Cell/Shape chứa ảnh trùng bằng win32com.
           Nếu máy không hỗ trợ win32com, hệ thống sẽ tự động hạ cấp xuống os.startfile."""
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
                
            # Chọn (Focus) vào đúng ô Cell hoặc Shape chứa ảnh
            try:
                found_shape = False
                
                # Biến đổi địa chỉ ô về chuẩn địa chỉ tuyệt đối của Excel (ví dụ $B$3)
                target_address = None
                if cell_address and cell_address != "N/A" and not cell_address.startswith("Float"):
                    try:
                        target_address = ws.Range(cell_address).Address
                    except Exception:
                        pass
                
                # 1. Thử tìm bằng Shape có vị trí TopLeftCell trùng khớp với ô Range
                if target_address:
                    try:
                        for shape in ws.Shapes:
                            try:
                                if shape.TopLeftCell.Address == target_address:
                                    shape.Select()
                                    found_shape = True
                                    break
                            except Exception:
                                pass
                    except Exception:
                        pass
                
                # 2. Thử tìm bằng pic_name nếu có
                if not found_shape and pic_name:
                    try:
                        ws.Shapes(pic_name).Select()
                        found_shape = True
                    except Exception:
                        pass
                        
                    if not found_shape:
                        try:
                            for shape in ws.Shapes:
                                if shape.Name.lower() == pic_name.lower():
                                    shape.Select()
                                    found_shape = True
                                    break
                        except Exception:
                            pass
                
                # 3. Fallback: Nếu không tìm thấy Shape, chọn ô Range thông thường
                if not found_shape and cell_address and cell_address != "N/A" and not cell_address.startswith("Float"):
                    ws.Range(cell_address).Select()
            except Exception:
                pass
                
            # Đưa cửa sổ Active lên trước
            try:
                excel.ActiveWindow.Activate()
            except Exception:
                pass
        except Exception:
            # Fallback về mở file mặc định nếu có bất cứ lỗi gì xảy ra
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
            
            if self.current_mode == "IMAGE AUDIT":
                worksheet = workbook.add_worksheet("Image Audit Results")
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
            else:
                hardcoded_violations = self.duplicates_cache.get("hardcoded", [])
                inconsistencies = self.duplicates_cache.get("inconsistencies", [])
                
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0F172A', 'font_color': 'white', 'border': 1, 'align': 'center'})
                cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                fail_fmt = workbook.add_format({'bold': True, 'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'border': 1, 'align': 'center'})
                pass_fmt = workbook.add_format({'bold': True, 'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1, 'align': 'center'})
                mismatch_fmt = workbook.add_format({'bg_color': '#FEF3C7', 'font_color': '#D97706', 'border': 1, 'align': 'center'})
                
                # 1. Sheet ô trị tĩnh Pass/Fail
                if hardcoded_violations or not inconsistencies:
                    ws_hc = workbook.add_worksheet("Hardcoded Pass-Fail")
                    headers_hc = ["FILE NAME", "SHEET NAME", "CELL ADDRESS", "HARDCODED VALUE", "ISSUE TYPE"]
                    for col, text in enumerate(headers_hc):
                        ws_hc.write(0, col, text, header_fmt)
                    ws_hc.set_column('A:E', 25)
                    
                    row_hc = 1
                    for item in hardcoded_violations:
                        val_fmt = pass_fmt if item['value'].strip().upper() == "PASS" else fail_fmt
                        ws_hc.write(row_hc, 0, item['file'], cell_fmt)
                        ws_hc.write(row_hc, 1, item['sheet'], cell_fmt)
                        ws_hc.write(row_hc, 2, item['cell'], cell_fmt)
                        ws_hc.write(row_hc, 3, item['value'], val_fmt)
                        ws_hc.write(row_hc, 4, item['type'], cell_fmt)
                        row_hc += 1
                        
                # 2. Sheet không đồng nhất công thức chéo
                if inconsistencies:
                    ws_inc = workbook.add_worksheet("Formula Inconsistencies")
                    headers_inc = ["COORDINATE", "SHEET NAME", "CELL ADDRESS"]
                    file_columns = [Path(f).name for f in self.selected_files]
                    headers_inc.extend(file_columns)
                    
                    for col, text in enumerate(headers_inc):
                        ws_inc.write(0, col, text, header_fmt)
                    ws_inc.set_column(0, len(headers_inc) - 1, 25)
                    
                    row_inc = 1
                    for inc in inconsistencies:
                        ws_inc.write(row_inc, 0, inc['coordinate'], cell_fmt)
                        ws_inc.write(row_inc, 1, inc['sheet'], cell_fmt)
                        ws_inc.write(row_inc, 2, inc['cell'], cell_fmt)
                        
                        for col_idx, f_path in enumerate(self.selected_files):
                            f_path_str = str(f_path)
                            f_detail = inc['files'].get(f_path_str, {"state": "missing", "value": "(No formula / Empty)"})
                            
                            f_fmt = cell_fmt
                            if f_detail['state'] == 'hardcoded':
                                f_fmt = fail_fmt
                            elif f_detail['state'] == 'missing':
                                f_fmt = mismatch_fmt
                                
                            ws_inc.write(row_inc, 3 + col_idx, f_detail['value'], f_fmt)
                        row_inc += 1

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
            "1. CHỌN CHẾ ĐỘ QUÉT: Tại 'AUDIT MODE', nhấp chọn 'IMAGE AUDIT' để tìm ảnh trùng lặp hoặc 'FORMULA AUDIT' để phát hiện các ô nhập Pass/Fail thủ công.\n\n"
            "2. CHỌN DỮ LIỆU: Nhấn 'SELECT DATA SOURCE' hoặc kéo thả các file Excel (.xlsx) trực tiếp vào ứng dụng.\n\n"
            "3. THỰC HIỆN: Nhấn 'EXECUTE AUDIT' để ứng dụng tiến hành phân tích siêu tốc.\n\n"
            "4. ĐỊNH VỊ SÂU: Nhấp chuột vào dòng chi tiết, bảng điều khiển bên phải sẽ cung cấp thông tin và nút 'OPEN'/'OPEN EXCEL FILE' để tự động mở tệp Excel và bôi đen tiêu điểm vào đúng ô tọa độ!"
        )
        ctk.CTkLabel(scroll, text=vn_text, font=ctk.CTkFont(size=13), justify="left", wraplength=520, text_color=COLORS["text_dark"]).pack(anchor="w", pady=(0,30))

        ctk.CTkLabel(scroll, text="🇺🇸 USER GUIDE", font=ctk.CTkFont(size=18, weight="bold"), text_color=COLORS["secondary"]).pack(anchor="w", pady=(0,10))
        en_text = (
            "1. SELECT AUDIT MODE: At 'AUDIT MODE', choose 'IMAGE AUDIT' to detect duplicate images or 'FORMULA AUDIT' to discover manual Pass/Fail cells.\n\n"
            "2. SELECT DATA: Click 'SELECT DATA SOURCE' or drag and drop Excel files (.xlsx) into the app.\n\n"
            "3. EXECUTE: Click 'EXECUTE AUDIT' to trigger the fast audit processor.\n\n"
            "4. DEEP FOCUS: Click any result item, and the right panel will show details along with 'OPEN'/'OPEN EXCEL FILE' buttons to automatically launch Excel and focus exact coordinates!"
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
