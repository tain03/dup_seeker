import os
import zipfile
import hashlib
import json
import threading
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
import io
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- ELITE PREMIUM DESIGN SYSTEM ---
COLORS = {
    "bg_canvas": "#F8FAFC",      # Lace White
    "sidebar_primary": "#0F172A", # Deep Slate (Luxury Dark)
    "sidebar_accent": "#1E293B",
    "primary": "#0284C7",        # Azure Blue
    "primary_glow": "#38BDF8",
    "secondary": "#6366F1",      # Indigo Accent
    "success": "#10B981",        # Emerald
    "danger": "#EF4444",         # Rose Red
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

class ExcelProcessor:
    @staticmethod
    def extract_data(file_path, ignore_sheets=None):
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
                                results.append({
                                    "hash": hashlib.md5(data).hexdigest(),
                                    "file": Path(file_path).name,
                                    "full_path": str(file_path),
                                    "sheet": s_real_name,
                                    "cell": cell_addr,
                                    "img_data": data
                                })
        except Exception as e:
            print(f"Error processing {file_path}: {e}")
        return results

class DuplicateApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("DUP-SEEKER - Elite Integrity Suite")
        
        # --- RESPONSIVE WINDOW SCALING ---
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Đặt kích thước mặc định phù hợp với cả laptop nhỏ
        width = min(1150, int(screen_width * 0.8))
        height = min(720, int(screen_height * 0.8))
        
        # Căn giữa màn hình
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(1000, 600)
        
        # Set Window Icon
        try:
            self.iconbitmap("assets/app.ico")
        except:
            pass
        
        self.configure(fg_color=COLORS["bg_canvas"])

        # State
        self.selected_files = []
        self.ignore_sheets_str = tk.StringVar(value="Sheet1, Form_Mau")
        self.is_scanning = False
        
        # Grid Structure
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._draw_sidebar()
        self._draw_main()
        
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.handle_drop)

    def _draw_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=COLORS["sidebar_primary"])
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        # Branding
        branding = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        branding.pack(pady=(30, 30), padx=25, fill="x")
        
        ctk.CTkLabel(branding, text="DUP-", 
                     font=ctk.CTkFont(size=24, weight="bold"), 
                     text_color=COLORS["white"]).pack(side="left")
        ctk.CTkLabel(branding, text="SEEKER", 
                     font=ctk.CTkFont(size=24, weight="bold"), 
                     text_color=COLORS["primary"]).pack(side="left")

        # Action Group
        group_lbl = ctk.CTkLabel(self.sidebar, text="MAIN ACTIONS", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text_muted"])
        group_lbl.pack(anchor="w", padx=35, pady=(0, 10))

        self.btn_select = ctk.CTkButton(self.sidebar, text="SELECT DATA SOURCE", 
                                       fg_color=COLORS["sidebar_accent"],
                                       hover_color=COLORS["primary"],
                                       border_width=1, border_color=COLORS["primary"],
                                       font=ctk.CTkFont(weight="bold"),
                                       height=50, corner_radius=12,
                                       command=self.select_files)
        self.btn_select.pack(pady=5, padx=30, fill="x")

        self.file_count_lbl = ctk.CTkLabel(self.sidebar, text="System standby", text_color=COLORS["text_muted"], font=ctk.CTkFont(size=12))
        self.file_count_lbl.pack(pady=5)

        # Configuration
        config_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        config_box.pack(pady=25, padx=25, fill="x")
        
        ctk.CTkLabel(config_box, text="EXCLUSION SETS", font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"]).pack(anchor="w", pady=(0,8))
        self.entry_ignore = ctk.CTkEntry(config_box, textvariable=self.ignore_sheets_str,
                                        fg_color=COLORS["sidebar_accent"], border_color=COLORS["sidebar_accent"],
                                        height=35, corner_radius=10, text_color=COLORS["white"])
        self.entry_ignore.pack(fill="x")

        # Bottom Actions (Reduced bottom padding)
        self.btn_scan = ctk.CTkButton(self.sidebar, text="EXECUTE AUDIT", 
                                     command=self.start_scan, 
                                     height=50, 
                                     fg_color=COLORS["primary"],
                                     hover_color=COLORS["primary_glow"],
                                     font=ctk.CTkFont(size=15, weight="bold"),
                                     corner_radius=12)
        self.btn_scan.pack(side="bottom", pady=30, padx=25, fill="x")

        self.progress = ctk.CTkProgressBar(self.sidebar, progress_color=COLORS["primary"], height=8, corner_radius=4)
        self.progress.pack(side="bottom", pady=(0, 20), padx=30, fill="x")
        self.progress.set(0)
        
        # Developer Credit
        ctk.CTkLabel(self.sidebar, text="Developed by ductai.nguyen", 
                     font=ctk.CTkFont(size=10), 
                     text_color=COLORS["text_muted"]).pack(side="bottom", pady=(0, 10))

    def _draw_main(self):
        self.main_view = ctk.CTkFrame(self, fg_color="transparent")
        self.main_view.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.main_view.grid_columnconfigure(0, weight=1)
        self.main_view.grid_rowconfigure(1, weight=1)

        # Dashboard Header (Metrics) - Reduced height
        metrics_grid = ctk.CTkFrame(self.main_view, fg_color="transparent", height=100)
        metrics_grid.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        metrics_grid.grid_propagate(False)
        metrics_grid.grid_columnconfigure((0,1,2), weight=1)

        self.m_files = self._metric_card(metrics_grid, "0", "TOTAL RECORDS", 0)
        self.m_dups = self._metric_card(metrics_grid, "0", "INTEGRITY ISSUES", 1, COLORS["danger"])
        self.m_total = self._metric_card(metrics_grid, "0", "PARSED OBJECTS", 2)

        # Content View
        self.scroll_view = ctk.CTkScrollableFrame(self.main_view, fg_color="transparent")
        self.scroll_view.grid(row=1, column=0, sticky="nsew")
        
        self.empty_msg = ctk.CTkLabel(self.scroll_view, text="Please provide data sources to begin analysis", 
                                     font=ctk.CTkFont(size=18, slant="italic"), text_color=COLORS["text_muted"])
        self.empty_msg.pack(pady=220)

    def _metric_card(self, parent, val, label, col, color=None):
        card = ctk.CTkFrame(parent, fg_color=COLORS["white"], corner_radius=16, border_width=1, border_color=COLORS["border"])
        card.grid(row=0, column=col, sticky="nsew", padx=10)
        
        v_lbl = ctk.CTkLabel(card, text=val, font=ctk.CTkFont(size=32, weight="bold"), text_color=color or COLORS["text_dark"])
        v_lbl.pack(pady=(25, 0))
        ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=10, weight="bold"), text_color=COLORS["text_muted"]).pack(pady=(0,20))
        return v_lbl

    def select_files(self):
        files = filedialog.askopenfilenames(title="System Selection", filetypes=[("Excel files", "*.xlsx")])
        if files:
            self.selected_files = list(files)
            self.file_count_lbl.configure(text=f"{len(self.selected_files)} resources active", text_color=COLORS["primary"])
            self.m_files.configure(text=str(len(self.selected_files)))
            if hasattr(self, 'empty_msg') and self.empty_msg: self.empty_msg.destroy(); self.empty_msg=None

    def handle_drop(self, event):
        import re
        matches = [f[0] if f[0] else f[1] for f in re.findall(r'\{(.*?)\}|(\S+)', event.data)]
        xlsx_files = [f for f in matches if f.lower().endswith('.xlsx')]
        if xlsx_files:
            self.selected_files = list(set(self.selected_files + xlsx_files))
            self.file_count_lbl.configure(text=f"{len(self.selected_files)} resources active", text_color=COLORS["primary"])
            self.m_files.configure(text=str(len(self.selected_files)))
            if hasattr(self, 'empty_msg') and self.empty_msg: self.empty_msg.destroy(); self.empty_msg=None

    def start_scan(self):
        if self.is_scanning or not self.selected_files: return
        self.is_scanning = True
        self.btn_scan.configure(state="disabled", text="ANALYZING SYSTEM...")
        threading.Thread(target=self.scan_logic, daemon=True).start()

    def scan_logic(self):
        try:
            for child in self.scroll_view.winfo_children(): child.destroy()
            ignore_list = [s.strip() for s in self.ignore_sheets_str.get().split(",") if s.strip()]
            all_images = []
            for i, f in enumerate(self.selected_files):
                all_images.extend(ExcelProcessor.extract_data(f, ignore_list))
                self.after(0, lambda i=i: self.progress.set((i + 1) / len(self.selected_files)))

            db = defaultdict(list)
            for img in all_images: db[img['hash']].append(img)
            duplicates = [locs for h, locs in db.items() if len(locs) > 1]
            self.after(0, lambda: self.render_results(duplicates, len(all_images)))
        except Exception as e:
            self.after(0, lambda e=e: messagebox.showerror("System Error", f"{e}"))
            self.after(0, self._reset_state)

    def render_results(self, duplicates, total_count):
        self.m_dups.configure(text=str(len(duplicates)))
        self.m_total.configure(text=str(total_count))
        
        if not duplicates:
            ctk.CTkLabel(self.scroll_view, text="SYSTEM INTEGRITY VERIFIED", font=ctk.CTkFont(size=20, weight="bold"), text_color=COLORS["success"]).pack(pady=150)
        else:
            for i, group in enumerate(duplicates):
                card = ctk.CTkFrame(self.scroll_view, fg_color=COLORS["white"], corner_radius=16, border_width=1, border_color=COLORS["border"])
                card.pack(fill="x", pady=12, padx=5)
                
                # Accent Bar
                accent = ctk.CTkFrame(card, width=6, fg_color=COLORS["danger"], corner_radius=0)
                accent.place(relx=0, rely=0, relheight=1)

                try:
                    img_data = group[0]['img_data']
                    pil_img = Image.open(io.BytesIO(img_data))
                    pil_img.thumbnail((150, 150))
                    ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
                    img_lbl = ctk.CTkLabel(card, image=ctk_img, text="")
                    img_lbl.grid(row=0, column=0, rowspan=len(group)+1, padx=(25, 20), pady=25)
                except: pass

                header = ctk.CTkLabel(card, text=f"DUPLICATE SET #{i+1}", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLORS["danger"])
                header.grid(row=0, column=1, sticky="w", pady=(20,10))

                for j, loc in enumerate(group):
                    row = ctk.CTkFrame(card, fg_color="transparent")
                    row.grid(row=j+1, column=1, sticky="w", padx=10, pady=5)
                    
                    ctk.CTkLabel(row, text=loc['file'], font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=5)
                    ctk.CTkLabel(row, text=f"[{loc['sheet']}]", font=ctk.CTkFont(size=11), text_color=COLORS["text_muted"]).pack(side="left", padx=5)
                    ctk.CTkLabel(row, text=loc['cell'], font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["primary"]).pack(side="left", padx=5)
                    
                    btn = ctk.CTkButton(card, text="VIEW FILE", width=85, height=28, fg_color=COLORS["bg_canvas"], hover_color=COLORS["border"], text_color=COLORS["text_dark"], border_width=1, border_color=COLORS["border"], corner_radius=8,
                                       font=ctk.CTkFont(size=10, weight="bold"), command=lambda p=loc['full_path']: os.startfile(p))
                    btn.grid(row=j+1, column=2, padx=30)

        self._reset_state()

    def _reset_state(self):
        self.is_scanning = False
        self.btn_scan.configure(state="normal", text="EXECUTE AUDIT")
        self.progress.set(0)

if __name__ == "__main__":
    app = DuplicateApp()
    app.mainloop()
