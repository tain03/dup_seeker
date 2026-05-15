import os
import zipfile
import hashlib
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

# Namespace trong file XML của Excel
NS = {
    'sh': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'dr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

def get_cell_address(col, row):
    """Chuyển đổi số cột/dòng thành địa chỉ Excel (ví dụ: 0,0 -> A1)"""
    string = ""
    col += 1
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        string = chr(65 + remainder) + string
    return f"{string}{row + 1}"

class ExcelImageChecker:
    def __init__(self, data_dir, output_dir):
        self.data_dir = Path(data_dir)
        self.output_dir = Path(output_dir)
        self.image_db = defaultdict(list) # hash -> [list of locations]

    def get_hash(self, data):
        return hashlib.md5(data).hexdigest()

    def process_file(self, file_path):
        print(f"--- Đang xử lý: {file_path.name} ---")
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                # 1. Tìm quan hệ giữa Sheet và Drawing
                sheet_to_drawing = {}
                for f in z.namelist():
                    if f.startswith('xl/worksheets/_rels/sheet') and f.endswith('.xml.rels'):
                        sheet_num = f.replace('xl/worksheets/_rels/sheet', '').replace('.xml.rels', '')
                        content = z.read(f)
                        root = ET.fromstring(content)
                        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if 'drawing' in rel.get('Type'):
                                target = rel.get('Target').replace('../', 'xl/')
                                sheet_to_drawing[target] = f"Sheet{sheet_num}"

                # 2. Lấy tên thực của các Sheet (từ workbook.xml)
                sheet_names = {}
                workbook_xml = z.read('xl/workbook.xml')
                root = ET.fromstring(workbook_xml)
                for s in root.findall('.//sh:sheet', NS):
                    sheet_id = s.get('sheetId')
                    name = s.get('name')
                    sheet_names[f"Sheet{sheet_id}"] = name

                # 3. Phân tích từng Drawing để tìm ảnh và tọa độ
                for drawing_path, sheet_id in sheet_to_drawing.items():
                    if drawing_path not in z.namelist():
                        continue
                        
                    sheet_name = sheet_names.get(sheet_id, sheet_id)
                    
                    # Tìm quan hệ trong drawing (để map rId ra file ảnh)
                    rel_path = f"xl/drawings/_rels/{Path(drawing_path).name}.rels"
                    drawing_to_media = {}
                    if rel_path in z.namelist():
                        rel_content = z.read(rel_path)
                        rel_root = ET.fromstring(rel_content)
                        for rel in rel_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            target = rel.get('Target').replace('../', 'xl/')
                            drawing_to_media[rel.get('Id')] = target

                    # Đọc tọa độ ảnh trong drawing.xml
                    drawing_content = z.read(drawing_path)
                    d_root = ET.fromstring(drawing_content)
                    
                    # 1. Xử lý Two Cell và One Cell Anchor
                    for anchor in d_root.findall('.//dr:twoCellAnchor', NS) + d_root.findall('.//dr:oneCellAnchor', NS):
                        from_tag = anchor.find('dr:from', NS)
                        if from_tag is not None:
                            col = int(from_tag.find('dr:col', NS).text)
                            row = int(from_tag.find('dr:row', NS).text)
                            cell_addr = get_cell_address(col, row)
                            self._add_image_from_anchor(anchor, z, drawing_to_media, file_path.name, sheet_name, cell_addr)

                    # 2. Xử lý Absolute Anchor
                    for anchor in d_root.findall('.//dr:absoluteAnchor', NS):
                        pos = anchor.find('dr:pos', NS)
                        if pos is not None:
                            x = pos.get('x')
                            y = pos.get('y')
                            cell_addr = f"Floating (X:{x}, Y:{y})"
                            self._add_image_from_anchor(anchor, z, drawing_to_media, file_path.name, sheet_name, cell_addr)
                            
        except Exception as e:
            print(f"Lỗi khi xử lý file {file_path.name}: {e}")

    def _add_image_from_anchor(self, anchor, z_file, drawing_to_media, file_name, sheet_name, cell_addr):
        # Tìm rId của ảnh
        blip = anchor.find('.//a:blip', NS)
        if blip is not None:
            r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            media_path = drawing_to_media.get(r_id)
            
            if media_path and media_path in z_file.namelist():
                img_data = z_file.read(media_path)
                img_hash = self.get_hash(img_data)
                
                self.image_db[img_hash].append({
                    "file": file_name,
                    "sheet": sheet_name,
                    "cell": cell_addr,
                    "media_path": media_path
                })

    def run(self):
        files = list(self.data_dir.glob("*.xlsx"))
        if not files:
            print("Không tìm thấy file .xlsx nào trong thư mục DATA!")
            return

        for f in files:
            self.process_file(f)

        # Lọc ra các ảnh trùng lặp
        duplicates = {h: locs for h, locs in self.image_db.items() if len(locs) > 1}

        # Lưu kết quả
        report_path = self.output_dir / "duplicate_report.json"
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(duplicates, f, indent=4, ensure_ascii=False)
        
        print(f"\n--- HOÀN THÀNH ---")
        print(f"Đã quét xong {len(files)} file.")
        print(f"Tìm thấy {len(duplicates)} bộ ảnh bị dùng lại.")
        print(f"Kết quả lưu tại: {report_path}")

if __name__ == "__main__":
    checker = ExcelImageChecker("DATA", "OUTPUT")
    checker.run()
