# 🚀 DUP-SEEKER: Premium Excel Image Integrity Suite

> [🇻🇳 Tiếng Việt / Vietnamese Edition](README_VN.md)

---

**DUP-SEEKER** is an Enterprise-Grade Desktop solution designed to automate duplicate image auditing and formula integrity scanning across bulk Excel reports (`.xlsx`). It guarantees absolute data integrity and layout protection in corporate technical documents.

![GitHub release (latest by date)](https://img.shields.io/badge/release-v3.1--Final-emerald.svg?style=flat-square)
![Python](https://img.shields.io/badge/python-3.12-sky.svg?style=flat-square)
![UI](https://img.shields.io/badge/UI-CustomTkinter-indigo.svg?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)

---

## 🌟 Core Excellence (The Best of DUP-SEEKER)

*   **🔌 COM Deep-Linking & Auto-Focus:** Rather than just opening the file, DUP-SEEKER hooks directly into **Microsoft Excel** using Windows COM interfaces. Upon clicking **`OPEN`**, it dynamically activates the workbook, switches to the exact sheet, and **directly selects & scroll-focuses on the exact Picture Shape** (or cell range fallback) containing the duplicate image.
*   **🛡️ One-Touch Formula Protection & 🔓 Bulk Unprotect:** Secure all formulas instantly with a new **`🛡️ PROTECT FORMULAS`** sidebar button. It utilizes Excel's native `SpecialCells` to automatically lock all formula cells and protect worksheets with password `'1'`, while keeping raw input data and shape drawings fully editable. Use the new **`🔓 UNPROTECT FORMULAS`** button to remove protection from all selected workbooks in bulk!
*   **🔍 Smart Search & Real-time Filter Bar:** Quickly filter through massive audit tables. A classy modern search box at the top of the Treeview dynamically matches results against filenames, sheets, cell locations, or values as you type.
*   **⚡ Hyper-Speed Double-Hash Engine:** Integrates a dual image-hashing network (**MD5 + AI Vision pHash**) backed by a **permanent SQLite cache database**. It processes thousands of images in seconds without duplicate calculations.
*   **💎 Elite Corporate UI/UX:** Built on CustomTkinter with a premium Slate-Blue palette. Features:
    *   **Zebra-striped Hierarchical Treeview** with capped dynamic column width (`380px`) to prevent long filenames from breaking the layout.
    *   **Right Panel Spreadsheet Grid Table** showing all occurrences in a clean cell-bordered table with immediate action buttons.
    *   **Taskbar Icon Sync & DPI Maximization Fix** ensuring native Windows integration on startup.

---

## 🛠️ Installation & Usage Guide

> [!NOTE]
> **No installation is required!** DUP-SEEKER is fully portable.

### Option 1: Run Pre-compiled Standalone EXE (For End-Users)
1. Download and extract **`DUP-SEEKER_v3.1_Final.zip`**.
2. Double-click **`dist/DUP-SEEKER/DUP-SEEKER.exe`** to launch instantly.

### Option 2: Run from Python Source (For Developers)
1. Install required library dependencies:
   ```bash
   pip install customtkinter pillow tkinterdnd2-universal imagehash xlsxwriter pywin32
   ```
2. Launch the application:
   ```bash
   python app.py
   ```

---

## 👥 Authors & License
*   **Developed by:** [ductai.nguyen](https://github.com/ductai-nguyen)
*   **License:** MIT License. All rights reserved.
