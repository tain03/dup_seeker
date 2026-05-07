# 🛠️ DUP-SEEKER Development Log

This file tracks the historical development progress of the Elite Excel Audit Suite.

* **[2026-05-03 15:28:24]** refactor: minor code cleanup and optimization pass #93
* **[2026-05-03 16:57:53]** style: design custom wrap lengths for long filenames in panel
* **[2026-05-03 15:55:06]** style: add alternating zebra striping colors
* **[2026-05-03 15:32:00]** refactor: optimize drawing relative ID mappings
* **[2026-05-03 17:27:17]** feat: implement capped dynamic column width (380px)
* **[2026-05-03 14:18:58]** style: add drop zone visual border feedback
* **[2026-05-03 11:20:25]** fix: handle multiple excel instances cleanly
* **[2026-05-04 18:07:09]** fix: resolve floating image anchor cell coordinates correctly
* **[2026-05-04 11:22:00]** feat: implement pywin32 COM connection utility
* **[2026-05-04 15:48:09]** test: verify excel focus automation on locked sheets
* **[2026-05-04 16:43:54]** test: benchmark md5 hashing vs phash generation
* **[2026-05-04 17:30:20]** fix: trim whitespaces in exclusion sheets input
* **[2026-05-04 18:48:56]** feat: add main layout split frame design
* **[2026-05-04 16:06:19]** feat: implement drag and drop wrapper using tkinterdnd2
* **[2026-05-05 16:11:59]** style: customize sidebar panel backgrounds
* **[2026-05-05 18:34:48]** feat: implement base ZIP reader for excel workbook
* **[2026-05-05 12:00:15]** fix: close SQLite connection safely on worker exit
* **[2026-05-05 10:09:58]** feat: integrate Pillow for image format decoding
* **[2026-05-05 12:44:21]** feat: create current_scan temporary table for bulk write
* **[2026-05-05 11:06:18]** style: set cell content anchors to center for cell addresses
* **[2026-05-05 17:15:32]** fix: clear old tree nodes upon new audit trigger
* **[2026-05-05 17:28:05]** perf: reuse SQL connection statement pool
* **[2026-05-05 18:55:23]** feat: add tkinterdnd2 DLL paths to binaries array in spec
* **[2026-05-06 17:57:04]** feat: parse worksheet relation files for drawing targets
* **[2026-05-06 13:54:58]** fix: fix file path parsing for drop paths on Windows
* **[2026-05-06 14:15:34]** test: add unit test for multi-sheet drawing parsers
* **[2026-05-06 15:04:45]** feat: integrate imagehash for perceptual hashing
* **[2026-05-06 10:45:35]** refactor: abstract color definitions to config structure
* **[2026-05-07 18:24:50]** fix: handle file-not-found fallback to os.startfile
* **[2026-05-07 10:44:44]** fix: release excel COM objects safely to avoid zombie processes
* **[2026-05-07 09:19:31]** docs: split Vietnamese version to README_VN.md with flags
* **[2026-05-07 17:08:14]** refactor: minor code cleanup and optimization pass #91
