# OJT Report Scanner

A local website that scans cover-page photos of OJT training narrative reports. It extracts the company, address, duration, student, program, and section; lets you correct OCR results; filters records by OJT year; and downloads an Excel workbook.

## Run on Windows

1. Install Python 3.10 or newer and dependencies: `python -m pip install -r requirements.txt`.
2. Double-click `start.bat`, or run `python app.py` in this folder.
3. Open `http://127.0.0.1:8765` if your browser does not open automatically.

Windows' built-in English OCR is used. No cloud account or API key is needed. Upload JPG, PNG, BMP, TIFF, or WebP cover pages, including multiple images or a folder. Images stay on your computer; temporary OCR copies are deleted after each scan. Your corrected records are saved in this browser's local storage. Download Excel to keep a portable copy.

**Review every scanned record.** OCR can confuse letters and dates in dim or angled photos. The Review / edit dialog shows the OCR text and lets you correct all fields. The exported workbook marks unreviewed records as `Needs review`.

Program majors are classified under their base program: for example, `BSCS-DA` becomes `BSCS` and `BSIT-GD` becomes `BSIT`. Full program names and spaced or dotted abbreviations are also recognized. The class section is kept separately.

Sections can appear on the next line or alongside the program: `BSCS - J4S` gives program `BSCS` and section `J4S`, including when a major precedes the section.

The scanner preserves `ñ` and `Ñ` through review and Excel export. It corrects the confirmed OCR misreads `Bifian`, `Nifio`, and `Dufigo` (or `Duigo`) to `Biñan`, `Niño`, and `Duñgo`, preserving capitalization and the original OCR text for review. Other ambiguous names still need manual review; genuine `fi` letter pairs are kept. Existing saved records need to be corrected or rescanned.

If a duration spans two calendar years, the workbook contains one row for each year so Excel's Year filter shows the report in both years. Choosing a year on the website downloads only that year's rows.

## Tests

Run `python -m unittest discover -s tests -v`.
