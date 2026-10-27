"""Local OJT cover-page scanner and Excel exporter. Run: python app.py"""
from __future__ import annotations

import base64
import datetime as dt
import io
import json
import re
import subprocess
import tempfile
import threading
import unicodedata
import uuid
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent
MAX_UPLOAD = 16 * 1024 * 1024
MONTHS = {m.lower(): i for i, m in enumerate(('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'), 1)}
MONTHS.update({m[:3]: n for m, n in list(MONTHS.items())})
DATE = re.compile(r'\b(' + '|'.join(sorted(MONTHS, key=len, reverse=True)) + r')\.?\s+(\d{1,2})\s*[,\.]?\s*(20\d{2})\b', re.I)
FIELDS = ('company', 'address', 'duration', 'student', 'program', 'section')
PROGRAM = re.compile(
    r'\b(?:(?P<cs>B[ .]*S[ .]*C[ .]*S|Bachelor\s+of\s+Science\s+in\s+Computer\s+Science)'
    r'|(?P<it>B[ .]*S[ .]*I[ .]*T|Bachelor\s+of\s+Science\s+in\s+Information\s+Technology))\b',
    re.I,
)
SECTION = re.compile(r'\bJ\d{1,2}[A-Z]\b', re.I)
# Confirmed against the supplied cover photos; never replace every "fi".
ENYE_CORRECTIONS = {'bifian': 'biñan', 'nifio': 'niño', 'niöo': 'niño', 'dufigo': 'duñgo', 'duigo': 'duñgo'}
ENYE_OCR_WORDS = re.compile(r'\b(?:' + '|'.join(ENYE_CORRECTIONS) + r')\b', re.I)


def restore_enye(match: re.Match[str]) -> str:
    original = match.group()
    corrected = ENYE_CORRECTIONS[original.lower()]
    if original.isupper():
        return corrected.upper()
    return corrected.capitalize() if original[0].isupper() else corrected


def clean(value: str) -> str:
    value = ENYE_OCR_WORDS.sub(restore_enye, unicodedata.normalize('NFC', value))
    return re.sub(r'\s+', ' ', value.strip().strip('<>‹›[]{} ')).strip()


def parse_date(match: re.Match[str]) -> dt.date | None:
    try:
        return dt.date(int(match.group(3)), MONTHS[match.group(1).lower().rstrip('.')], int(match.group(2)))
    except ValueError:
        return None


def parse_report(text: str, filename: str) -> dict:
    lines = [clean(line) for line in text.splitlines() if clean(line)]
    lower = [line.lower() for line in lines]
    row = {key: '' for key in FIELDS}
    row.update({'id': str(uuid.uuid4()), 'source': filename, 'start_date': '', 'end_date': '', 'raw_text': text, 'verified': False})
    title = next((i for i, line in enumerate(lower) if 'narrative' in line and 'report' in line), -1)
    if title < 0:
        title = next((i for i, line in enumerate(lower) if 'train' in line and 'report' in line), -1)
    if title >= 0:
        block_end = next((i for i in range(title + 1, len(lines)) if any(marker in lower[i] for marker in ('duration', 'submitted by', 'submitted to'))), len(lines))
        following = lines[title + 1:block_end]
        if following:
            row['company'] = following[0]
        row['address'] = ' '.join(line for line in following[1:] if not re.search(r'\b(?:department|dept\.?)\b', line, re.I))
    duration_index = next((i for i, line in enumerate(lower) if 'duration' in line and ('ojt' in line or line == 'duration')), -1)
    if duration_index >= 0 and duration_index + 1 < len(lines):
        row['duration'] = lines[duration_index + 1]
    dates = [parse_date(m) for m in DATE.finditer(row['duration'])]
    dates = [date for date in dates if date]
    if len(dates) >= 2 and dates[0] <= dates[1]:
        row['start_date'], row['end_date'] = (date.isoformat() for date in dates[:2])
    submit = next((i for i, line in enumerate(lower) if 'submitted by' in line), -1)
    if submit >= 0 and submit + 1 < len(lines):
        row['student'] = lines[submit + 1]
    program_match = next(((i, match) for i in range(max(0, submit + 1), len(lines)) if (match := PROGRAM.search(lines[i]))), None)
    if program_match is not None:
        program, match = program_match
        row['program'] = 'BSCS' if match.group('cs') else 'BSIT'
        inline_section = SECTION.search(lines[program], match.end())
        if inline_section:
            row['section'] = inline_section.group().upper()
        elif program + 1 < len(lines) and not lower[program + 1].startswith('submitted to'):
            candidate = lines[program + 1]
            if re.fullmatch(r'[A-Za-z0-9/ -]{2,18}', candidate):
                row['section'] = candidate
    row['needs_review'] = [field for field in FIELDS if not row[field]]
    if not row['start_date'] or not row['end_date']:
        row['needs_review'].append('dates')
    return row


def repair_saved_record(record: dict) -> dict:
    if record.get('verified') or not record.get('raw_text'):
        return dict(record)
    parsed = parse_report(record['raw_text'].split('[Student area OCR]')[0], record.get('source', ''))
    repaired = dict(record)
    if parsed['company']:
        repaired['company'] = parsed['company']
        repaired['address'] = parsed['address']
    if repaired.get('student'):
        repaired['student'] = clean(repaired['student'])
    repaired['needs_review'] = [field for field in FIELDS if not repaired.get(field)]
    if not repaired.get('start_date') or not repaired.get('end_date'):
        repaired['needs_review'].append('dates')
    return repaired


def years_for_record(row: dict) -> list[int]:
    try:
        start, end = (dt.date.fromisoformat(row[key]) for key in ('start_date', 'end_date'))
        if start > end or end.year - start.year > 10:
            return []
        return list(range(start.year, end.year + 1))
    except (KeyError, ValueError, TypeError):
        return []


def ocr_image(path: Path) -> str:
    result = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ROOT / 'ocr.ps1'), '-Path', str(path)], capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=45)
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or 'Windows OCR failed')
    return result.stdout.strip()


def scan_image(data: bytes, filename: str) -> dict:
    from PIL import Image, ImageEnhance, ImageOps
    with Image.open(io.BytesIO(data)) as original:
        image = original.convert('RGB')
    with tempfile.TemporaryDirectory(dir=ROOT) as temp:
        folder = Path(temp)
        original_path = folder / 'original.png'
        image.save(original_path)
        plain = ocr_image(original_path)
        enhanced = ImageEnhance.Sharpness(ImageEnhance.Contrast(ImageOps.autocontrast(image, cutoff=2)).enhance(1.4)).enhance(1.8)
        enhanced_path = folder / 'enhanced.png'
        enhanced.save(enhanced_path)
        enhanced_text = ocr_image(enhanced_path)
        candidates = [parse_report(plain, filename), parse_report(enhanced_text, filename)]
        chosen = min(candidates, key=lambda row: (len(row['needs_review']), -len(row['raw_text'])))
        width, height = image.size
        crop = image.crop((int(width * .25), int(height * .65), int(width * .78), int(height * .87)))
        crop = ImageEnhance.Contrast(ImageOps.autocontrast(crop.convert('L'), cutoff=1).resize((crop.width * 3, crop.height * 3))).enhance(2)
        crop_path = folder / 'student.png'
        crop.save(crop_path)
        crop_text = ocr_image(crop_path)
        crop_row = parse_report('TRAINING NARRATIVE REPORT\n\n' + crop_text, filename)
        if not chosen['program'] and crop_row['program']:
            chosen['program'] = crop_row['program']
        if not chosen['section'] and crop_row['section']:
            chosen['section'] = crop_row['section']
        chosen['needs_review'] = [field for field in FIELDS if not chosen[field]]
        if not chosen['start_date'] or not chosen['end_date']:
            chosen['needs_review'].append('dates')
        chosen['raw_text'] = chosen['raw_text'] + '\n\n[Student area OCR]\n' + crop_text
        return chosen


def xml_text(value: object) -> str:
    return escape(str(value), {'"': '&quot;', "'": '&apos;'})


def make_workbook(records: list[dict], selected_year: int | None = None) -> bytes:
    headers = ['Year', 'Company', 'Address', 'OJT Duration', 'Start Date', 'End Date', 'Student', 'Program', 'Section', 'Source File', 'Review Status']
    rows = [headers]
    for record in records:
        years = years_for_record(record) or ([''] if selected_year is None else [])
        if selected_year is not None:
            years = [year for year in years if year == selected_year]
        for year in years:
            rows.append([year, record.get('company', ''), record.get('address', ''), record.get('duration', ''), record.get('start_date', ''), record.get('end_date', ''), record.get('student', ''), record.get('program', ''), record.get('section', ''), record.get('source', ''), 'Reviewed' if record.get('verified') else 'Needs review'])
    cells = []
    for row_num, values in enumerate(rows, 1):
        parts = []
        for index, value in enumerate(values):
            column = chr(65 + index)
            if index == 0 and row_num > 1 and isinstance(value, int):
                parts.append(f'<c r="{column}{row_num}" s="{1 if row_num == 1 else 0}"><v>{value}</v></c>')
            else:
                style = ' s="1"' if row_num == 1 else ''
                parts.append(f'<c r="{column}{row_num}" t="inlineStr"{style}><is><t>{xml_text(value)}</t></is></c>')
        cells.append(f'<row r="{row_num}">{"".join(parts)}</row>')
    last = max(1, len(rows))
    sheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i,w in enumerate([10,30,52,32,16,16,32,13,16,24,18],1))}</cols><sheetData>{''.join(cells)}</sheetData><autoFilter ref="A1:K{last}"/></worksheet>'''
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        archive.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="OJT Records" sheetId="1" r:id="rId1"/></sheets></workbook>')
        archive.writestr('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        archive.writestr('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF18334A"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>')
        archive.writestr('xl/worksheets/sheet1.xml', sheet)
    return output.getvalue()


class Handler(BaseHTTPRequestHandler):
    def send_bytes(self, data: bytes, content_type: str, status: int = 200, filename: str | None = None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        if filename:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if urlparse(self.path).path in ('/', '/index.html'):
            self.send_bytes((ROOT / 'index.html').read_bytes(), 'text/html; charset=utf-8')
        else:
            self.send_error(404)

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= MAX_UPLOAD * 2:
                raise ValueError('Upload is too large.')
            request = json.loads(self.rfile.read(length))
            if self.path == '/api/scan':
                filename = Path(str(request['filename'])).name
                if Path(filename).suffix.lower() not in ('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp'):
                    raise ValueError('Use a JPG, PNG, BMP, TIFF, or WebP image.')
                data = base64.b64decode(request['image'], validate=True)
                if len(data) > MAX_UPLOAD:
                    raise ValueError('Image is too large (16 MB maximum).')
                result = scan_image(data, filename)
                self.send_bytes(json.dumps(result).encode(), 'application/json; charset=utf-8')
            elif self.path == '/api/reparse':
                record = request['record']
                if not isinstance(record, dict):
                    raise ValueError('Invalid record.')
                self.send_bytes(json.dumps(repair_saved_record(record)).encode(), 'application/json; charset=utf-8')
            elif self.path == '/api/export':
                records = request['records']
                if not isinstance(records, list) or len(records) > 10000:
                    raise ValueError('Invalid record list.')
                selected_year = request.get('year', 'all')
                if selected_year != 'all':
                    selected_year = int(selected_year)
                    if not 1900 <= selected_year <= 2100:
                        raise ValueError('Invalid year.')
                else:
                    selected_year = None
                self.send_bytes(make_workbook(records, selected_year), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='ojt-records.xlsx')
            else:
                self.send_error(404)
        except Exception as exc:
            self.send_bytes(json.dumps({'error': str(exc)}).encode(), 'application/json; charset=utf-8', 400)


if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 8765), Handler)
    print('OJT Scanner: http://127.0.0.1:8765', flush=True)
    threading.Timer(1, lambda: webbrowser.open('http://127.0.0.1:8765')).start()
    server.serve_forever()
