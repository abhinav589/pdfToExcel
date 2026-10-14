#!/usr/bin/env python3
"""End-to-end OCR extraction for the attached Hindi voter-list PDF.

Usage: python3 279_pdf_to_excel.py input.pdf output.xlsx [--expected-records N] [--first-page N] [--last-page N]
This supports PDFs using the same three-column, ten-row voter-card layout.
Requirements: Tesseract with Hindi and English traineddata, Poppler (pdftoppm/pdfinfo),
Pillow, and openpyxl. OCR output can contain errors; review the resulting sheet
against the source PDF before relying on individual voter details.
"""
import os
import re
import ast
import json
import subprocess
import collections
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('pdf', help='Input voter-list PDF')
parser.add_argument('xlsx', help='Output Excel workbook path')
parser.add_argument('--expected-records', type=int, default=None, help='Optional expected number of records; reports whether extraction matched')
parser.add_argument('--first-page', type=int, default=3, help='First voter-list PDF page (default: 3)')
parser.add_argument('--last-page', type=int, default=None, help='Last voter-list PDF page (default: final PDF page)')
args = parser.parse_args()
pdf_path = Path(args.pdf).expanduser().resolve()
xlsx_path = Path(args.xlsx).expanduser().resolve()
if not pdf_path.is_file():
    raise FileNotFoundError(pdf_path)
xlsx_path.parent.mkdir(parents=True, exist_ok=True)
# Keep OCR artifacts out of the deployed source tree. The web app supplies a
# per-job temporary parent; direct CLI runs default to a work/ folder beside this script.
script_dir = Path(__file__).resolve().parent
work_parent = Path(os.environ.get('VOTER_OCR_WORK_PARENT', str(script_dir))).expanduser().resolve()
work_parent.mkdir(parents=True, exist_ok=True)
os.chdir(work_parent)
Path('work/scan').mkdir(parents=True, exist_ok=True)
Path('work/scan300').mkdir(parents=True, exist_ok=True)
Path('work/cards').mkdir(parents=True, exist_ok=True)

# Determine the page range. The expected record count is an optional completeness check.
if args.expected_records is not None and args.expected_records < 0:
    parser.error('--expected-records must be zero or greater')
if args.first_page < 1:
    parser.error('--first-page must be 1 or greater')
if args.last_page is None:
    info = subprocess.run(['pdfinfo', str(pdf_path)], capture_output=True, text=True, check=True).stdout
    match = re.search(r'^Pages:\s+(\d+)', info, re.MULTILINE)
    if not match:
        raise RuntimeError('Could not determine PDF page count with pdfinfo; pass --last-page explicitly')
    last_page = int(match.group(1))
else:
    last_page = args.last_page
if last_page < args.first_page:
    parser.error('--last-page must be greater than or equal to --first-page')
page_numbers = range(args.first_page, last_page + 1)

# Render voter-list pages at 150 dpi for primary OCR and 300 dpi for age review.
subprocess.run(['pdftoppm', '-f', str(args.first_page), '-l', str(last_page), '-jpeg', '-r', '150', str(pdf_path), 'work/scan/rec'], check=True)
subprocess.run(['pdftoppm', '-f', str(args.first_page), '-l', str(last_page), '-jpeg', '-r', '300', '-jpegopt', 'quality=95', str(pdf_path), 'work/scan300/page'], check=True)

def rendered_page(prefix, page_number):
    # Find the rendered page regardless of whether Poppler zero-pads page numbers.
    prefix_path = Path(prefix)
    pattern = re.compile(re.escape(prefix_path.name) + r'-(\d+)\.jpg$')
    for candidate in prefix_path.parent.glob(prefix_path.name + '-*.jpg'):
        match = pattern.fullmatch(candidate.name)
        if match and int(match.group(1)) == page_number:
            return candidate
    raise FileNotFoundError(
        f'Poppler did not create an image for PDF page {page_number} '
        f'(looked for {prefix_path.name}-{page_number}.jpg with or without zero padding). '
        'Check the selected voter-list page range.'
    )

# Crop the 3-column by 10-row voter cards and run primary Hindi/English OCR.
jobs = []
for p in page_numbers:
    im = Image.open(rendered_page('work/scan/rec', p))
    for r in range(10):
        for c, x in enumerate((55, 885, 1715), start=1):
            y = 118 + r * 344
            out = f'work/cards/p{p:02}_r{r:02}_c{c}.png'
            im.crop((x, y, x + 805, y + 330)).save(out)
            jobs.append(out)
def ocr_primary(f):
    q = subprocess.run(['tesseract', f, 'stdout', '-l', 'hin+eng', '--psm', '6'], capture_output=True, text=True, check=True)
    return f, q.stdout.strip()
with ThreadPoolExecutor(max_workers=8) as ex:
    for f, txt in ex.map(ocr_primary, jobs):
        Path(f.rsplit('.', 1)[0] + '.txt').write_text(txt)

# Alternate text OCR pass for names and relations.
def ocr_alt(f):
    q = subprocess.run(['tesseract', f, 'stdout', '-l', 'hin+eng', '--psm', '11'], capture_output=True, text=True, check=True)
    return f, q.stdout.strip()
with ThreadPoolExecutor(max_workers=8) as ex:
    for f, txt in ex.map(ocr_alt, jobs):
        Path(f.rsplit('.', 1)[0] + '_alt.txt').write_text(txt)

# Two independent English OCR passes over the voter-ID header region.
def ocr_id(f):
    q = subprocess.run(['tesseract', f, 'stdout', '-l', 'eng', '--psm', '11'], capture_output=True, text=True, check=True)
    return f, q.stdout.strip()
with ThreadPoolExecutor(max_workers=8) as ex:
    for f, txt in ex.map(ocr_id, jobs):
        Path(f.rsplit('.', 1)[0] + '_id.txt').write_text(txt)
def ocr_id_roi(f):
    im = Image.open(f)
    roi = im.crop((500, 0, 805, 90)).resize((915, 270))
    temp = f.rsplit('.', 1)[0] + '_roi.png'
    roi.save(temp)
    q = subprocess.run(['tesseract', temp, 'stdout', '-l', 'eng', '--psm', '7'], capture_output=True, text=True, check=True)
    return f, q.stdout.strip()
with ThreadPoolExecutor(max_workers=8) as ex:
    for f, txt in ex.map(ocr_id_roi, jobs):
        Path(f.rsplit('.', 1)[0] + '_roi.txt').write_text(txt)

# High-resolution card crops used to re-read ambiguous/missing ages.
hi_jobs = []
for p in page_numbers:
    im = Image.open(rendered_page('work/scan300/page', p))
    for r in range(10):
        for c, x in enumerate((110, 1770, 3430), start=1):
            y = 236 + r * 688
            out = f'work/cards/p{p:02}_r{r:02}_c{c}_hi.png'
            im.crop((x, y, x + 1610, y + 660)).save(out)
            hi_jobs.append(out)
def ocr_high(f):
    q = subprocess.run(['tesseract', f, 'stdout', '-l', 'hin+eng', '--psm', '6'], capture_output=True, text=True, check=True)
    return f, q.stdout.strip()
with ThreadPoolExecutor(max_workers=8) as ex:
    for f, txt in ex.map(ocr_high, hi_jobs):
        Path(f.rsplit('.', 1)[0] + '.txt').write_text(txt)

# Detect missing/under-18 OCR ages, then use three enlarged numeric-area OCR crops.
suspects = []
for p in page_numbers:
    for r in range(10):
        for c in range(1, 4):
            base = f'work/cards/p{p:02}_r{r:02}_c{c}'
            txt = Path(base + '.txt').read_text(errors='replace')
            m = re.search(r'आयु\s*[:;]?\s*(\d{1,3})', txt)
            if not m or int(m.group(1)) < 18:
                suspects.append((p, r, c, base))
def age_ocr(item):
    p, r, c, base = item
    im = Image.open(base + '_hi.png')
    outs = []
    for i, (x1, x2) in enumerate(((130, 350), (160, 270), (120, 330))):
        crop = im.crop((x1, 390, x2, 480)).resize(((x2-x1)*6, 540))
        f = base + f'_ageocr{i}.png'
        crop.save(f)
        q = subprocess.run(['tesseract', f, 'stdout', '-l', 'eng', '--psm', '7'], capture_output=True, text=True, check=True)
        outs.append(q.stdout.strip().replace('\n', ' '))
    return p, r + 1, c, outs
with ThreadPoolExecutor(max_workers=8) as ex:
    age_results = list(ex.map(age_ocr, suspects))
with open('work/age_suspects.txt', 'w') as f:
    for item in age_results:
        f.write(repr(item) + '\n')

from PIL import Image
from pathlib import Path
import re, json, ast, collections, unicodedata
records=[]
manual_sex={
(4,3,1):'पुरुष',(4,4,1):'पुरुष',(5,2,3):'महिला',(5,3,3):'पुरुष',(7,1,2):'पुरुष',(7,4,3):'पुरुष',(7,8,1):'पुरुष',(8,9,3):'पुरुष',(9,2,3):'महिला',(13,3,1):'महिला',(13,7,2):'पुरुष',(13,8,2):'महिला',(13,9,3):'पुरुष',(14,1,2):'महिला',(15,4,2):'महिला',
(15,5,3):'पुरुष',(15,7,1):'पुरुष',(15,7,2):'महिला',(18,1,2):'पुरुष',(18,1,3):'महिला',(18,7,3):'पुरुष',(19,2,3):'पुरुष',(19,3,1):'महिला',(21,7,2):'महिला',(22,4,1):'पुरुष',(23,4,2):'पुरुष',(23,6,1):'पुरुष',(24,4,1):'महिला',(24,8,1):'महिला',(26,9,2):'महिला',
(27,3,2):'महिला',(27,9,3):'महिला',(28,6,3):'महिला',(30,5,1):'पुरुष',(30,5,2):'महिला',(30,7,3):'महिला',(30,9,3):'महिला',(31,9,2):'महिला',(33,1,3):'पुरुष',(33,9,3):'महिला',(34,4,3):'महिला',(34,6,2):'पुरुष',(36,2,1):'महिला',(37,4,3):'पुरुष',(37,9,1):'पुरुष',(38,7,2):'महिला',(38,8,1):'महिला'
}
manual_age={(6,6,2):21,(11,1,3):31,(14,6,3):31,(17,3,3):19,(17,6,3):31,(22,5,3):71,(28,5,1):31,(30,6,3):21,(34,6,3):31,(34,10,3):19,(35,9,3):51,(42,2,3):41}
if pdf_path.name.lower() != '279.pdf':
 manual_sex = {}
 manual_age = {}
# OCR candidate corrections for age rows where the leading digit was dropped.
ocr_age=[ast.literal_eval(s) for s in open('work/age_suspects.txt')]
agefix={}
for p,r,c,outs in ocr_age:
 candidates=[]
 for s in outs:
  candidates.extend(int(x) for x in re.findall(r'(?<!\d)\d{2,3}(?!\d)',s) if 18<=int(x)<=110)
 co=collections.Counter(candidates)
 if co:
  age=co.most_common(1)[0][0]
  # reject a tied first place for visual review
  if len(co)>1 and co.most_common()[0][1]==co.most_common()[1][1]: continue
  agefix[(p,r,c)]=age

def tidy(s):
 s=s.strip().strip(' :;|.!।,')
 s=re.sub(r'\s+',' ',s)
 s=re.sub(r'\s*(?:८०|८०2|७०|९\)|«८०2|८०\s*)$','',s).strip()
 s=re.sub(r'\s*(?:फोटो उपलब्ध.*|उपलब्ध है.*)$','',s).strip()
 return s.strip(' :;|.!।,')
name_label=re.compile(r'^(?:नाम|माम|नाम॑|नामः|am|ant|are|aa|art|ar|ara)\s*[:;ः]\s*(.+)$',re.I)
rel_label=re.compile(r'(?:पिता|पति|माता|पित्ता|प्रिता|प्रति|पितृ|पिला|घति|पत्ति|fret|Ufa|Ufa|Ufa).{0,24}?(?:नाम|ara|am|ATA|ATH|art|AT)\s*[:;ः!]?\s*(.+)$',re.I)
for p in page_numbers:
 for i in range(30):
  r,c=divmod(i,3);r+=1;c+=1;base=Path(f'work/cards/p{p:02}_r{r-1:02}_c{c}')
  t=Path(str(base)+'.txt').read_text(errors='replace'); alt=Path(str(base)+'_alt.txt').read_text(errors='replace')
  # voter IDs are captured from the distinct top-right header in two independent OCR passes.
  idtxt=Path(str(base)+'_id.txt').read_text(errors='replace')+' '+Path(str(base)+'_roi.txt').read_text(errors='replace')
  idtxt=re.sub(r'\s+','',idtxt)
  ids=re.findall(r'(?i)[A-Z]{3}[A-Z0-9]{7}',idtxt)
  voter_id=ids[0].upper() if ids else ''
  # Prefer exact OCR labels, then recover the one name line from common label variants.
  name=''
  for tx in (t,alt):
   lines=tx.splitlines()
   for j,line in enumerate(lines):
    if re.search(r'(?:पिता|पति|माता)\s*का',line): continue
    if re.search(r'(मकान|आयु|लिंग|फोटो|उपलब्ध)',line): continue
    if re.search(r'(?i)[A-Z]{3}[A-Z0-9]{5,}',line): continue
    m=re.search(r'(?:नाम|माम|नाम॑|नामः|am|ant|are|aa|art|ar|ara|amt|ary)\s*[:;ः!]\s*(.+)',line,re.I)
    if not m: m=re.match(r'^\s*[^:;ः!]{0,16}[:;ः!]\s*(.+)$',line)
    if m:
     name=tidy(m.group(1));break
   if name:break
  relation=''
  for tx in (t,alt):
   for line in tx.splitlines():
    if not re.search(r'(पिता|पति|माता|पित्ता|प्रिता|प्रति|पितृ|पिला|घति|पत्ति|fret|Ufa|पिता)',line,re.I):continue
    if re.search(r'(मकान|आयु|लिंग|फोटो|उपलब्ध)',line): continue
    m=rel_label.search(line)
    if not m:
     m=re.search(r'(?:पिता|पति|माता|पित्ता|प्रिता|प्रति|पितृ|पिला|घति|पत्ति|fret|Ufa)[^:;ः!]{0,32}[:;ः!]\s*(.+)$',line,re.I)
    if not m:
     m=re.search(r'(?:का\s*(?:नाम|ara|am|ATH|AT|bing)\s*)[:;ः]?\s*(.+)$',line,re.I)
    if m:
     relation=tidy(m.group(1));break
   if relation:break
  # Age: patch dropped leading digits using a second OCR of the isolated numeric area.
  m=re.search(r'आयु\s*[:;]?\s*(\d{1,3})',t) or re.search(r'आयु\s*[:;]?\s*(\d{1,3})',alt)
  age=int(m.group(1)) if m else None
  if (p,r,c) in manual_age:age=manual_age[(p,r,c)]
  elif (p,r,c) in agefix:age=agefix[(p,r,c)]
  # Sex: first-pass/alternate OCR; manually read the 47 low-confidence sex labels from the source cards.
  sex='महिला' if ('महिला' in t or 'महिला' in alt) else ('पुरुष' if ('पुरुष' in t or 'पुरुष' in alt) else '')
  if (p,r,c) in manual_sex: sex=manual_sex[(p,r,c)]
  records.append({'voter id':voter_id,'Name':name,"Father's name":relation,'age':age,'sex':sex,'page':p,'row':r,'col':c,'raw':t,'alt':alt})
# Ignore empty card positions so the workbook size follows the actual extracted records.
records = [x for x in records if any((x['voter id'], x['Name'], x["Father's name"], x['age'] is not None, x['sex']))]
unique_ids = len({x['voter id'] for x in records if x['voter id']})
missing_ids = sum(not x['voter id'] for x in records)
print('records',len(records),'unique IDs',unique_ids,'ID missing',missing_ids,'name missing',sum(not x['Name'] for x in records),'relation missing',sum(not x["Father's name"] for x in records),'age missing',sum(x['age'] is None for x in records),'sex missing',sum(not x['sex'] for x in records))
if args.expected_records is not None:
 if len(records) == args.expected_records:
  print(f'Expected record count matched: {args.expected_records}')
 else:
  print(f'WARNING: expected {args.expected_records} records, extracted {len(records)}. Review page range/layout and OCR output.')
if missing_ids or unique_ids != len(records):
 print('WARNING: voter IDs are missing or duplicated; review the workbook against the PDF.')
for k,v in collections.Counter(x['age'] for x in records).most_common(5): pass
print('under18',[(x['voter id'],x['age'],x['page'],x['row'],x['col']) for x in records if x['age'] is not None and x['age']<18])
with open('work/records.json','w') as f:json.dump(records,f,ensure_ascii=False,indent=1)

# Write the requested Excel workbook using Python/openpyxl.
wb = Workbook()
ws = wb.active
ws.title = 'Voter Records'
headers = ['Voter ID', 'Name', "Father's name", 'Age', 'Sex']
ws.append(headers)
for rec in records:
    ws.append([
        rec['voter id'], rec['Name'], rec["Father's name"], rec['age'], rec['sex']
    ])
ws.freeze_panes = 'A2'
ws.auto_filter.ref = f'A1:E{ws.max_row}'
for cell in ws[1]:
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='1F4E78')
    cell.alignment = Alignment(horizontal='center')
for col, width in {'A': 18, 'B': 34, 'C': 34, 'D': 10, 'E': 14}.items():
    ws.column_dimensions[col].width = width
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(vertical='top', wrap_text=True)
if ws.max_row > 1:
    table = Table(displayName='VoterRecords', ref=f'A1:E{ws.max_row}')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
wb.save(xlsx_path)
print(f'Saved {len(records)} voter records to {xlsx_path}')
