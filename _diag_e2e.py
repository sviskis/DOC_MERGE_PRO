"""E2E pārbaude: visi DOC/DOCX -> viens DOCX ar reālu Microsoft Word.

Izveido paraugus (.docx, vecais .doc, ignorējamu .txt), uzliek vienam avotam
īstu Zone.Identifier marķējumu, palaiž CLI `--merge-all` un pārbauda izvadi.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORK = ROOT / 'runtime' / '_e2e'
SRC = WORK / 'source'
OUT = WORK / 'APVIENOTS.docx'
WD_FORMAT_DOCUMENT = 0
WD_FORMAT_DOCX = 16

if WORK.exists():
    shutil.rmtree(WORK, ignore_errors=True)
SRC.mkdir(parents=True, exist_ok=True)

import pythoncom  # noqa: E402
import win32com  # noqa: E402
import win32com.client  # noqa: E402

from docmerge.engines.word_com import dispatch_word  # noqa: E402

pythoncom.CoInitialize()
word = dispatch_word(win32com)


def make(path, text, file_format):
    last = None
    for attempt in range(1, 5):
        try:
            doc = word.Documents.Add()
            break
        except Exception as exc:  # noqa: BLE001 - WRD6ER32.CNV ir vienreizēja kļūda
            last = exc
            print(f'  Documents.Add() {attempt}. mēģinājums: {exc}')
    else:
        raise last
    doc.Content.Text = text
    doc.SaveAs2(str(path), FileFormat=file_format)
    doc.Close(False)
    print('radīts:', path.name, os.path.getsize(path), 'B')


try:
    make(SRC / '01_alpha.docx', 'ALFA markieris 1', WD_FORMAT_DOCX)
    make(SRC / '02_beta.doc', 'BETA markieris 2', WD_FORMAT_DOCUMENT)
    make(SRC / '03_gamma.docx', 'GAMMA markieris 3', WD_FORMAT_DOCX)
    (SRC / '04_ignore.txt').write_text('TXT nav jāapvieno', encoding='utf-8')
    with open(f'{SRC / "03_gamma.docx"}:Zone.Identifier', 'w', encoding='utf-8') as fh:
        fh.write('[ZoneTransfer]\r\nZoneId=3\r\n')
    print('Zone.Identifier uz 03_gamma.docx: True')
finally:
    word.Quit()
    pythoncom.CoUninitialize()

cp = subprocess.run(
    [sys.executable, '-m', 'docmerge.cli.main', '--merge-all', str(SRC), '--output', str(OUT)],
    capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=str(ROOT),
)
print('--- CLI rc', cp.returncode)
print((cp.stdout or '')[-2500:])
print((cp.stderr or '')[-2500:])

with zipfile.ZipFile(OUT) as z:
    xml = z.read('word/document.xml').decode('utf-8')
print('--- IZVADE')
print('apvienotais fails eksistē:', OUT.exists(), os.path.getsize(OUT), 'B')
for marker in ('ALFA markieris 1', 'BETA markieris 2', 'GAMMA markieris 3'):
    print(f'  satur "{marker}":', marker in xml)
print('  .txt saturs nav iekšā:', 'TXT nav' not in xml)
print('  virsraksti (failu nosaukumi):', [t for t in ('01_alpha', '02_beta', '03_gamma') if f'>{t}<' in xml])
print('  lappuses pārtraukumi:', xml.count('w:type="page"'))
print('  Zone.Identifier uz izvades:', os.path.exists(f'{OUT}:Zone.Identifier'))
result_path = ROOT / 'reports' / 'result.json'
if result_path.exists():
    data = json.loads(result_path.read_text(encoding='utf-8'))
    print('--- result.json:', json.dumps({k: data.get(k) for k in ('status', 'total', 'merged', 'skipped', 'errors', 'unlocked', 'normalized')}, ensure_ascii=False))
