"""Diagnostika: vai Word COM `Range.InsertFile` pieņem minimālos OOXML DOCX?

Ja jā, offline testi var ātri ģenerēt simtiem avotu bez Word (`write_minimal_docx`)
un tomēr testēt reālu Word apvienošanu.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from docmerge.core.ooxml import write_minimal_docx  # noqa: E402

WORK = ROOT / 'runtime' / '_min'
if WORK.exists():
    shutil.rmtree(WORK, ignore_errors=True)
SRC = WORK / 'source'
SRC.mkdir(parents=True, exist_ok=True)

write_minimal_docx(SRC / 'A_min.docx', ('MARKER_ALFA minimālais OOXML',))
write_minimal_docx(SRC / 'B_min.docx', ('MARKER_BETA minimālais OOXML',))
OUT = WORK / 'MIN_APVIENOTS.docx'

cp = subprocess.run([sys.executable, '-m', 'docmerge.cli.main', '--merge-all', str(SRC), '--output', str(OUT)],
                    capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=str(ROOT))
print('rc', cp.returncode)
print((cp.stdout or '')[-1500:])
print((cp.stderr or '')[-1500:])

if OUT.exists():
    with zipfile.ZipFile(OUT) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    print('ALFA:', 'MARKER_ALFA' in xml)
    print('BETA:', 'MARKER_BETA' in xml)
    print('page breaks:', xml.count('w:type="page"'))
    print('titles:', [t for t in ('A_min', 'B_min') if f'>{t}<' in xml])
else:
    print('OUTPUT MISSING')
result = ROOT / 'reports' / 'result.json'
if result.exists():
    data = json.loads(result.read_text(encoding='utf-8'))
    print('result:', json.dumps({k: data.get(k) for k in ('status', 'total', 'merged', 'errors')}, ensure_ascii=False))
