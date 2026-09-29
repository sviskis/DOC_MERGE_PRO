from pathlib import Path
import zipfile, xml.etree.ElementTree as ET
from docmerge.domain.enums import DocumentStatus
from docmerge.core.format_detection import detect_format
from docmerge.core.hashing import sha256_file
from docmerge.core.paths import describe_path_risks
from docmerge.core.text_encoding import detect_text_encoding
from docmerge.core.unlock import has_zone_identifier

WORD_OWNER_SIGNATURE=b'\x15Microsoft'


def _is_word_owner_file(path):
    """Word bloķēšanas (~$ / owner) fails — nav dokuments, Word to nevar atvērt."""
    try:
        with Path(path).open('rb') as fh: head=fh.read(64)
    except OSError:
        return False
    return head.startswith(WORD_OWNER_SIGNATURE)

def _docx(path):
    e=[]
    try:
        with zipfile.ZipFile(path) as z:
            bad=z.testzip(); n=set(z.namelist())
            if bad: e.append(f'ZIP kļūda: {bad}')
            for req in ('[Content_Types].xml','word/document.xml'):
                if req not in n: e.append(f'Trūkst {req}')
            if 'word/document.xml' in n:
                try: ET.fromstring(z.read('word/document.xml'))
                except ET.ParseError as ex: e.append(f'XML kļūda: {ex}')
    except zipfile.BadZipFile: e.append('Bojāts DOCX ZIP konteiners')
    return e

def _odt(path):
    e=[]
    try:
        with zipfile.ZipFile(path) as z:
            bad=z.testzip(); n=set(z.namelist())
            if bad: e.append(f'ZIP kļūda: {bad}')
            for req in ('mimetype','content.xml','META-INF/manifest.xml'):
                if req not in n: e.append(f'Trūkst {req}')
            if 'content.xml' in n:
                try: ET.fromstring(z.read('content.xml'))
                except ET.ParseError as ex: e.append(f'ODT XML kļūda: {ex}')
    except zipfile.BadZipFile: e.append('Bojāts ODT ZIP konteiners')
    return e

def preflight_item(item,calculate_hash=True):
    item.health_status=DocumentStatus.CHECKING; item.warnings.clear(); item.errors.clear(); item.hydrate_from_path(); p=Path(item.source_path)
    if not p.exists(): item.errors.append('Fails neeksistē')
    elif p.name.startswith('~$'): item.errors.append('Word pagaidu fails')
    elif item.size_bytes==0: item.errors.append('Fails ir tukšs (0 B)')
    elif not p.stem: item.errors.append('Nederīgs faila nosaukums (nav pamatnosaukuma) — Word to nevar atvērt')
    elif _is_word_owner_file(item.source_path): item.errors.append('Word bloķēšanas (owner) fails, nevis dokuments')
    if item.errors: item.health_status=DocumentStatus.ERROR; return item
    item.warnings.extend(describe_path_risks(item.source_path))
    if has_zone_identifier(item.source_path):
        item.warnings.append('Protected View marķējums (Zone.Identifier) — apvienošanai tiks izmantota atbloķēta kopija')
    item.detected_format=detect_format(item.source_path)
    expected={'.docx':{'DOCX_PACKAGE'},'.docm':{'DOCX_PACKAGE'},'.dotx':{'DOCX_PACKAGE'},'.dotm':{'DOCX_PACKAGE'},'.doc':{'OLE'},'.dot':{'OLE'},'.odt':{'ODT_PACKAGE'},'.rtf':{'RTF'},'.txt':{'TXT'}}.get(item.extension,set())
    if expected and item.detected_format not in expected: item.warnings.append(f'Paplašinājums {item.extension} neatbilst formātam {item.detected_format}')
    if item.detected_format=='DOCX_PACKAGE': item.errors.extend(_docx(item.source_path))
    elif item.detected_format=='ODT_PACKAGE': item.errors.extend(_odt(item.source_path))
    elif item.detected_format=='ZIP_CORRUPT': item.errors.append('Bojāts ZIP konteiners')
    if item.extension=='.txt':
        enc,conf,warns,_=detect_text_encoding(item.source_path); item.encoding=enc; item.encoding_confidence=conf; item.warnings.extend(warns)
    if calculate_hash and not item.errors:
        try: item.binary_hash=sha256_file(item.source_path)
        except OSError as ex: item.warnings.append(f'Hash kļūda: {ex}')
    item.health_status=DocumentStatus.ERROR if item.errors else (DocumentStatus.WARNING if item.warnings else DocumentStatus.OK)
    return item

def preflight_all(items): return [preflight_item(x) for x in items]
