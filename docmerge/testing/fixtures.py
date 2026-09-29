"""Offline testu fixtures: dokumentu ģenerēšana, reāli Word fixtures, pārbaudes.

Minimālie OOXML DOCX (`write_minimal_docx`) tiek izmantoti simtiem avotu, jo
Word COM `Range.InsertFile` tos pieņem (pārbaudīts ar `_diag_minimal.py`).
Reāliem E2E testiem izmanto `WordFixtureMaker` (Word COM).
"""
from __future__ import annotations

import ctypes
import gc
import hashlib
import os
import shutil
import stat
import subprocess
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from docmerge.core.ooxml import write_minimal_docx

W_NS='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ZONE_STREAM_SUFFIX=':Zone.Identifier'
ZONE_CONTENT='[ZoneTransfer]\r\nZoneId=3\r\n'
WORD_PROCESS_NAME='WINWORD.EXE'


# ------------------------------------------------------------------ files / dirs
def _remove_readonly(func,path,exc_info):
    """shutil.rmtree kļūdu apstrāde: noņem Read-only un mēģina vēlreiz."""
    try:
        os.chmod(path,stat.S_IWRITE); func(path)
    except OSError: pass


def clean_dir(path):
    """Izdzēš un izveido mapi (droši arī read-only failiem un ja ceļš ir fails)."""
    target=Path(path)
    if target.is_dir(): shutil.rmtree(target,onerror=_remove_readonly)
    elif target.exists():
        try: target.unlink()
        except OSError:
            try: os.chmod(target,stat.S_IWRITE); target.unlink()
            except OSError: pass
    target.mkdir(parents=True,exist_ok=True)
    return target


def sha256(path,chunk_size=1024*1024):
    digest=hashlib.sha256()
    with open(path,'rb') as fh:
        for chunk in iter(lambda:fh.read(chunk_size),b''): digest.update(chunk)
    return digest.hexdigest()


def hashes_of(paths):
    return {str(Path(p)):sha256(p) for p in paths}


# ------------------------------------------------------------------ docx helpers
def make_docx(path,text,paragraphs=None):
    """Izveido derīgu .docx ar minimālu OOXML (bez Word)."""
    body=list(paragraphs) if paragraphs else ([text] if text else [])
    return write_minimal_docx(path,body)


def docx_xml(path):
    with zipfile.ZipFile(path) as z: return z.read('word/document.xml').decode('utf-8')


def docx_texts(path):
    """Visi teksta gabali (w:t) dokumenta secībā."""
    root=ET.fromstring(docx_xml(path))
    return [node.text or '' for node in root.iter(f'{W_NS}t')]


def docx_text(path):
    return '\n'.join(docx_texts(path))


def page_break_count(path):
    return docx_xml(path).count('w:type="page"')


def is_valid_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            names=set(z.namelist())
            return '[Content_Types].xml' in names and 'word/document.xml' in names
    except (OSError,zipfile.BadZipFile): return False


def paragraphs_of(path):
    """w:p elementu tekstu saraksts (tukšie izlaisti)."""
    root=ET.fromstring(docx_xml(path))
    out=[]
    for paragraph in root.iter(f'{W_NS}p'):
        text=''.join(node.text or '' for node in paragraph.iter(f'{W_NS}t'))
        if text.strip(): out.append(text)
    return out


# ------------------------------------------------------------------ zone / attrs
def write_zone_identifier(path,content=ZONE_CONTENT):
    with open(f'{path}{ZONE_STREAM_SUFFIX}','w',encoding='utf-8') as fh: fh.write(content)
    return str(path)


def has_zone_identifier(path):
    if os.name!='nt': return False
    return os.path.exists(f'{path}{ZONE_STREAM_SUFFIX}')


def make_read_only(path):
    os.chmod(path,os.stat(path).st_mode&~stat.S_IWRITE)
    return path


def make_writable(path):
    os.chmod(path,os.stat(path).st_mode|stat.S_IWRITE)
    return path


# ------------------------------------------------------------------ process info
def winword_pids():
    if os.name!='nt': return set()
    try:
        cp=subprocess.run(['tasklist','/FI',f'IMAGENAME eq {WORD_PROCESS_NAME}','/NH'],
                          capture_output=True,text=True,timeout=20)
    except (OSError,subprocess.SubprocessError): return set()
    pids=set()
    for line in (cp.stdout or '').splitlines():
        parts=line.split()
        if len(parts)>=2 and parts[0].upper()==WORD_PROCESS_NAME:
            try: pids.add(int(parts[1].replace(',','')))
            except ValueError: pass
    return pids


def process_memory_mb():
    """Šī procesa RSS (MB). None, ja noteikt nevar."""
    if os.name!='nt': return None
    try:
        class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
            _fields_=[('cb',ctypes.c_ulong),('PageFaultCount',ctypes.c_ulong),
                      ('PeakWorkingSetSize',ctypes.c_size_t),('WorkingSetSize',ctypes.c_size_t),
                      ('QuotaPeakPagedPoolUsage',ctypes.c_size_t),('QuotaPagedPoolUsage',ctypes.c_size_t),
                      ('QuotaPeakNonPagedPoolUsage',ctypes.c_size_t),('QuotaNonPagedPoolUsage',ctypes.c_size_t),
                      ('PagefileUsage',ctypes.c_size_t),('PeakPagefileUsage',ctypes.c_size_t)]
        counters=PROCESS_MEMORY_COUNTERS(); counters.cb=ctypes.sizeof(PROCESS_MEMORY_COUNTERS)
        ok=ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.windll.kernel32.GetCurrentProcess(),
                                                    ctypes.byref(counters),counters.cb)
        if not ok: return None
        return round(counters.WorkingSetSize/1024/1024,1)
    except Exception:  # noqa: BLE001 - bez ctypes API atgriežam None
        return None


def object_count():
    return len(gc.get_objects())


def leftover_files(root,patterns=('*.tmp','*.tmp.docx')):
    """Atlikušie temp faili mapē (rekursīvi)."""
    base=Path(root); found=[]
    if not base.exists(): return found
    for pattern in patterns: found.extend(str(p) for p in base.rglob(pattern))
    return sorted(found)


def can_open_for_write(path):
    """True, ja failu var atvērt rakstīšanai (nav aizņemts)."""
    try:
        with open(path,'r+b'): return True
    except OSError: return False


# ------------------------------------------------------------- real Word fixtures
class WordFixtureMaker:
    """Veido reālus Word dokumentus (DOCX/DOC) ar Word COM; droši aizver sesiju.

    Izmanto tikai lokālo Word (OFFLINE atļauts) un nekad neaizver lietotāja
    sesiju, ja pieslēgšanās notika ar `Dispatch` fallback.
    """

    WD_FORMAT_DOCUMENT=0
    WD_FORMAT_DOCX=16
    WD_FORMAT_TEXT=2

    def __init__(self,log=None):
        self.log=log; self.session=None; self.created=[]; self.add_attempts=0

    def __enter__(self):
        import pythoncom
        from docmerge.engines.word_com import open_word_session
        pythoncom.CoInitialize()
        self._pythoncom=pythoncom
        self.session=open_word_session(self._pythoncom_client(),log=self.log)
        return self

    def __exit__(self,exc_type,exc,tb):
        try:
            if self.session is not None: self.session.close(self.log)
        except Exception: pass
        try: self._pythoncom.CoUninitialize()
        except Exception: pass
        gc.collect()
        return False

    @staticmethod
    def _pythoncom_client():
        import win32com.client
        return win32com.client

    @property
    def word(self): return self.session.word

    @property
    def owned(self): return bool(self.session and self.session.owned)

    def make(self,path,text,file_format=None):
        """Izveido dokumentu ar doto tekstu. Atgriež Path.

        Word COM relatīvos ceļus risina pret SAVU darba mapi
        (`C:\\WINDOWS\\system32`), tāpēc `SaveAs2` saņem ABSOLŪTU ceļu.
        """
        target=Path(path); target.parent.mkdir(parents=True,exist_ok=True)
        absolute=str(target.resolve())
        fmt=file_format if file_format is not None else (self.WD_FORMAT_DOCX if target.suffix.lower()=='.docx' else self.WD_FORMAT_DOCUMENT)
        doc=None; last=None
        for attempt in range(1,5):
            try:
                doc=self.word.Documents.Add(); break
            except Exception as exc:  # noqa: BLE001 - WRD6ER32.CNV var būt vienreizēja kļūda
                last=exc; self.add_attempts+=1
        if doc is None: raise last
        try:
            doc.Content.Text=text or ''
            doc.SaveAs2(absolute,FileFormat=fmt)
        finally:
            try: doc.Close(False)
            except Exception: pass
        self.created.append(target)
        return target

    def open_document(self,path,read_only=True):
        """Atver dokumentu un atstāj to atvērtu (atgriež Word dokumentu objektu)."""
        return self.word.Documents.Open(str(Path(path).resolve()),ReadOnly=read_only,
                                        AddToRecentFiles=False,ConfirmConversions=False)
