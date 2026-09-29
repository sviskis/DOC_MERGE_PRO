from pathlib import Path
import zipfile
OLE_MAGIC=bytes.fromhex('D0CF11E0A1B11AE1')

def detect_format(path):
    p=Path(path)
    if not p.is_file(): return 'MISSING'
    with p.open('rb') as f: head=f.read(16)
    if head.startswith(OLE_MAGIC): return 'OLE'
    if head.startswith(b'PK'):
        try:
            with zipfile.ZipFile(p) as z:
                n=set(z.namelist())
                if '[Content_Types].xml' in n and 'word/document.xml' in n: return 'DOCX_PACKAGE'
                if 'mimetype' in n and 'content.xml' in n:
                    try: mime=z.read('mimetype').decode('utf-8',errors='replace')
                    except Exception: mime=''
                    if 'opendocument' in mime: return 'ODT_PACKAGE'
                return 'ZIP_UNKNOWN'
        except zipfile.BadZipFile: return 'ZIP_CORRUPT'
    if p.suffix.lower()=='.rtf':
        try:
            if p.read_bytes()[:16].lstrip().startswith(b'{\\rtf'): return 'RTF'
        except OSError: pass
    if p.suffix.lower()=='.txt': return 'TXT'
    return 'UNKNOWN'
