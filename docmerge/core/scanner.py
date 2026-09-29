from pathlib import Path
from docmerge.domain.models import DocumentItem
SUPPORTED_EXTENSIONS={'.doc','.docx','.docm','.dot','.dotx','.dotm','.rtf','.odt','.txt'}
# Word dokumenti, ko apvieno "visi DOC/DOCX -> viens DOCX" režīmā.
WORD_DOCUMENT_EXTENSIONS={'.doc','.docx'}


def is_word_document(path):
    """True, ja ceļam ir .doc vai .docx paplašinājums un tas nav Word ~$ fails."""
    p=Path(path)
    return p.suffix.lower() in WORD_DOCUMENT_EXTENSIONS and not p.name.startswith('~$')


def filter_word_documents(items):
    """Atstāj tikai .doc/.docx vienumus (DOC/DOCX -> viens DOCX apvienošanai)."""
    return [x for x in items if is_word_document(x.source_path or x.filename or '')]

def scan_folder(folder:str, recursive:bool=False):
    root=Path(folder)
    if not root.is_dir(): raise ValueError(f'Mape neeksistē: {folder}')
    it=root.rglob('*') if recursive else root.glob('*'); out=[]
    for p in it:
        if not p.is_file() or p.name.startswith('~$') or p.suffix.lower() not in SUPPORTED_EXTENSIONS: continue
        x=DocumentItem(str(p)); x.hydrate_from_path(); x.manual_order=len(out)+1; out.append(x)
    return out

def add_files(paths):
    out=[]
    for raw in paths:
        p=Path(raw)
        if not p.is_file() or p.name.startswith('~$') or p.suffix.lower() not in SUPPORTED_EXTENSIONS: continue
        x=DocumentItem(str(p)); x.hydrate_from_path(); out.append(x)
    return out
