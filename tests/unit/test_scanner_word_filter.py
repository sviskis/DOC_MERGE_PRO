from pathlib import Path

from docmerge.core.scanner import (WORD_DOCUMENT_EXTENSIONS, filter_word_documents,
                                   is_word_document, scan_folder)


def test_word_document_extension_set():
    assert WORD_DOCUMENT_EXTENSIONS=={'.doc','.docx'}


def test_is_word_document_accepts_doc_and_docx_only():
    assert is_word_document('C:/docs/vecais.doc') is True
    assert is_word_document('C:/docs/jaunais.DOCX') is True
    assert is_word_document('C:/docs/teikums.txt') is False
    assert is_word_document('C:/docs/template.dot') is False
    assert is_word_document('C:/docs/~$slēpts.docx') is False


def test_filter_keeps_only_doc_docx_items(tmp_path):
    (tmp_path/'a.docx').write_bytes(b'PK\x03\x04')
    (tmp_path/'b.doc').write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
    (tmp_path/'c.txt').write_text('x',encoding='utf-8')
    (tmp_path/'d.rtf').write_bytes(b'{\\rtf1}')
    (tmp_path/'~$e.docx').write_bytes(b'x')  # skeneris jau izlaiž Word ~$ failus
    items=scan_folder(str(tmp_path),recursive=True)
    assert len(items)==4
    kept=filter_word_documents(items)
    assert sorted(Path(x.source_path).name for x in kept)==['a.docx','b.doc']
