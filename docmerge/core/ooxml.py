"""Minimāla OOXML (.docx) ģenerēšana bez python-docx.

Izmanto:
- tukša master dokumenta izveidei, ja Word `Documents.Add()` nokrīt;
- testu avota dokumentu izveidei (offline, bez Word).

NEIZMANTO kā production merge engine — production merge paliek Word COM
`Range.InsertFile`.
"""
import zipfile
from pathlib import Path

CONTENT_TYPES='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''

ROOT_RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''

DOCUMENT_TEMPLATE='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>'''


def xml_escape(text):
    return (str(text).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;'))


def paragraph_xml(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:p>'


def document_xml(paragraphs):
    body=''.join(paragraph_xml(p) for p in paragraphs) or '<w:p/>'
    return DOCUMENT_TEMPLATE.format(body=body)


def write_minimal_docx(path,paragraphs=()):
    """Ieraksta derīgu .docx pakotni. Atgriež Path."""
    target=Path(path); target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml',CONTENT_TYPES)
        zf.writestr('_rels/.rels',ROOT_RELS)
        zf.writestr('word/document.xml',document_xml(paragraphs))
    return target
