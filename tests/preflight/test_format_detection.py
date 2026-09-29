import zipfile
from docmerge.core.format_detection import detect_format

def test_txt(tmp_path):
    p=tmp_path/'a.txt'; p.write_text('Sveiki',encoding='utf-8'); assert detect_format(str(p))=='TXT'

def test_docx(tmp_path):
    p=tmp_path/'a.docx'
    with zipfile.ZipFile(p,'w') as z:z.writestr('[Content_Types].xml','<Types/>'); z.writestr('word/document.xml',"<w:document xmlns:w='x'/>")
    assert detect_format(str(p))=='DOCX_PACKAGE'
