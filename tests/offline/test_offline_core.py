from docmerge.core.text_encoding import detect_text_encoding
from docmerge.core.hashing import sha256_file
from docmerge.core.sorter import natural_key

def test_offline(tmp_path):
    p=tmp_path/'lv.txt'; p.write_text('āčēģīķļņšūž',encoding='utf-8'); enc,conf,warn,text=detect_text_encoding(str(p)); assert enc in {'utf-8','utf-8-sig'}; assert conf>0; assert len(sha256_file(str(p)))==64; assert sorted(['10','2','1'],key=natural_key)==['1','2','10']
