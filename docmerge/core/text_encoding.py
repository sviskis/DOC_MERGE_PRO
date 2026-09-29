from pathlib import Path
CANDIDATES=('utf-8-sig','utf-8','cp1257','cp1252'); BAD=('�','Ã','Ä','Å')
def detect_text_encoding(path):
    raw=Path(path).read_bytes()
    if raw.startswith(b'\xef\xbb\xbf'): return 'utf-8-sig',1.0,[],raw.decode('utf-8-sig')
    best=None
    for enc in CANDIDATES[1:]:
        try: text=raw.decode(enc)
        except UnicodeDecodeError: continue
        bad=sum(text.count(m) for m in BAD); lv=sum(text.count(c) for c in 'āčēģīķļņšūžĀČĒĢĪĶĻŅŠŪŽ'); score=1-min(bad*.1,.8)+min(lv/max(len(text),1),.05)
        row=(score,enc,text,bad)
        if best is None or row[0]>best[0]: best=row
    if best is None: return None,0.0,['Neizdevās noteikt kodējumu'],''
    score,enc,text,bad=best; warns=[]
    if bad: warns.append(f'Iespējamas mojibake pazīmes: {bad}')
    return enc,max(0,min(score,1)),warns,text
