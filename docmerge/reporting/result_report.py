from pathlib import Path
from datetime import datetime
import json

def write_result(path,**kw):
    data={'timestamp':datetime.now().isoformat(timespec='seconds'),**kw}; p=Path(path); p.parent.mkdir(parents=True,exist_ok=True); p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8'); return data
