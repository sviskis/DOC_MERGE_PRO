from pathlib import Path
import json
from docmerge.domain.models import Project

def save_project(project,path):
    p=Path(path); p.parent.mkdir(parents=True,exist_ok=True); p.write_text(json.dumps(project.to_dict(),ensure_ascii=False,indent=2),encoding='utf-8')

def load_project(path): return Project.from_dict(json.loads(Path(path).read_text(encoding='utf-8')))
