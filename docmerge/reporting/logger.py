from pathlib import Path
from datetime import datetime
import logging

def create_logger(log_dir='logs'):
    Path(log_dir).mkdir(parents=True,exist_ok=True); path=Path(log_dir)/f'docmerge_{datetime.now():%Y%m%d_%H%M%S}.log'; logger=logging.getLogger(path.stem); logger.setLevel(logging.INFO); logger.propagate=False
    h=logging.FileHandler(path,encoding='utf-8'); h.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')); logger.addHandler(h); return logger,path
