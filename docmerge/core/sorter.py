import re
from pathlib import Path
from docmerge.domain.enums import SortMode

def natural_key(v): return [int(x) if x.isdigit() else x for x in re.split(r'(\d+)',v.casefold())]

def sort_items(items,mode):
    data=list(items)
    if mode==SortMode.NATURAL_NAME: data.sort(key=lambda x:natural_key(x.filename or Path(x.source_path).name))
    elif mode==SortMode.MODIFIED_DATE_ASC: data.sort(key=lambda x:x.modified_time)
    elif mode==SortMode.MODIFIED_DATE_DESC: data.sort(key=lambda x:x.modified_time,reverse=True)
    # ORDER_LIST seība nāk no CSV/JSON un ir ABSOLŪTA — natural sort netiek pielietots.
    elif mode in (SortMode.MANUAL,SortMode.ORDER_LIST): data.sort(key=lambda x:x.manual_order)
    for i,x in enumerate(data,1): x.manual_order=i
    return data
