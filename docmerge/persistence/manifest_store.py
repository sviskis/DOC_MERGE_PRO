"""Merge manifest: source failu relative path, izmērs, mtime un SHA-256.

Manifesta salīdzinājums (`compare_manifest`) ļauj pēc atkārtota merge ar to pašu
order sarakstu parādīt katram avotam statusu: UNCHANGED / CHANGED / MISSING
(+ ADDED jauniem failiem). Tas ir informatīvs un merge NEbloķē (arī STRICT režīmā
bloķē tikai order saraksta MISSING/AMBIGUOUS/DUPLICATE ieraksti).
"""
from pathlib import Path
from datetime import datetime
import json

from docmerge.core.order_list import common_base, relative_for

DIFF_UNCHANGED='UNCHANGED'; DIFF_CHANGED='CHANGED'; DIFF_MISSING='MISSING'; DIFF_ADDED='ADDED'


def build_manifest(project,items):
    """Pilns manifests: katram avotam relative path, size, mtime, sha256."""
    base=common_base(items)
    documents=[]
    for i,item in enumerate(items,1):
        documents.append({'index':i,'id':item.id,'source':item.source_path,
                          'relative_path':relative_for(item.source_path,base),
                          'size_bytes':int(item.size_bytes or 0),'modified_time':float(item.modified_time or 0.0),
                          'sha256':item.binary_hash,
                          'status':'READY' if item.eligible_for_merge else 'SKIPPED'})
    return {'job_id':project.id,'created_at':datetime.now().isoformat(timespec='seconds'),
            'output_path':project.output_path,
            'order_list':{'enabled':bool(project.options.order_list_enabled),
                          'path':project.options.order_list_path,
                          'strict':bool(project.options.strict_order_mode),
                          'entries':list(project.options.order_list_entries)},
            'document_count':len(documents),'documents':documents}


def save_manifest(path,data):
    p=Path(path); p.parent.mkdir(parents=True,exist_ok=True); p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    return p


def load_manifest(path):
    """Nolasa manifestu; bojāts/neesošs fails atgriež None (nekad nemet)."""
    target=Path(path)
    if not target.is_file(): return None
    try: return json.loads(target.read_text(encoding='utf-8'))
    except (OSError,json.JSONDecodeError): return None


def manifest_key(document):
    """Salīdzināšanas atslēga: relative_path (casefold), citādi avota pamatnosaukums."""
    if not isinstance(document,dict): return ''
    value=document.get('relative_path') or document.get('source') or ''
    parts=[p for p in str(value).replace('\\','/').split('/') if p]
    return '/'.join(parts).casefold()


def _fingerprint(document):
    return (document.get('size_bytes'),document.get('sha256') or None)


def compare_manifest(previous,current):
    """Salīdzina divus manifestus un atgriež statusu kopsavilkumu."""
    result={'status':'FIRST_RUN','unchanged':0,'changed':0,'missing':0,'added':0,'documents':[]}
    if not previous or not isinstance(previous,dict) or not previous.get('documents'):
        for document in (current or {}).get('documents',[]):
            result['added']+=1
            result['documents'].append({'key':manifest_key(document),'status':DIFF_ADDED,'source':document.get('source')})
        return result
    old={manifest_key(d):d for d in previous.get('documents',[]) if manifest_key(d)}
    new={manifest_key(d):d for d in (current or {}).get('documents',[]) if manifest_key(d)}
    for key,document in new.items():
        old_document=old.get(key)
        if old_document is None:
            result['added']+=1
            result['documents'].append({'key':key,'status':DIFF_ADDED,'source':document.get('source')})
        elif _fingerprint(old_document)==_fingerprint(document):
            result['unchanged']+=1
            result['documents'].append({'key':key,'status':DIFF_UNCHANGED,'source':document.get('source')})
        else:
            result['changed']+=1
            result['documents'].append({'key':key,'status':DIFF_CHANGED,'source':document.get('source'),
                                        'previous_sha256':old_document.get('sha256'),'sha256':document.get('sha256')})
    for key,old_document in old.items():
        if key not in new:
            result['missing']+=1
            result['documents'].append({'key':key,'status':DIFF_MISSING,'source':old_document.get('source')})
    result['status']='OK' if not (result['changed'] or result['missing']) else 'WARN'
    return result
