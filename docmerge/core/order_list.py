"""CSV/JSON secības saraksts: parsēšana, matching, priekšskatījums un eksports.

Saraksta secība ir ABSOLŪTA merge secība — pēc tās natural sort netiek pielietots.

Matching prioritāte (kā specifikācijā):
1. relative path  ->  2. exact filename  ->  3. case-insensitive filename

Ja vienam nosaukumam atbilst vairāki faili dažādās mapēs, ieraksts kļūst
AMBIGUOUS; nejaušs fails NEDRĪKST tikt izvēlēts automātiski.

Statusi: FOUND, MISSING, AMBIGUOUS, DUPLICATE, IGNORED.
"""
from __future__ import annotations

import codecs
import csv
import io
import json
import os
from datetime import datetime
from pathlib import Path

from docmerge.core.scanner import is_word_document
from docmerge.domain.enums import BLOCKING_ORDER_STATUSES, OrderListStatus, SortMode
from docmerge.domain.errors import OrderListError, StrictOrderError
from docmerge.domain.models import OrderEntry, OrderMatch

CSV_EXTENSIONS={'.csv','.txt','.tsv'}
JSON_EXTENSIONS={'.json'}
# Galvenes nosaukumu aliasi (latviešu + angļu).
ORDER_ALIASES={'order','kartas_numurs','kartasnumurs','nr','numurs','position','seq','sequence'}
NAME_ALIASES={'filename','file','fails','faila_nosaukums','nosaukums','name'}
PATH_ALIASES={'path','cels','ceļš','rel_path','relative_path','filepath','file_path'}
COMMENT_PREFIXES=('#','//',';')
# Matching kvalitātes pakāpes -> redzams reportā.
RULE_RELATIVE_PATH='relative_path'
RULE_RELATIVE_PATH_CI='relative_path_ci'
RULE_FILENAME='filename'
RULE_FILENAME_CI='filename_ci'
RULE_NONE=''


def _looks_like_file(value):
    text=str(value or '').strip()
    if not text: return False
    return Path(text).suffix.lower() in {'.doc','.docx','.docm','.dot','.dotx','.dotm','.rtf','.odt','.txt','.csv','.json'}


def is_header_row(cells):
    """True, ja rinda ir galvene (visas ne-tukšās šūnas ir zināmi kolonnu nosaukumi)."""
    values=[str(c or '').strip().casefold() for c in cells]
    values=[v for v in values if v]
    if not values: return False
    if any(_looks_like_file(v) for v in values): return False
    known=ORDER_ALIASES|NAME_ALIASES|PATH_ALIASES
    return all(v in known for v in values)


def decode_order_bytes(raw):
    """Decode CSV/JSON baitus. Atgriež (text, encoding, warnings).

    Atbalsta UTF-8, UTF-8 BOM un Windows cp1257 (ja droši dekodējams);
    cp1252 tiek izmantots tikai kā pēdējais fallback ar brīdinājumu.
    """
    if raw.startswith(codecs.BOM_UTF8): return raw.decode('utf-8-sig'),'utf-8-sig',[]
    try: return raw.decode('utf-8'),'utf-8',[]
    except UnicodeDecodeError: pass
    try: return raw.decode('cp1257'),'cp1257',[]
    except UnicodeDecodeError:
        return raw.decode('cp1252',errors='replace'),'cp1252',['Kodējumu nevarēja droši noteikt — izmantots cp1252']


def detect_delimiter(text):
    """Nosaka CSV atdalītāju (',' ';' tab) pēc kolonnu skaita stabilitātes."""
    sample=[ln for ln in text.splitlines() if ln.strip() and not ln.lstrip().startswith(COMMENT_PREFIXES)][:10]
    if not sample: return ','
    best,best_score=',',-1.0
    for delim in (',',';','\t'):
        try: counts=[len(next(csv.reader([ln],delimiter=delim))) for ln in sample]
        except csv.Error: continue
        if not counts: continue
        most=counts.count(max(counts,key=counts.count))/len(counts)
        score=most+(1.0 if max(counts)>1 else 0.0)
        if score>best_score: best_score,best=score,delim
    return best


def _as_int(value):
    try: return int(str(value).strip())
    except (TypeError,ValueError): return None


def _header_columns(cells):
    """Atgriež (order_col, name_col, path_col) vai (None,None,None)."""
    order_col=name_col=path_col=None
    for i,cell in enumerate(cells):
        key=str(cell or '').strip().casefold()
        if not key: continue
        if order_col is None and key in ORDER_ALIASES: order_col=i
        elif name_col is None and key in NAME_ALIASES: name_col=i
        elif path_col is None and key in PATH_ALIASES: path_col=i
    return order_col,name_col,path_col


def finalize_orders(entries):
    """Piešķir secīgus order numurus tukšajiem; atkārtotos atzīmē kā problēmu."""
    problems=[]; used=set(); last=0
    for entry in entries:
        if entry.order and entry.order>0: last=entry.order
        else: last+=1; entry.order=last
    for entry in entries:
        if entry.order in used: problems.append(f'Atkārtots order numurs {entry.order} (ieraksts "{entry.requested}")')
        used.add(entry.order)
    return entries,problems


def parse_csv_entries(text, delimiter=None):
    """Parsē CSV saturu. Atgriež (entries, problems). Par saturu izņēmumu nemet."""
    entries=[]; problems=[]
    delimiter=delimiter or detect_delimiter(text)
    reader=csv.reader(io.StringIO(text),delimiter=delimiter)
    header=None; order_col=name_col=path_col=None; row_no=0
    for cells in reader:
        row_no+=1
        if not cells or all(not str(c or '').strip() for c in cells): continue
        if str(cells[0] or '').lstrip().startswith(COMMENT_PREFIXES): continue
        if header is None and is_header_row(cells):
            header=cells; order_col,name_col,path_col=_header_columns(cells); continue
        try:
            if header is not None:
                requested=''
                if path_col is not None and path_col<len(cells) and str(cells[path_col]).strip(): requested=str(cells[path_col]).strip()
                elif name_col is not None and name_col<len(cells) and str(cells[name_col]).strip(): requested=str(cells[name_col]).strip()
                else: requested=' '.join(str(c).strip() for c in cells if str(c).strip())
                order=_as_int(cells[order_col]) if (order_col is not None and order_col<len(cells)) else None
                if order_col is not None and order is None and order_col<len(cells) and str(cells[order_col]).strip():
                    problems.append(f'{row_no}. rindā nederīgs order numurs: {cells[order_col]!r}')
            else:
                values=[str(c).strip() for c in cells]
                first=values[0]; parsed=_as_int(first)
                if parsed is not None and len(values)>1 and any(values[1:]):
                    order=parsed; requested=next((v for v in values[1:] if v),'')
                else:
                    order=None; requested=first if first else next((v for v in values if v),'')
            if not requested:
                problems.append(f'{row_no}. rindā nav faila nosaukuma'); continue
            if order is not None and order<=0:
                problems.append(f'{row_no}. rindā nederīgs order numurs: {order}'); order=None
            entries.append(OrderEntry(order=order or 0,requested=requested,raw=','.join(cells),line=row_no))
        except Exception as exc:  # noqa: BLE001 - neviena rinda nedrīkst pazust
            problems.append(f'{row_no}. rindu nevarēja nolasīt: {exc!r}')
    if not entries: problems.append('CSV saraksts ir tukšs — nav neviena ieraksta')
    entries,order_problems=finalize_orders(entries)
    problems.extend(order_problems)
    return entries,problems


def parse_json_entries(text):
    """Parsē JSON sarakstu. Struktūras problēmas atgriež, bojātu JSON met kļūdu."""
    problems=[]; entries=[]
    try: data=json.loads(text)
    except json.JSONDecodeError as exc:
        raise OrderListError(f'Bojāts JSON saraksts: {exc}',context={'line':exc.lineno,'column':exc.colno}) from exc
    if isinstance(data,list): rows=data
    elif isinstance(data,dict):
        rows=None
        for key in ('files','documents','order','items','list'):
            if isinstance(data.get(key),list): rows=data[key]; break
        if rows is None:
            raise OrderListError('JSON sarakstā nav atrasts "files" masīvs',
                                 context={'keys':','.join(sorted(data.keys()))[:200]})
    else:
        raise OrderListError(f'JSON saraksta saknei jābūt masīvam vai objektam, nevis {type(data).__name__}')
    for index,row in enumerate(rows,1):
        try:
            if isinstance(row,str):
                requested=row.strip()
                if not requested: problems.append(f'{index}. ieraksts ir tukša virkne'); continue
                entries.append(OrderEntry(order=index,requested=requested,raw=row,line=index))
            elif isinstance(row,dict):
                requested=''
                for key in ('path','relative_path','rel_path','filename','file','name','nosaukums'):
                    value=row.get(key)
                    if value and str(value).strip(): requested=str(value).strip(); break
                order=_as_int(row.get('order'))
                if order is None and 'order' in row: problems.append(f'{index}. ierakstā nederīgs order: {row.get("order")!r}')
                if not requested: problems.append(f'{index}. ierakstā nav faila nosaukuma/path'); continue
                entries.append(OrderEntry(order=order or 0,requested=requested,raw=json.dumps(row,ensure_ascii=False),line=index))
            else:
                problems.append(f'{index}. ieraksts nav ne virkne, ne objekts: {type(row).__name__}')
        except Exception as exc:  # noqa: BLE001 - neviens ieraksts nedrīkst pazust
            problems.append(f'{index}. ierakstu nevarēja nolasīt: {exc!r}')
    if not entries: problems.append('JSON saraksts ir tukšs — nav neviena ieraksta')
    entries,order_problems=finalize_orders(entries)
    problems.extend(order_problems)
    return entries,problems


def format_for_path(path):
    """'JSON' vai 'CSV' pēc paplašinājuma (noklusēti CSV)."""
    return 'JSON' if Path(path).suffix.lower() in JSON_EXTENSIONS else 'CSV'


def parse_order_text(text, fmt):
    """Parsē tekstu pēc formāta nosaukuma ('CSV'/'JSON')."""
    fmt=str(fmt or '').upper()
    if fmt=='JSON': return parse_json_entries(text)
    if fmt=='CSV': return parse_csv_entries(text)
    raise OrderListError(f'Nezināms saraksta formāts: {fmt}')


def load_order_list(path):
    """Nolasa CSV/JSON sarakstu no diska. Atgriež (entries, meta).

    meta: {'path','format','encoding','problems','count','parsed_at'}
    """
    target=Path(path)
    if not target.is_file():
        raise OrderListError(f'Saraksta fails neeksistē: {path}',context={'path':str(path)})
    try: raw=target.read_bytes()
    except OSError as exc:
        raise OrderListError(f'Sarakstu nevar nolasīt: {exc}',original=exc,context={'path':str(path)}) from exc
    text,encoding,warnings=decode_order_bytes(raw)
    fmt=format_for_path(target)
    entries,problems=parse_order_text(text,fmt)
    meta={'path':str(target),'format':fmt,'encoding':encoding,'problems':list(warnings)+list(problems),
          'count':len(entries),'parsed_at':datetime.now().isoformat(timespec='seconds')}
    return entries,meta


# --------------------------------------------------------------------- matching
def normalize_parts(value):
    """'apaksmape\\01_ievads.docx' -> ('apaksmape','01_ievads.docx')."""
    text=str(value or '').replace('\\','/').strip().strip('"')
    return tuple(p for p in text.split('/') if p not in ('','.'))


def _index_items(items):
    index=[]
    for item in items:
        name=item.filename or Path(item.source_path).name
        parts=normalize_parts(item.source_path)
        index.append({'item':item,'parts':parts,'parts_ci':tuple(p.casefold() for p in parts),
                      'name':name,'name_ci':name.casefold()})
    return index


def _suffix_matches(index, parts, case_sensitive=True):
    if not parts: return []
    requested=parts if case_sensitive else tuple(p.casefold() for p in parts)
    found=[]
    for row in index:
        candidate=row['parts'] if case_sensitive else row['parts_ci']
        if len(candidate)>=len(requested) and candidate[-len(requested):]==requested: found.append(row)
    return found


def _match_one(entry, index, used_item_ids, order_seen):
    requested=str(entry.requested or '').strip()
    parts=normalize_parts(requested)
    rule=RULE_NONE; candidates=[]
    if len(parts)>1 or '/' in requested or '\\' in requested:
        candidates=_suffix_matches(index, parts, True); rule=RULE_RELATIVE_PATH
        if not candidates:
            candidates=_suffix_matches(index, parts, False); rule=RULE_RELATIVE_PATH_CI
    if not candidates:
        candidates=[r for r in index if r['name']==requested]
        if candidates: rule=RULE_FILENAME
    if not candidates:
        candidates=[r for r in index if r['name_ci']==requested.casefold()]
        if candidates: rule=RULE_FILENAME_CI
    if not candidates:
        return OrderMatch(order=entry.order,requested=requested,status=OrderListStatus.MISSING,
                          message='Nav atrasts neviens fails ar šo nosaukumu/ceļu')
    if len(candidates)>1:
        paths=', '.join(r['item'].source_path for r in candidates[:5])
        return OrderMatch(order=entry.order,requested=requested,status=OrderListStatus.AMBIGUOUS,rule=rule,
                          message=f'{len(candidates)} faili atbilst — tiek prasīts precīzs relative path: {paths}')
    row=candidates[0]; item=row['item']
    base=dict(order=entry.order,requested=requested,resolved_path=item.source_path,resolved_name=row['name'],
              detected_format=item.detected_format,item_id=item.id,rule=rule)
    if not is_word_document(item.source_path):
        if entry.order in order_seen:
            return OrderMatch(status=OrderListStatus.DUPLICATE,
                              message=f'Atkārtots order numurs {entry.order}',**base)
        return OrderMatch(status=OrderListStatus.IGNORED,message='Fails nav .doc/.docx — netiek apvienots;',**base)
    if entry.order in order_seen:
        return OrderMatch(status=OrderListStatus.DUPLICATE,
                          message=f'Atkārtots order numurs {entry.order}',**base)
    if item.id in used_item_ids:
        return OrderMatch(status=OrderListStatus.DUPLICATE,
                          message=f'Fails jau izmantots ar #{used_item_ids[item.id]}',**base)
    order_seen[entry.order]=requested; used_item_ids[item.id]=entry.order
    return OrderMatch(status=OrderListStatus.FOUND,message='',**base)


def match_entries(entries, items):
    """Sasaista saraksta ierakstus ar projekta failiem. Atgriež OrderMatch sarakstu."""
    index=_index_items(items); matches=[]; used_item_ids={}; order_seen={}
    for entry in entries:
        matches.append(_match_one(entry, index, used_item_ids, order_seen))
    return matches


def summarize(matches):
    """Kopsavilkums priekšskatījumam/reportam (Atrasti/Trūkst/Dublikāti/Neskaidri)."""
    counts={s.value:0 for s in OrderListStatus}
    for match in matches: counts[match.status.value]=counts.get(match.status.value,0)+1
    blocking=[m for m in matches if m.status in BLOCKING_ORDER_STATUSES]
    return {'total':len(matches),'found':counts.get(OrderListStatus.FOUND.value,0),
            'missing':counts.get(OrderListStatus.MISSING.value,0),
            'ambiguous':counts.get(OrderListStatus.AMBIGUOUS.value,0),
            'duplicate':counts.get(OrderListStatus.DUPLICATE.value,0),
            'ignored':counts.get(OrderListStatus.IGNORED.value,0),
            'blocking_count':len(blocking),'ok':not blocking,
            'problems':[f'#{m.order} {m.requested}: {m.message}' for m in blocking],
            'blocking':[m.to_dict() for m in blocking]}


def enforce_strict(matches, strict):
    """STRICT MODE: MISSING/AMBIGUOUS/DUPLICATE bloķē merge (met StrictOrderError)."""
    if not strict: return
    blocking=[m for m in matches if m.status in BLOCKING_ORDER_STATUSES]
    if not blocking: return
    lines=[f'#{m.order} {m.requested} [{m.status.value}] {m.message}' for m in blocking[:25]]
    raise StrictOrderError('STRICT MODE: merge bloķēts — '+f'{len(blocking)} problemātiski saraksta ieraksti:\n'+'\n'.join(lines),
                           context={'blocking_count':len(blocking),'checked':len(matches)})


# ----------------------------------------------------------------- apply / evaluate
def entries_from_project(project):
    """Saglabāto saraksta ierakstu (options.order_list_entries) atjaunošana."""
    return [OrderEntry(order=i,requested=str(r)) for i,r in enumerate(project.options.order_list_entries or [],1)]


def evaluate_order_list(project):
    """Atkārtoti novērtē saglabāto sarakstu pret projekta failiem -> (matches, summary)."""
    if not project.options.order_list_enabled: return [],summarize([])
    matches=match_entries(entries_from_project(project),project.items)
    return matches,summarize(matches)


def apply_order(project, entries, matches, strict=False, order_path=''):
    """Sakārto projekta failus ABSOLŪTI pēc saraksta.

    Sarakstā iekļautie faili tiek ieslēgti un sakārtoti saraksta secībā; pārējie
    paliek redzami, bet tiek izslēgti (`enabled=False`) ar brīdinājumu, jo
    saraksta secība ir galīgā merge secība.
    """
    by_id={item.id:item for item in project.items}
    ordered=[]; used=set()
    for match in sorted(matches,key=lambda m:(m.order,)):
        if match.status!=OrderListStatus.FOUND: continue
        item=by_id.get(match.item_id)
        if item is None or item.id in used: continue
        item.enabled=True; ordered.append(item); used.add(item.id)
    others=[item for item in project.items if item.id not in used]
    for item in others:
        item.enabled=False
        item.warnings.append('Nav order sarakstā — tiks izlaists no apvienošanas')
    project.items=ordered+others
    for i,item in enumerate(project.items,1): item.manual_order=i
    options=project.options
    options.order_list_enabled=True; options.order_list_path=str(order_path or '')
    options.order_list_entries=[str(e.requested) for e in entries]
    options.strict_order_mode=bool(strict); options.sort_mode=SortMode.ORDER_LIST
    return ordered,others


def clear_order_list(project):
    """Izslēdz saraksta režīmu un ieslēdz visus failus atpakaļ."""
    for item in project.items: item.enabled=True
    options=project.options
    options.order_list_enabled=False; options.order_list_path=''
    options.order_list_entries=[]; options.strict_order_mode=False
    if options.sort_mode==SortMode.ORDER_LIST: options.sort_mode=SortMode.NATURAL_NAME
    return project.items


def order_list_report(matches, meta=None, strict=False):
    """Reportam paredzēts saraksta kopsavilkums (nekad nezaudē problēmas)."""
    summary=summarize(matches)
    return {'meta':dict(meta or {}),'strict':bool(strict),'summary':summary,
            'entries':[m.to_dict() for m in matches]}


# ------------------------------------------------------------------------ export
def common_base(items):
    """Kopīgā mape relatīvo ceļu aprēķinam (vai None, ja nav iespējams)."""
    folders=[]
    for item in items:
        try: folders.append(Path(item.source_path).parent)
        except (TypeError,AttributeError): continue
    if not folders: return None
    try: return Path(os.path.commonpath([str(f) for f in folders]))
    except (ValueError,OSError): return folders[0]


def relative_for(path, base=None):
    """Relatīvs ceļš pret bāzi; ja nav iespējams, tikai faila nosaukums."""
    target=Path(path)
    if base is None: return target.name
    try: return str(target.relative_to(base))
    except ValueError: return target.name


def export_order(path, items, fmt=None):
    """Saglabā pašreizējo secību CSV vai JSON (lai vēlāk varētu atkārtot identiski).

    Atgriež (target_path, format).
    """
    target=Path(path); fmt=str(fmt or format_for_path(target)).upper()
    if fmt not in ('CSV','JSON'):
        raise OrderListError(f'Neatbalstīts eksporta formāts: {fmt}')
    base=common_base(items)
    rows=[{'order':i,'filename':(item.filename or Path(item.source_path).name),
           'path':relative_for(item.source_path,base)} for i,item in enumerate(items,1)]
    target.parent.mkdir(parents=True,exist_ok=True)
    if fmt=='JSON':
        payload={'exported_at':datetime.now().isoformat(timespec='seconds'),'count':len(rows),'files':rows}
        target.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
    else:
        # utf-8-sig (BOM), lai Excel atvērtu latviešu burtus pareizi; lasītājs BOM saprot.
        with target.open('w',encoding='utf-8-sig',newline='') as fh:
            writer=csv.writer(fh); writer.writerow(['order','filename','path'])
            for row in rows: writer.writerow([row['order'],row['filename'],row['path']])
    return target,fmt
