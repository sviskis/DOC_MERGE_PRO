"""OFFLINE QUICK testi: imports, compileall, config, saraksti, matching, strict, manifest.

Tie neizmanto tīklu, API, cloud vai lejupielādes (skat. `guard.py`) un neprasa Word.
"""
from __future__ import annotations

import compileall
import importlib
import json
import os
from pathlib import Path
from unittest.mock import patch

from docmerge.core import order_list as ol
from docmerge.core.hashing import sha256_file
from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.paths import ensure_word_compatible_path, path_too_long
from docmerge.core.preflight import preflight_all
from docmerge.core.scanner import filter_word_documents, is_word_document, scan_folder
from docmerge.core.sorter import natural_key, sort_items
from docmerge.domain.enums import OrderListStatus, SortMode
from docmerge.domain.errors import ErrorCode, MergeStage, OrderListError, StrictOrderError, ValidationError
from docmerge.domain.models import DocumentItem, Project
from docmerge.persistence.manifest_store import build_manifest, compare_manifest, load_manifest, save_manifest
from docmerge.persistence.project_store import load_project, save_project
from docmerge.reporting.result_report import write_result
from docmerge.testing.harness import Suite, check, skip, warn
from docmerge.testing.fixtures import (clean_dir, make_docx, make_read_only, make_writable, write_zone_identifier,
                                       has_zone_identifier)

SECTION_IMPORTS='imports'
SECTION_CONFIG='config'
SECTION_PROJECT='project'
SECTION_SCANNER='scanner'
SECTION_SORT='sort'
SECTION_CSV='csv'
SECTION_JSON='json'
SECTION_MATCH='matching'
SECTION_STRICT='strict'
SECTION_PATHS='paths'
SECTION_UNLOCK='unlock'
SECTION_MANIFEST='manifest'
SECTION_REPORT='report'
PROJECT_ROOT=Path(__file__).resolve().parent.parent.parent


def package_modules():
    """Visi docmerge pakotnes moduļi."""
    root=Path(__file__).resolve().parent.parent
    names=[]
    for path in sorted(root.rglob('*.py')):
        if '__pycache__' in path.parts: continue
        rel=path.relative_to(root.parent).with_suffix('')
        parts=list(rel.parts)
        if parts[-1]=='__init__': parts=parts[:-1]
        if parts: names.append('.'.join(parts))
    return sorted(set(names))


def add_import_and_config_tests(suite:Suite):
    for module in package_modules():
        with suite.test(f'import {module}',section=SECTION_IMPORTS):
            importlib.import_module(module)
    with suite.test('root launcher DOC_MERGE_PRO.py',section=SECTION_IMPORTS):
        launcher=PROJECT_ROOT/'DOC_MERGE_PRO.py'
        check(launcher.is_file(),'DOC_MERGE_PRO.py nav projekta root mapē')
        text=launcher.read_text(encoding='utf-8')
        check('run_app' in text,'Launcher nepalaiž GUI entry point (run_app)')
    with suite.test('app.py compatibility launcher',section=SECTION_IMPORTS):
        text=(PROJECT_ROOT/'app.py').read_text(encoding='utf-8')
        check('run_app' in text,'app.py nepalaiž to pašu GUI')
    with suite.test('compileall docmerge',section=SECTION_CONFIG):
        check(compileall.compile_dir(str(Path(__file__).resolve().parent.parent),quiet=2,force=True),
              'compileall atrada sintakses kļūdas docmerge pakotnē')
    with suite.test('compileall tests',section=SECTION_CONFIG):
        check(compileall.compile_dir(str(PROJECT_ROOT/'tests'),quiet=2,force=True),'compileall kļūda tests mapē')
    with suite.test('pyproject.toml konfigurācija',section=SECTION_CONFIG):
        import tomllib
        data=tomllib.loads((PROJECT_ROOT/'pyproject.toml').read_text(encoding='utf-8'))
        check(data['project']['name']=='doc-merge-pro',f"Nepareizs projekta nosaukums: {data['project'].get('name')}")
        ini=data.get('tool',{}).get('pytest',{}).get('ini_options',{})
        check('tests' in ini.get('testpaths',[]),'pytest testpaths nav definēts (tool.pytest.ini_options)')
    with suite.test('requirements.txt',section=SECTION_CONFIG):
        text=(PROJECT_ROOT/'requirements.txt').read_text(encoding='utf-8')
        check('pywin32' in text and 'pytest' in text,'requirements.txt trūkst pywin32/pytest')
        for line in text.splitlines():
            if line.strip() and not line.strip().startswith('#'):
                check(any(token in line for token in ('==','>=',';')),f'Nepašaubāma prasība: {line}')


def add_project_and_scanner_tests(suite:Suite,work):
    sources=clean_dir(work/'sources')
    with suite.test('projekta roundtrip (saglabāt/ielādēt)',section=SECTION_PROJECT):
        for i in range(2): make_docx(sources/f'proj_{i}.docx',f'Teksts {i}')
        items=preflight_all([DocumentItem(str(sources/f'proj_{i}.docx')) for i in range(2)])
        project=Project(name='Roundtrip',items=items,output_path=str(work/'out.docx'))
        project.options.sort_mode=SortMode.ORDER_LIST; project.options.order_list_enabled=True
        project.options.order_list_path=str(work/'order.json')
        project.options.order_list_entries=['proj_1.docx','proj_0.docx']
        project.options.strict_order_mode=True; project.options.open_output_when_finished=True
        path=work/'project.docmerge.json'; save_project(project,path)
        loaded=load_project(path)
        check(loaded.name=='Roundtrip','Projekta nosaukums nav saglabāts')
        check(loaded.output_path==project.output_path,'Izvades ceļš nav saglabāts')
        check(len(loaded.items)==2,'Vienumu skaits nav saglabāts')
        check(loaded.options.order_list_entries==['proj_1.docx','proj_0.docx'],'Saraksta ieraksti nav saglabāti')
        check(loaded.options.strict_order_mode is True,'STRICT MODE nav saglabāts')
        check(loaded.options.open_output_when_finished is True,'OPEN OUTPUT nav saglabāts')
        check(loaded.options.sort_mode==SortMode.ORDER_LIST,'Secības režīms nav saglabāts')
        check(loaded.items[0].filename==items[0].filename,'Faila nosaukums nav saglabāts')

    with suite.test('skeneris atrod .doc/.docx un izlaiž ~$',section=SECTION_SCANNER):
        scan=clean_dir(work/'scan')
        make_docx(scan/'a.docx','A'); (scan/'b.doc').write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
        (scan/'c.txt').write_text('x',encoding='utf-8'); (scan/'~$temp.docx').write_bytes(b'x')
        (scan/'sub').mkdir(exist_ok=True); make_docx(scan/'sub'/'d.docx','D')
        items=scan_folder(str(scan),recursive=True)
        names=sorted(Path(x.source_path).name for x in items)
        check(names==['a.docx','b.doc','c.txt','d.docx'],f'Skeneris atgrieza nepareizus failus: {names}')
        kept=[Path(x.source_path).name for x in filter_word_documents(items)]
        check(kept==['a.docx','b.doc','d.docx'],f'Word filtrs kļūdains: {kept}')
        check(is_word_document('X:/a.DOC') and not is_word_document('X:/a.rtf'),'is_word_document kļūda')
    with suite.test('skeneris: neeksistējoša mape met kļūdu',section=SECTION_SCANNER):
        try: scan_folder(str(work/'neeksiste_999'))
        except ValueError as exc:
            check('neeksistē' in str(exc),f'Nepareizs kļūdas teksts: {exc}')
        else: raise AssertionError('scan_folder neeksistējošai mapei bija jāmet ValueError')
    with suite.test('preflight atzīmē trūkstošu un 0 B failu',section=SECTION_SCANNER):
        empty=work/'empty.docx'; empty.write_bytes(b'')
        missing=DocumentItem(str(work/'pazudis.docx'))
        items=preflight_all([DocumentItem(str(empty)),missing])
        check(any('tukšs' in e for e in items[0].errors),'0 B fails netika atzīmēts')
        check(any('neeksistē' in e for e in items[1].errors),'Trūkstošs fails netika atzīmēts')
        check(not items[1].eligible_for_merge,'Trūkstošs fails nedrīkst būt eligible')


def add_sort_tests(suite:Suite,work):
    with suite.test('natural sort kārtība',section=SECTION_SORT):
        check(sorted(['10.docx','2.docx','1.docx'],key=natural_key)==['1.docx','2.docx','10.docx'],'natural_key kļūda')
        items=[DocumentItem(str(work/f'{n}.docx')) for n in ('10','2','1')]
        for item in items: item.hydrate_from_path()
        ordered=sort_items(items,SortMode.NATURAL_NAME)
        check([Path(x.source_path).name for x in ordered]==['1.docx','2.docx','10.docx'],'NATURAL_NAME kārtošana kļūda')
    with suite.test('manuālā secība (manual_order)',section=SECTION_SORT):
        items=[]
        for index,name in enumerate(('a.docx','b.docx','c.docx')):
            item=DocumentItem(str(work/name)); item.manual_order=3-index; items.append(item)
        ordered=sort_items(items,SortMode.MANUAL)
        check([x.manual_order for x in ordered]==[1,2,3],'MANUAL kārtošana kļūda')
    with suite.test('ORDER_LIST secība ir absolūta (bez natural sort)',section=SECTION_SORT):
        items=[]
        for index,name in enumerate(('C.docx','A.docx','B.docx')):
            item=DocumentItem(str(work/name)); item.filename=name; item.manual_order=index+1; items.append(item)
        ordered=sort_items(items,SortMode.ORDER_LIST)
        check([x.filename for x in ordered]==['C.docx','A.docx','B.docx'],'ORDER_LIST secība netika ievērota')
        natural=sort_items(list(items),SortMode.NATURAL_NAME)
        check([x.filename for x in natural]==['A.docx','B.docx','C.docx'],'Kontroles pārbaude (natural sort) neizdevās')
    with suite.test('ORDER_LIST ir idempotenta',section=SECTION_SORT):
        items=[DocumentItem(str(work/f'{n}.docx')) for n in ('demo_2','demo_10','demo_1')]
        for index,item in enumerate(items): item.filename=Path(item.source_path).name; item.manual_order=index+1
        first=sort_items(items,SortMode.ORDER_LIST); second=sort_items(first,SortMode.ORDER_LIST)
        check([x.filename for x in second]==[x.filename for x in first],'ORDER_LIST secība nav idempotenta')


def _write(path,text,encoding='utf-8'):
    Path(path).write_text(text,encoding=encoding); return path


def add_csv_tests(suite:Suite,work):
    with suite.test('CSV: vienas kolonnas formāts (filename)',section=SECTION_CSV):
        entries,problems=ol.parse_csv_entries('filename\n01_ievads.docx\n05_stasts.doc\n02_nodala.docx\n')
        check(len(entries)==3,f'Gaidīti 3 ieraksti, iegūti {len(entries)}')
        check([e.requested for e in entries]==['01_ievads.docx','05_stasts.doc','02_nodala.docx'],'CSV secība kļūdaina')
        check([e.order for e in entries]==[1,2,3],'Order numuri kļūdaini')
    with suite.test('CSV: bez galvenes (tikai nosaukumi)',section=SECTION_CSV):
        entries,_=ol.parse_csv_entries('01_ievads.docx\n05_stasts.doc\n')
        check([e.requested for e in entries]==['01_ievads.docx','05_stasts.doc'],'CSV bez galvenes kļūdains')
    with suite.test('CSV: order,filename formāts',section=SECTION_CSV):
        entries,_=ol.parse_csv_entries('order,filename\n1,01_ievads.docx\n2,05_stasts.doc\n3,02_nodala.docx\n')
        check([(e.order,e.requested) for e in entries]==[(1,'01_ievads.docx'),(2,'05_stasts.doc'),(3,'02_nodala.docx')],
              'order,filename formāts kļūdains')
    with suite.test('CSV: path formāts ar apakšmapēm',section=SECTION_CSV):
        entries,_=ol.parse_csv_entries('path\napaksmape\\01_ievads.docx\napaksmape\\05_stasts.doc\n')
        check([e.requested for e in entries]==['apaksmape\\01_ievads.docx','apaksmape\\05_stasts.doc'],'path formāts kļūdains')
    with suite.test('CSV: order;path ar ; atdalītāju',section=SECTION_CSV):
        entries,_=ol.parse_csv_entries('order;path\n2;b\\x.docx\n1;a\\y.docx\n')
        check([(e.order,e.requested) for e in entries]==[(2,'b\\x.docx'),(1,'a\\y.docx')],'order;path formāts kļūdains')
    with suite.test('CSV: UTF-8 BOM',section=SECTION_CSV):
        path=work/'bom.csv'; path.write_text('filename\nāčēģī_01.docx\n',encoding='utf-8-sig')
        entries,meta=ol.load_order_list(path)
        check(meta['encoding']=='utf-8-sig',f"BOM kodējums nav atpazīts: {meta['encoding']}")
        check(entries[0].requested=='āčēģī_01.docx',f"BOM nosaukums kļūdains: {entries[0].requested}")
    with suite.test('CSV: cp1257 kodējums',section=SECTION_CSV):
        path=work/'cp1257.csv'; path.write_bytes('filename\nāčēģīķļņšūž_01.docx\n'.encode('cp1257'))
        entries,meta=ol.load_order_list(path)
        check(meta['encoding']=='cp1257',f"cp1257 nav atpazīts: {meta['encoding']}")
        check(entries[0].requested=='āčēģīķļņšūž_01.docx',f'cp1257 teksts kļūdains: {entries[0].requested}')
    with suite.test('CSV: komentāri un tukšas rindas',section=SECTION_CSV):
        entries,problems=ol.parse_csv_entries('# komentars\n\nfilename\n01_a.docx\n')
        check([e.requested for e in entries]==['01_a.docx'],f'Komentāri netika apstrādāti: {entries}')
    with suite.test('CSV: dublikāti un bojātas rindas',section=SECTION_CSV):
        entries,problems=ol.parse_csv_entries('order,filename\n1,a.docx\n1,b.docx\nx,c.docx\n,e.docx\n')
        check([e.order for e in entries]==[1,1,2,3],f'Order numuri kļūdaini: {[e.order for e in entries]}')
        check(any('Atkārtots order' in p for p in problems),f'Atkārtots order netika atzīmēts: {problems}')
        check(any('nederīgs order' in p for p in problems),f'Nederīgs order netika atzīmēts: {problems}')
        # Tukša order šūna nav kļūda: ieraksts saņem secīgu numuru (dokumentēta uzvedība).
        check(entries[-1].requested=='e.docx' and entries[-1].order==3,f'Tukša order šūna: {entries[-1]}')
    with suite.test('CSV: tukšs fails',section=SECTION_CSV):
        entries,problems=ol.parse_csv_entries('')
        check(entries==[],'Tukšam CSV jābūt bez ierakstiem')
        check(any('tukšs' in p.lower() for p in problems),f'Tukšs CSV netika atzīmēts: {problems}')
    with suite.test('CSV: neesošs fails met OrderListError',section=SECTION_CSV):
        try: ol.load_order_list(work/'neeksiste_123.csv')
        except OrderListError as exc:
            check(exc.error_code==ErrorCode.E_ORDER_LIST,f'Nepareizs kļūdas kods: {exc.error_code}')
        else: raise AssertionError('Neeksistējošam sarakstam bija jāmet OrderListError')


def add_json_tests(suite:Suite,work):
    with suite.test('JSON: masīvs ar virknēm',section=SECTION_JSON):
        entries,_=ol.parse_json_entries('["01_ievads.docx","05_stasts.doc","02_nodala.docx"]')
        check([e.requested for e in entries]==['01_ievads.docx','05_stasts.doc','02_nodala.docx'],'JSON masīvs kļūdains')
        check([e.order for e in entries]==[1,2,3],'JSON order numuri kļūdaini')
    with suite.test('JSON: {"files": [...]}',section=SECTION_JSON):
        entries,_=ol.parse_json_entries('{"files":["01_ievads.docx","05_stasts.doc"]}')
        check([e.requested for e in entries]==['01_ievads.docx','05_stasts.doc'],'files masīvs kļūdains')
    with suite.test('JSON: {"files":[{order,filename}]}',section=SECTION_JSON):
        raw='{"files":[{"order":1,"filename":"01_ievads.docx"},{"order":2,"filename":"05_stasts.doc"}]}'
        entries,_=ol.parse_json_entries(raw)
        check([(e.order,e.requested) for e in entries]==[(1,'01_ievads.docx'),(2,'05_stasts.doc')],'JSON objektu masīvs kļūdains')
    with suite.test('JSON: faila lasīšana un formāta noteikšana',section=SECTION_JSON):
        raw='[{"order":3,"filename":"c.docx"},{"order":1,"filename":"a.docx"},{"order":2,"filename":"b.docx"}]'
        entries,meta=ol.load_order_list(_write(work/'order_unsorted.json',raw))
        check(meta['format']=='JSON','Formāts nav JSON')
        check([e.requested for e in entries]==['c.docx','a.docx','b.docx'],'JSON ierakstu secība netika saglabāta')
    with suite.test('JSON: bojāts saturs met OrderListError',section=SECTION_JSON):
        try: ol.parse_json_entries('{"files":[{"order":1,')
        except OrderListError as exc:
            check('Bojāts JSON' in exc.message,f'Nepareizs kļūdas teksts: {exc.message}')
            check(exc.error_code==ErrorCode.E_ORDER_LIST,f'Nepareizs kods: {exc.error_code}')
        else: raise AssertionError('Bojātam JSON bija jāmet OrderListError')
    with suite.test('JSON: nederīga sakne un tukšs masīvs',section=SECTION_JSON):
        try: ol.parse_json_entries('123')
        except OrderListError as exc: check('saknei' in exc.message,f'Kļūda: {exc.message}')
        else: raise AssertionError('Nederīgai JSON saknei bija jāmet kļūda')
        entries,problems=ol.parse_json_entries('{"files":[]}')
        check(entries==[] and any('tukšs' in p.lower() for p in problems),f'Tukšs JSON kļūdains: {problems}')
    with suite.test('JSON: dublikāti un nederīgi ieraksti',section=SECTION_JSON):
        raw='{"files":[{"order":1,"filename":"a.docx"},{"order":1,"filename":"b.docx"},42,{"filename":""}]}'
        entries,problems=ol.parse_json_entries(raw)
        check(len(entries)==2,f'Gaidīti 2 derīgi ieraksti: {entries}')
        check(any('Atkārtots order' in p for p in problems),f'Atkārtots order netika atzīmēts: {problems}')
        check(any('nav ne virkne' in p for p in problems),f'Nederīgs ieraksts netika atzīmēts: {problems}')
        check(any('nav faila nosaukuma' in p for p in problems),f'Tukšs ieraksts netika atzīmēts: {problems}')


def _items(paths):
    items=[DocumentItem(str(p)) for p in paths]
    for item in items: item.hydrate_from_path()
    return items


def build_match_fixture(work):
    """Avoti matching testiem: apakšmapes, .doc/.docx/.txt, dublikāts citā mapē."""
    root=clean_dir(work/'match')
    sub=root/'apaksmape'; sub.mkdir(parents=True,exist_ok=True)
    other=root/'cita'; other.mkdir(parents=True,exist_ok=True)
    make_docx(root/'01_ievads.docx','IEVADS')
    (root/'05_stasts.doc').write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
    make_docx(sub/'02_nodala.docx','NODALA')
    (root/'piezimes.txt').write_text('nav Word',encoding='utf-8')
    make_docx(other/'01_ievads.docx','CITS IEVADS')
    return root


def add_matching_tests(suite:Suite,work):
    root=build_match_fixture(work)
    items=_items([root/'01_ievads.docx',root/'05_stasts.doc',root/'piezimes.txt',
                  root/'apaksmape'/'02_nodala.docx',root/'cita'/'01_ievads.docx'])
    with suite.test('matching: FOUND pēc precīza faila nosaukuma',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.FOUND,f'Statuss: {matches[0].status} {matches[0].message}')
        check(Path(matches[0].resolved_path).name=='05_stasts.doc','Nepareizs atrisinātais fails')
        check(matches[0].rule==ol.RULE_FILENAME,f'Nepareizs matching noteikums: {matches[0].rule}')
    with suite.test('matching: relative path ir prioritāte 1',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('path\napaksmape\\02_nodala.docx\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.FOUND,f'Statuss: {matches[0].status}')
        check(matches[0].rule==ol.RULE_RELATIVE_PATH,f'Nepareizs noteikums: {matches[0].rule}')
        check('apaksmape' in matches[0].resolved_path,'Nepareizs ceļš')
    with suite.test('matching: case-insensitive faila nosaukums',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\n05_STASTS.DOC\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.FOUND,f'Statuss: {matches[0].status}')
        check(matches[0].rule in (ol.RULE_FILENAME_CI,ol.RULE_RELATIVE_PATH_CI),f'Noteikums: {matches[0].rule}')
    with suite.test('matching: MISSING',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\nxyz_neeksiste.docx\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.MISSING,f'Statuss: {matches[0].status}')
        check('Nav atrasts' in matches[0].message,'Trūkst paskaidrojuma')
    with suite.test('matching: AMBIGUOUS (vairākas mapes)',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\n01_ievads.docx\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.AMBIGUOUS,f'Statuss: {matches[0].status} {matches[0].message}')
        check(matches[0].resolved_path=='','Neskaidrā gadījumā nedrīkst automātiski izvēlēties failu')
        check('faili atbilst' in matches[0].message,'Trūkst paskaidrojuma')
    with suite.test('matching: IGNORED (nav Word fails)',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\npiezimes.txt\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.IGNORED,f'Statuss: {matches[0].status}')
    with suite.test('matching: DUPLICATE (viens fails divreiz)',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\n05_stasts.doc\n')
        matches=ol.match_entries(entries,items)
        check(matches[0].status==OrderListStatus.FOUND,'Pirmais ieraksts nedrīkst būt dublikāts')
        check(matches[1].status==OrderListStatus.DUPLICATE,f'Statuss: {matches[1].status}')
    with suite.test('matching: DUPLICATE (atkārtots order numurs)',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('order,filename\n1,05_stasts.doc\n1,piezimes.txt\n')
        matches=ol.match_entries(entries,items)
        check(matches[1].status==OrderListStatus.DUPLICATE,f'Statuss: {matches[1].status}')
        check('Atkārtots order' in matches[1].message,'Trūkst paskaidrojuma')
    with suite.test('matching: kopsavilkums (Atrasti/Trūkst/Dublikāti/Neskaidri)',section=SECTION_MATCH):
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\nxyz.docx\n01_ievads.docx\n05_stasts.doc\n')
        summary=ol.summarize(ol.match_entries(entries,items))
        check(summary['found']==1 and summary['missing']==1 and summary['ambiguous']==1 and summary['duplicate']==1,
              f'Kopsavilkums kļūdains: {summary}')
        check(summary['blocking_count']==3,f"Bloķējošo skaits: {summary['blocking_count']}")
        check(summary['ok'] is False,'Kopsavilkumam jābūt ar kļūdām')
    with suite.test('apply_order: saraksts nosaka secību, pārējie tiek izslēgti',section=SECTION_MATCH):
        project=Project(name='Match',items=list(items))
        entries,_=ol.parse_csv_entries('order,filename\n1,05_stasts.doc\n2,piezimes.txt\n')
        matches=ol.match_entries(entries,project.items)
        ordered,others=ol.apply_order(project,entries,matches,strict=False,order_path=str(work/'order.csv'))
        check([Path(x.source_path).name for x in ordered]==['05_stasts.doc'],f'Sakārtotie: {[x.filename for x in ordered]}')
        check(len(others)==len(items)-1,f'Ārpus saraksta: {len(others)}')
        check(all(not item.enabled for item in others),'Faili ārpus saraksta nedrīkst būt ieslēgti')
        check(project.items[0].filename=='05_stasts.doc','Saraksta fails nav pirmais')
        check(project.options.sort_mode==SortMode.ORDER_LIST,'Secības režīms nav ORDER_LIST')
        check(project.options.order_list_entries==['05_stasts.doc','piezimes.txt'],'Saraksta ieraksti nav saglabāti')
        rematches,summary=ol.evaluate_order_list(project)
        check(len(rematches)==2,'Saglabātā saraksta atkārtota novērtēšana neizdevās')
    with suite.test('eksporta roundtrip CSV un JSON',section=SECTION_MATCH):
        # Roundtrip pārbauda unikālus nosaukumus; dublikāti dažādās mapēs ir
        # apzināti AMBIGUOUS (skat. matching testus augstāk).
        root=clean_dir(work/'export_roundtrip')
        (root/'viena').mkdir(parents=True,exist_ok=True); (root/'otra').mkdir(parents=True,exist_ok=True)
        paths=[make_docx(root/'viena'/'pirmais.docx','EXPORT 1'),
               make_docx(root/'otra'/'otrais.docx','EXPORT 2'),
               make_docx(root/'tresais.docx','EXPORT 3')]
        reversed_items=list(reversed(_items(paths)))
        project=Project(name='Export',items=reversed_items)
        for suffix in ('csv','json'):
            target=work/f'export_order.{suffix}'
            exported,fmt=ol.export_order(target,project.items)
            check(exported.is_file(),f'Eksports netika izveidots: {exported}')
            entries,meta=ol.load_order_list(exported)
            check(len(entries)==len(project.items),f'{fmt} eksportā {len(entries)} ieraksti, gaidīti {len(project.items)}')
            replayed=ol.match_entries(entries,project.items)
            summary=ol.summarize(replayed)
            check(summary['found']==len(project.items),f'{fmt} roundtrip neatrada visus failus: {summary}')
            check([m.resolved_name for m in replayed]==[x.filename for x in project.items],
                  f'{fmt} roundtrip secība atšķiras: {[m.resolved_name for m in replayed]}')


def add_strict_tests(suite:Suite,work):
    root=build_match_fixture(work)
    # Preflight ir obligāts: bez tā faili nav 'eligible_for_merge' (tāpat kā produkcijā,
    # kur MergeService.preflight() tiek izsaukts pirms run()).
    items=preflight_all(_items([root/'01_ievads.docx',root/'05_stasts.doc',
                                root/'apaksmape'/'02_nodala.docx']))
    with suite.test('STRICT: enforce_strict met StrictOrderError',section=SECTION_STRICT):
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\nxyz.docx\n')
        matches=ol.match_entries(entries,items)
        try: ol.enforce_strict(matches,True)
        except StrictOrderError as exc:
            check(exc.error_code==ErrorCode.E_STRICT_ORDER,f'Kods: {exc.error_code}')
            check('STRICT MODE' in exc.message,'Kļūdas tekstā nav STRICT MODE')
            check('xyz.docx' in exc.message,'Kļūdas tekstā nav problemātiskā ieraksta')
            check(exc.stage==MergeStage.ORDER_LIST,f'Stage: {exc.stage}')
        else: raise AssertionError('STRICT režīmā bija jāmet StrictOrderError')
    with suite.test('STRICT: izslēgts režīms nebloķē',section=SECTION_STRICT):
        entries,_=ol.parse_csv_entries('filename\nxyz.docx\n')
        ol.enforce_strict(ol.match_entries(entries,items),False)
    with suite.test('STRICT: tīrs saraksts iztur STRICT',section=SECTION_STRICT):
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\n')
        ol.enforce_strict(ol.match_entries(entries,items),True)
    with suite.test('STRICT: merge tiek bloķēts PIRMS Word',section=SECTION_STRICT):
        from docmerge.core.merge_service import MergeService
        project=Project(name='Strict',items=list(items),output_path=str(work/'strict_out.docx'))
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\npazudis_999.docx\n')
        matches=ol.match_entries(entries,project.items)
        ol.apply_order(project,entries,matches,strict=True,order_path=str(work/'strict.csv'))
        engine_calls=[]
        with patch('docmerge.core.merge_service.WordComEngine') as engine:
            engine.side_effect=lambda *a,**k: engine_calls.append(a)
            try: MergeService().run(project)
            except StrictOrderError as exc: check('STRICT MODE' in exc.message,'Nepareizs kļūdas teksts')
            else: raise AssertionError('MergeService.run bija jāmet StrictOrderError')
        check(engine_calls==[],'Word dzinējs tika izsaukts, lai gan STRICT bloķēja merge')
        check(not (work/'strict_out.docx').exists(),'Izvades fails tika izveidots, lai gan merge bija bloķēts')
    with suite.test('STRICT: izslēgts režīms ļauj merge ar izlaistiem ierakstiem',section=SECTION_STRICT):
        from docmerge.core.merge_service import MergeService
        project=Project(name='Lenient',items=list(items),output_path=str(work/'lenient_out.docx'))
        entries,_=ol.parse_csv_entries('filename\n05_stasts.doc\npazudis_999.docx\n')
        matches=ol.match_entries(entries,project.items)
        ol.apply_order(project,entries,matches,strict=False,order_path=str(work/'lenient.csv'))
        captured={}
        class FakeEngine:
            def __init__(self,logger=None): pass
            def merge(self,merge_items,output_path,*a,**k):
                # Produkcijā WordComEngine pats izfiltrē tikai eligible failus.
                captured['items']=[Path(x.source_path).name for x in merge_items if x.eligible_for_merge]
                Path(output_path).write_text('fake',encoding='utf-8')
                return {'ok':True,'output_path':str(output_path),'merged':len(captured['items']),'errors':[],
                        'normalized':[],'unlocked':[],'documents_add_attempts':1,'warning':None,
                        'total':len(captured['items'])}
        with patch('docmerge.core.merge_service.WordComEngine',FakeEngine):
            result=MergeService().run(project)
        check(captured['items']==['05_stasts.doc'],f"Merge saņēma: {captured['items']}")
        check(result['order_list']['summary']['missing']==1,'Trūkstošais ieraksts nav reportā')


def add_paths_tests(suite:Suite,work):
    with suite.test('absolūti ceļi Word COM',section=SECTION_PATHS):
        absolute=ensure_word_compatible_path('runtime/tests/relatīvs.docx')
        check(os.path.isabs(absolute),f'Ceļš nav absolūts: {absolute}')
        check(Path(absolute).is_absolute(),'Path nav absolūts')
    with suite.test('tukšs ceļš -> E_INVALID_PATH',section=SECTION_PATHS):
        try: ensure_word_compatible_path('')
        except ValidationError as exc: check(exc.error_code==ErrorCode.E_INVALID_PATH,f'Kods: {exc.error_code}')
        else: raise AssertionError('Tukšam ceļam bija jāmet ValidationError')
    with suite.test('pārāk garš ceļš -> E_PATH_TOO_LONG',section=SECTION_PATHS):
        long_path=str(work/('x'*300)/'faila.docx')
        check(path_too_long(long_path),'path_too_long neatpazina garu ceļu')
        try: ensure_word_compatible_path(long_path)
        except ValidationError as exc: check(exc.error_code==ErrorCode.E_PATH_TOO_LONG,f'Kods: {exc.error_code}')
        else: raise AssertionError('Garam ceļam bija jāmet ValidationError')
    with suite.test('unicode un iekavas ceļā netiek bojātas',section=SECTION_PATHS):
        tricky=ensure_word_compatible_path(str(work/'Latviešu (tests) 01.docx'))
        check('Latviešu (tests) 01.docx' in tricky,'Ceļš tika bojāts')


def add_unlock_tests(suite:Suite,work):
    with suite.test('Zone.Identifier: uzlikšana un noņemšana',section=SECTION_UNLOCK):
        if os.name!='nt': skip('Zone.Identifier ir Windows NTFS funkcija')
        path=work/'zone_test.docx'; make_docx(path,'ZONE')
        write_zone_identifier(path)
        check(has_zone_identifier(path) is True,'Zone.Identifier netika uzlikts')
        before=sha256_file(str(path))
        from docmerge.core.unlock import unlock_file
        zone,read_only=unlock_file(str(path))
        check(has_zone_identifier(path) is False,'Zone.Identifier netika noņemts')
        check(sha256_file(str(path))==before,'Faila saturs tika mainīts')
    with suite.test('Read-only + Zone.Identifier kopija, avots nemainīts',section=SECTION_UNLOCK):
        if os.name!='nt': skip('Zone.Identifier ir Windows NTFS funkcija')
        from docmerge.core.unlock import copy_unlocked
        source=work/'readonly_source.docx'
        make_writable(source) if source.exists() else None
        make_docx(source,'READONLY')
        try:
            write_zone_identifier(source); make_read_only(source)
            target=work/'unlocked_copy.docx'
            copy_unlocked(source,target)
            check(target.is_file(),'Atbloķētā kopija netika izveidota')
            check(has_zone_identifier(target) is False,'Kopijai palika Zone.Identifier')
            check(has_zone_identifier(source) is True,'Avotam tika noņemts marķējums (avots ir immutable)')
        finally:
            make_writable(source)  # temp artifacts jāpaliek dzēšamiem
    with suite.test('unlock_file tīram failam neko nemaina',section=SECTION_UNLOCK):
        if os.name!='nt': skip('Zone.Identifier ir Windows NTFS funkcija')
        from docmerge.core.unlock import unlock_file
        path=work/'unlock_flags.docx'; make_docx(path,'FLAGS')
        check(unlock_file(str(path))==(False,False),'Tīram failam jābūt (False, False)')


def add_manifest_tests(suite:Suite,work):
    with suite.test('manifestā relative path, size, mtime un SHA-256',section=SECTION_MANIFEST):
        root=clean_dir(work/'manifest')
        paths=[root/f'm_{i}.docx' for i in range(2)]
        for path in paths: make_docx(path,f'MANIFEST {path.name}')
        items=preflight_all(_items(paths))
        project=Project(name='Manifest',items=items,output_path=str(work/'manifest_out.docx'))
        manifest=build_manifest(project,items)
        check(len(manifest['documents'])==2,'Manifesta dokumentu skaits kļūdains')
        document=manifest['documents'][0]
        for key in ('relative_path','size_bytes','modified_time','sha256'):
            check(key in document,F'Manifestā trūkst lauka {key}')
        check(document['sha256']==sha256_file(str(paths[0])),'SHA-256 neatbilst')
        check(document['size_bytes']==paths[0].stat().st_size,'Izmērs neatbilst')
        check(manifest['order_list']['enabled'] is False,'Order saraksta sadaļa kļūdaina')
    with suite.test('manifesta salīdzinājums: UNCHANGED/CHANGED/ADDED',section=SECTION_MANIFEST):
        root=clean_dir(work/'manifest_diff')
        first=root/'first.docx'; second=root/'second.docx'; third=root/'third.docx'
        for path in (first,second): make_docx(path,f'SATURS {path.name}')
        project=Project(name='Diff',output_path=str(work/'diff_out.docx'))
        previous=build_manifest(project,preflight_all(_items([first,second])))
        make_docx(second,'SATURS MAINĪTS (garāks nekā pirms tam)')
        make_docx(third,'JAUNS')
        current=build_manifest(project,preflight_all(_items([first,second,third])))
        diff=compare_manifest(previous,current)
        check(diff['unchanged']==1,f"UNCHANGED: {diff['unchanged']}")
        check(diff['changed']==1,f"CHANGED: {diff['changed']}")
        check(diff['added']==1,f"ADDED: {diff['added']}")
        check(diff['missing']==0,f"MISSING: {diff['missing']}")
        check(diff['status']=='WARN',f"Statuss: {diff['status']}")
        statuses={entry['status'] for entry in diff['documents']}
        check({'UNCHANGED','CHANGED','ADDED'}<=statuses,f'Statusi: {statuses}')
    with suite.test('manifests: MISSING, saglabāšana un pirmā palaišana',section=SECTION_MANIFEST):
        root=clean_dir(work/'manifest_missing')
        only=root/'only.docx'; make_docx(only,'ONLY')
        project=Project(name='Missing',output_path=str(work/'missing_out.docx'))
        previous=build_manifest(project,preflight_all(_items([only])))
        diff=compare_manifest(previous,build_manifest(project,preflight_all(_items([]))))
        check(diff['missing']==1,f"MISSING: {diff['missing']}")
        path=work/'manifest.json'
        check(load_manifest(path) is None,'Neesošam manifestam jāatgriež None')
        path.write_text('{bojāts',encoding='utf-8')
        check(load_manifest(path) is None,'Bojātam manifestam jāatgriež None')
        save_manifest(path,previous)
        check(load_manifest(path)['document_count']==1,'Manifests netika nolasīts')
        first_run=compare_manifest(None,previous)
        check(first_run['status']=='FIRST_RUN' and first_run['added']==1,f'Pirmā palaišana: {first_run}')


def add_report_tests(suite:Suite,work):
    with suite.test('result.json atskaite',section=SECTION_REPORT):
        path=work/'result.json'
        write_result(path,status='DONE',total=3,merged=3,errors=0,warnings=0,output_path=str(work/'out.docx'))
        data=json.loads(path.read_text(encoding='utf-8'))
        check(data['status']=='DONE','Statuss nav ierakstīts')
        check(data['total']==3 and data['merged']==3,'Skaitļi nav ierakstīti')
        check('timestamp' in data,'Trūkst timestamp')
    with suite.test('kļūdu reportsaglabā pilnu tekstu un traceback',section=SECTION_REPORT):
        from docmerge.reporting.error_report import capture_exception, write_error_log
        try:
            raise ValueError('tests kļūda ar specifisku tekstu')
        except ValueError as exc:
            report=capture_exception(exc,stage=MergeStage.INTERNAL,error_code=ErrorCode.E_INTERNAL)
        check(report.text.strip() and 'specifisku tekstu' in report.text,'Kļūdas teksts nav saglabāts')
        check(report.traceback_text and 'ValueError' in report.traceback_text,'Traceback nav saglabāts')
        log_path=write_error_log(report,log_path=work/'errors.jsonl')
        check(log_path and log_path.is_file(),'Kļūdu log netika ierakstīts')
        entry=json.loads(log_path.read_text(encoding='utf-8').strip().splitlines()[0])
        check(entry['stage']==MergeStage.INTERNAL.value,'Stage nav saglabāts')
        check('ValueError' in (entry['traceback_text'] or ''),'Traceback nav JSONL failā')


def build_quick_suite(report_dir='reports',work_dir=None):
    """Izveido un izpilda QUICK komplektu (bez Word)."""
    suite=Suite('OFFLINE_QUICK',report_dir=report_dir)
    work=clean_dir(work_dir or (Path('runtime')/'_offline_quick'))
    suite.meta['work_dir']=str(work)
    add_import_and_config_tests(suite)
    add_project_and_scanner_tests(suite,work)
    add_sort_tests(suite,work)
    add_csv_tests(suite,work)
    add_json_tests(suite,work)
    add_matching_tests(suite,work)
    add_strict_tests(suite,work)
    add_paths_tests(suite,work)
    add_unlock_tests(suite,work)
    add_manifest_tests(suite,work)
    add_report_tests(suite,work)
    return suite
