"""OFFLINE FULL testi: apjomi, nosaukumi, fault injection, Word E2E, esošā Word sesija.

Izņēmums: reāls Microsoft Word COM ir atļauts (lokāls, nav nepieciešams tīkls).
Visi pārējie resursi ir lokāli; tīkls tiek bloķēts ar `guard.py`.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

from docmerge.core import order_list as ol
from docmerge.core.merge_service import MergeService
from docmerge.core.preflight import preflight_all
from docmerge.core.scanner import scan_folder
from docmerge.core.sorter import sort_items
from docmerge.domain.enums import OrderListStatus, SortMode
from docmerge.domain.errors import ErrorCode, MergeError, MergeStage, OrderListError, WordUnavailableError
from docmerge.domain.models import DocumentItem, Project
from docmerge.engines import word_com
from docmerge.reporting.logger import create_logger
from docmerge.testing import quick as quick_tests
from docmerge.testing.fakes import (FakeWord, fake_com_environment, word_unavailable_client)
from docmerge.testing.fixtures import (WordFixtureMaker, can_open_for_write, clean_dir, docx_text, hashes_of,
                                       make_docx, make_read_only, make_writable, page_break_count,
                                       paragraphs_of, sha256, winword_pids, write_zone_identifier,
                                       has_zone_identifier)
from docmerge.testing.harness import Suite, check, skip, warn

SECTION_LIST='order-list'
SECTION_SCALING='scaling'
SECTION_NAMES='filenames'
SECTION_FAULTS='fault-injection'
SECTION_WORD='word-e2e'
SECTION_SESSION='word-session'
PROJECT_ROOT=Path(__file__).resolve().parent.parent.parent
RUNTIME=PROJECT_ROOT/'runtime'
LOGS=RUNTIME/'_logs'


def logger():
    LOGS.mkdir(parents=True,exist_ok=True)
    log,_=create_logger(log_dir=str(LOGS)); return log


def service():
    return MergeService(logger(),error_log_path=LOGS/'errors.jsonl')


def items_of(paths):
    items=[DocumentItem(str(p)) for p in paths]
    for item in items: item.hydrate_from_path()
    return preflight_all(items)


def project_of(paths,out,order_entries=None,strict=False,order_path=None):
    project=Project(name='FULL',items=items_of(paths),output_path=str(out))
    if order_entries is not None:
        entries=[ol.OrderEntry(order=i,requested=str(r)) for i,r in enumerate(order_entries,1)]
        matches=ol.match_entries(entries,project.items)
        ol.apply_order(project,entries,matches,strict=strict,order_path=str(order_path or ''))
    return project


def marker_positions(out,markers):
    """Atgriež (positions, paragraphs) vai (None, paragraphs), ja kāds marķieris trūkst."""
    paragraphs=paragraphs_of(out)
    positions=[]
    for marker in markers:
        try: positions.append(paragraphs.index(marker))
        except ValueError: return None,paragraphs
    return positions,paragraphs


def assert_ordered(suite,out,markers,label=''):
    positions,paragraphs=marker_positions(out,markers)
    check(positions is not None,
          f'{label} Izvadē nav visu marķieru. Marķieri: {markers}. Paragrāfi: {paragraphs[:12]}')
    check(positions==sorted(positions),f'{label} Secība nav pareiza: {positions} (paragrāfi: {paragraphs[:12]})')
    return positions,paragraphs


def no_leftover_word(baseline,label=''):
    """Pārbauda, ka pēc mūsu darba nav palicis jauns WINWORD process."""
    after=winword_pids()
    new=sorted(after-set(baseline))
    check(not new,f'{label} Palika jauni WINWORD procesi: {new} (pirms: {sorted(baseline)})')
    return after


def runtime_temp_leftovers(base=RUNTIME):
    """Temp faili runtime mapē, kuriem nevajadzētu palikt."""
    found=[]
    for pattern in ('*.tmp','*.tmp.docx','*.tmp.doc'):
        found.extend(base.rglob(pattern))
    return sorted(str(p) for p in found)


def temp_leftover_baseline():
    """Temp failu stāvoklis PIRMS testa (vecie diag atlikumi netiek skaitīti)."""
    return set(runtime_temp_leftovers())


def assert_no_new_temp_files(baseline,label=''):
    """Pārbauda, ka šī testa laikā nav radušies jauni temp faili."""
    new=sorted(set(runtime_temp_leftovers())-set(baseline or set()))
    check(not new,f'{label} Radās jauni temp faili: {new}')


def add_scaling_tests(suite:Suite,work,word_available):
    """1/2/10/50/100 ģenerēti DOCX + mixed DOC/DOCX."""
    for count in (1,2,10,50,100):
        with suite.test(f'Word merge: {count} DOCX',section=SECTION_SCALING) as case:
            if not word_available: skip('Word COM nav pieejams')
            root=clean_dir(work/f'scale_{count}')
            paths=[]
            for i in range(1,count+1):
                path=root/f'doc_{i:03d}.docx'
                make_docx(path,f'MARKERIS {i:03d}')
                paths.append(path)
            before=hashes_of(paths)
            baseline=winword_pids(); temp_before=temp_leftover_baseline()
            out=root/'APVIENOTS.docx'
            project=project_of(paths,out)
            result=service().run(project)
            check(result['merged']==count,f"Apvienoti {result['merged']}/{count}; kļūdas: {result['errors'][:2]}")
            check(len(result['errors'])==0,f'Kļūdas: {result["errors"][:2]}')
            check(out.is_file(),'Izvades fails nav izveidots')
            texts=docx_text(out)
            for i in range(1,count+1):
                check(f'MARKERIS {i:03d}' in texts,f'Trūkst marķiera {i:03d}')
            assert_ordered(suite,out,[f'MARKERIS {i:03d}' for i in range(1,count+1)],f'{count} faili:')
            breaks=page_break_count(out)
            check(breaks==count-1,f'Lappuses pārtraukumi: {breaks}, gaidīti {count-1}')
            check(hashes_of(paths)==before,'Avotu hash mainījās (source ir immutable)')
            assert_no_new_temp_files(temp_before,f'{count} faili:')
            no_leftover_word(baseline,f'{count} faili:')

    with suite.test('Word merge: mixed DOC + DOCX (Word fixtures)',section=SECTION_SCALING):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'mixed')
        baseline=winword_pids()
        with WordFixtureMaker() as maker:
            maker.make(root/'01_docx.docx','MARKERIS DOCX 1',maker.WD_FORMAT_DOCX)
            maker.make(root/'02_doc.doc','MARKERIS DOC 2',maker.WD_FORMAT_DOCUMENT)
            maker.make(root/'03_docx.docx','MARKERIS DOCX 3',maker.WD_FORMAT_DOCX)
        paths=sorted(root.glob('*'))
        before=hashes_of(paths)
        out=root/'MIXED.docx'
        result=service().run(project_of(paths,out))
        check(result['merged']==3,f"Apvienoti {result['merged']}/3: {result['errors'][:2]}")
        assert_ordered(suite,out,['MARKERIS DOCX 1','MARKERIS DOC 2','MARKERIS DOCX 3'],'Mixed:')
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        no_leftover_word(baseline,'Mixed:')

    with suite.test('Word merge: avotam Zone.Identifier (Protected View)',section=SECTION_SCALING):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'zone')
        baseline=winword_pids(); temp_before=temp_leftover_baseline()
        with WordFixtureMaker() as maker:
            maker.make(root/'01_alpha.docx','MARKERIS ALFA')
            maker.make(root/'02_beta.docx','MARKERIS BETA')
            maker.make(root/'03_gamma.docx','MARKERIS GAMMA')
        marked=root/'03_gamma.docx'
        write_zone_identifier(marked)
        paths=[root/'01_alpha.docx',root/'02_beta.docx',root/'03_gamma.docx']
        before=hashes_of(paths)
        out=root/'ZONE.docx'
        project=project_of(paths,out)
        result=service().run(project)
        check(result['merged']==3,f"Apvienoti {result['merged']}/3: {result['errors'][:2]}")
        check(result['unlocked']==[str(marked)],f"Atbloķētās kopijas: {result['unlocked']}")
        assert_ordered(suite,out,['MARKERIS ALFA','MARKERIS BETA','MARKERIS GAMMA'],'Zone:')
        check(has_zone_identifier(out) is False,'Izvadei ir Zone.Identifier marķējums')
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        check(has_zone_identifier(marked) is True,'Avota marķējums tika noņemts (avots ir immutable)')
        check(list((RUNTIME/'unlocked').glob('*'))==[],'Palika atbloķētās kopijas')
        assert_no_new_temp_files(temp_before,'Zone:')
        no_leftover_word(baseline,'Zone:')


def add_word_e2e_tests(suite:Suite,work,word_available):
    with suite.test('Word E2E: JSON secība C, A, B (absolūtā)',section=SECTION_WORD):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'e2e_order')
        baseline=winword_pids()
        with WordFixtureMaker() as maker:
            maker.make(root/'A_alpha.docx','MARKERIS A ALFA')
            maker.make(root/'B_beta.docx','MARKERIS B BETA')
            maker.make(root/'C_gamma.docx','MARKERIS C GAMMA')
        order_path=root/'order.json'
        order_path.write_text(json.dumps(['C_gamma.docx','A_alpha.docx','B_beta.docx'],ensure_ascii=False),encoding='utf-8')
        entries,meta=ol.load_order_list(order_path)
        check(meta['format']=='JSON','Order fails nav JSON')
        paths=sorted(root.glob('*_*.docx'))
        before=hashes_of(paths)
        out=root/'ORDERED.docx'
        project=project_of(paths,out,order_entries=[e.requested for e in entries],strict=True,order_path=order_path)
        check(project.items[0].filename=='C_gamma.docx',f'Pirmais fails: {project.items[0].filename}')
        result=service().run(project)
        check(result['merged']==3,f"Apvienoti {result['merged']}/3: {result['errors'][:2]}")
        check(result['word_owned'] is True,'Word sesija netika atzīta par mūsu (owned=False)')
        positions,paragraphs=assert_ordered(suite,out,['MARKERIS C GAMMA','MARKERIS A ALFA','MARKERIS B BETA'],'E2E:')
        check(positions[0]<positions[1]<positions[2],f'Secība nav C, A, B: {positions}')
        titles=[name for name in ('C_gamma','A_alpha','B_beta') if name in paragraphs]
        check(len(titles)==3,f'Virsraksti nav visi: {titles} (paragrāfi: {paragraphs[:10]})')
        check(page_break_count(out)==2,f"Lappuses pārtraukumi: {page_break_count(out)}")
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        check(has_zone_identifier(out) is False,'Izvadei ir Zone.Identifier')
        no_leftover_word(baseline,'E2E:')

    with suite.test('Word E2E: idempotents rerun un backup',section=SECTION_WORD):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'e2e_rerun')
        baseline=winword_pids()
        with WordFixtureMaker() as maker:
            maker.make(root/'01_first.docx','MARKERIS PIRMAIS')
            maker.make(root/'02_second.docx','MARKERIS O TRAIS')
        out=root/'RERUN.docx'
        entries=['02_second.docx','01_first.docx']
        project=project_of(sorted(root.glob('*.docx')),out,order_entries=entries,strict=True,order_path=root/'order.json')
        first=service().run(project)
        check(first['merged']==2,f"Pirmais cikls: {first['merged']}")
        check(out.is_file(),'Izvades fails nav izveidots')
        project2=project_of(sorted(root.glob('*.docx')),out,order_entries=entries,strict=True,order_path=root/'order.json')
        second=service().run(project2)
        check(second['merged']==2,f"Otrais cikls: {second['merged']}")
        backup=Path(str(out)+'.bak')
        check(backup.is_file(),'Backup (.bak) netika izveidots, lai gan izvade jau eksistēja')
        check(second['manifest_diff']['unchanged']==2,f"UNCHANGED: {second['manifest_diff']}")
        assert_ordered(suite,out,['MARKERIS O TRAIS','MARKERIS PIRMAIS'],'Rerun:')
        no_leftover_word(baseline,'Rerun:')


def add_name_tests(suite:Suite,work,word_available):
    LATVIAN='ĀČĒĢĶĻŅŠŪŽ'
    with suite.test('Failu nosaukumi: latviešu unicode burti',section=SECTION_NAMES):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'names_unicode')
        baseline=winword_pids()
        paths=[]
        for letter in LATVIAN:
            path=root/f'{letter}_dokuments.docx'
            make_docx(path,f'MARKERIS {letter}')
            paths.append(path)
        before=hashes_of(paths)
        out=root/'UNICODE.docx'
        result=service().run(project_of(paths,out))
        check(result['merged']==len(LATVIAN),f"Apvienoti {result['merged']}/{len(LATVIAN)}: {result['errors'][:2]}")
        assert_ordered(suite,out,[f'MARKERIS {letter}' for letter in LATVIAN],'Unicode:')
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        no_leftover_word(baseline,'Unicode:')
    with suite.test('Failu nosaukumi: atstarpes, iekavas, gari nosaukumi',section=SECTION_NAMES):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'names_tricky')
        baseline=winword_pids()
        paths=[
            make_docx(root/'01 ar atstarpēm.docx','MARKERIS ATSTRAPES'),
            make_docx(root/'02 (iekavas) un figūras.docx','MARKERIS IEKAVAS'),
            make_docx(root/f'{"X"*110}.docx','MARKERIS GARAIS'),
            make_docx(root/"04_apostrofs 'un' pedinas.docx",'MARKERIS PEDINAS'),
        ]
        out=root/'TRICKY.docx'
        result=service().run(project_of(paths,out))
        check(result['merged']==4,f"Apvienoti {result['merged']}/4: {result['errors'][:2]}")
        # Natural secība: '01 ...' < '02 ...' < '04_...' < 'XXXX...' 
        assert_ordered(suite,out,['MARKERIS ATSTRAPES','MARKERIS IEKAVAS','MARKERIS PEDINAS','MARKERIS GARAIS'],'Tricky:')
        check('MARKERIS GARAIS' in docx_text(out),'Garā nosaukuma saturs pazuda')
        no_leftover_word(baseline,'Tricky:')
    with suite.test('Apakšmapes un dublikāti dažādās mapēs',section=SECTION_NAMES):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'names_nested')
        baseline=winword_pids()
        (root/'pirma').mkdir(parents=True,exist_ok=True); (root/'otra').mkdir(parents=True,exist_ok=True)
        first=make_docx(root/'pirma'/'same.docx','MARKERIS PIRMA MAPE')
        second=make_docx(root/'otra'/'same.docx','MARKERIS OTRA MAPE')
        out=root/'NESTED.docx'
        project=project_of([first,second],out,order_entries=['otra/same.docx','pirma/same.docx'],strict=True,
                           order_path=root/'order.csv')
        result=service().run(project)
        check(result['merged']==2,f"Apvienoti {result['merged']}/2: {result['errors'][:2]}")
        assert_ordered(suite,out,['MARKERIS OTRA MAPE','MARKERIS PIRMA MAPE'],'Nested:')
        no_leftover_word(baseline,'Nested:')
    with suite.test('Read-only avots tiek apvienots',section=SECTION_NAMES):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'names_readonly')
        baseline=winword_pids()
        path=make_docx(root/'readonly.docx','MARKERIS READONLY')
        make_read_only(path)
        before=sha256(path)
        # Izvades nosaukums NAV tāds pats kā avotam (Windows ceļi nav reģistrjutīgi).
        out=root/'APVIENOTS_readonly.docx'
        try:
            result=service().run(project_of([path],out))
            check(result['merged']==1,f"Apvienoti {result['merged']}/1: {result['errors'][:2]}")
            assert_ordered(suite,out,['MARKERIS READONLY'],'Read-only:')
            check(sha256(path)==before,'Read-only avots tika mainīts')
        finally:
            make_writable(path)
        no_leftover_word(baseline,'Read-only:')
    with suite.test('Tukša mape: nav ko apvienot un kļūda netiek pazaudēta',section=SECTION_NAMES):
        root=clean_dir(work/'names_empty')
        empty=scan_folder(str(root),recursive=True)
        check(empty==[],'Tukšā mapē tika atrasti faili')
        project=Project(name='Empty',items=[],output_path=str(root/'EMPTY.docx'))
        try: service().run(project)
        except MergeError as exc:
            check(exc.error_code==ErrorCode.E_NO_DOCUMENTS,f'Kods: {exc.error_code}')
            check(exc.message.strip(),'Kļūdas teksts ir tukšs')
        else: raise AssertionError('Tukšai mapi bija jāmet MergeError (E_NO_DOCUMENTS)')


def add_order_list_tests(suite:Suite,work,word_available):
    with suite.test('Saraksts: trūkstošs ieraksts (bez STRICT) tiek izlaists',section=SECTION_LIST):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'list_missing')
        baseline=winword_pids()
        paths=[make_docx(root/f'{i}_doc.docx',f'MARKERIS LIST {i}') for i in (1,2,3)]
        order=root/'order.csv'
        order.write_text('order,filename\n1,1_doc.docx\n2,neeksiste_999.docx\n3,3_doc.docx\n',encoding='utf-8')
        entries,meta=ol.load_order_list(order)
        out=root/'LIST_MISSING.docx'
        project=project_of(paths,out,order_entries=[e.requested for e in entries],strict=False,order_path=order)
        result=service().run(project)
        check(result['merged']==2,f"Apvienoti {result['merged']}/2: {result['errors'][:2]}")
        check(result['order_list']['summary']['missing']==1,'Trūkstošais ieraksts nav reportā')
        assert_ordered(suite,out,['MARKERIS LIST 1','MARKERIS LIST 3'],'Missing:')
        no_leftover_word(baseline,'Missing:')
    with suite.test('Saraksts: dublikāts CSV (bez STRICT) tiek izlaists',section=SECTION_LIST):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'list_duplicate')
        baseline=winword_pids()
        paths=[make_docx(root/f'{i}_doc.docx',f'MARKERIS DUP {i}') for i in (1,2)]
        order=root/'order.csv'
        order.write_text('filename\n1_doc.docx\n1_doc.docx\n2_doc.docx\n',encoding='utf-8')
        entries,_=ol.load_order_list(order)
        out=root/'LIST_DUP.docx'
        project=project_of(paths,out,order_entries=[e.requested for e in entries],strict=False,order_path=order)
        result=service().run(project)
        check(result['merged']==2,f"Apvienoti {result['merged']}/2: {result['errors'][:2]}")
        check(result['order_list']['summary']['duplicate']==1,f"DUPLICATE: {result['order_list']['summary']}")
        assert_ordered(suite,out,['MARKERIS DUP 1','MARKERIS DUP 2'],'Duplicate:')
        no_leftover_word(baseline,'Duplicate:')
    with suite.test('Saraksts: exportētā secība atkārto merge identiski',section=SECTION_LIST):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'list_roundtrip')
        baseline=winword_pids()
        paths=[make_docx(root/f'{name}.docx',f'MARKERIS RT {name}') for name in ('a','b','c')]
        out=root/'RT.docx'
        project=project_of(paths,out)
        project.items=sort_items(project.items,SortMode.MANUAL)
        project.items[0],project.items[2]=project.items[2],project.items[0]
        for index,item in enumerate(project.items,1): item.manual_order=index
        exported,fmt=ol.export_order(root/'exported.csv',project.items)
        entries,meta=ol.load_order_list(exported)
        out2=root/'RT2.docx'
        project2=project_of(paths,out2,order_entries=[e.requested for e in entries],strict=True,order_path=exported)
        result=service().run(project2)
        check(result['merged']==3,f"Apvienoti {result['merged']}/3: {result['errors'][:2]}")
        expected=[Path(x.source_path).stem for x in project.items]
        assert_ordered(suite,out2,[f'MARKERIS RT {name}' for name in expected],'Roundtrip:')
        check(expected==['c','b','a'],f'Eksportētā secība: {expected}')
        no_leftover_word(baseline,'Roundtrip:')
    with suite.test('Saraksts: bojāts CSV — derīgie ieraksti paliek, problēmas netiek pazaudētas',section=SECTION_LIST):
        raw='order,filename\n1,ok_1.docx\nnevalidigs\n2,ok_2.docx\n\n'
        entries,problems=ol.parse_csv_entries(raw)
        check([e.requested for e in entries]==['ok_1.docx','nevalidigs','ok_2.docx'],f'Ieraksti: {[e.requested for e in entries]}')
        check(any('nederīgs order' in p for p in problems),f'Problēmas: {problems}')
    with suite.test('Saraksts: bojāts JSON met kļūdu, kas tiek arī ierakstīta logā',section=SECTION_LIST):
        root=clean_dir(work/'list_bad_json')
        bad=root/'order.json'; bad.write_text('{"files":[{"order":1,"filename":"a.docx"',encoding='utf-8')
        error_path=root/'errors.jsonl'
        try:
            ol.load_order_list(bad)
        except OrderListError as exc:
            report=quick_capture(exc,error_path)
            check(report['message'],'Kļūdas ziņojums ir tukšs')
            check('Bojāts JSON' in report['message'],f"Ziņojums: {report['message']}")
            check(error_path.is_file(),'Kļūda netika ierakstīta errors.jsonl')
            check(bad.exists(),'Bojātais fails tika mainīts')
        else: raise AssertionError('Bojātam JSON sarakstam bija jāmet OrderListError')
    with suite.test('Saraksts: tukšs CSV noved pie E_NO_DOCUMENTS (kļūda netiek pazaudēta)',section=SECTION_LIST):
        root=clean_dir(work/'list_empty')
        root.joinpath('order.csv').write_text('',encoding='utf-8')
        entries,problems=ol.parse_csv_entries('')
        check(entries==[],'Tukšam CSV jābūt bez ierakstiem')
        check(problems and 'tukšs' in problems[0].lower(),f'Problēmas: {problems}')
        project=Project(name='EmptyList',items=[],output_path=str(root/'EMPTY_LIST.docx'))
        try: service().run(project)
        except MergeError as exc: check(exc.error_code==ErrorCode.E_NO_DOCUMENTS,f'Kods: {exc.error_code}')
        else: raise AssertionError('Tukšam sarakstam bija jāmet E_NO_DOCUMENTS')


def quick_capture(exc,error_path):
    """Materializē kļūdu un ieraksta to errors.jsonl (kā to dara GUI/CLI)."""
    from docmerge.reporting.error_report import capture_exception, persist_error_log_safely
    report=capture_exception(exc,stage=MergeStage.ORDER_LIST)
    persist_error_log_safely(report,log_path=error_path,logger=logger())
    return report.to_dict()


def add_fault_tests(suite:Suite,work,word_available):
    with suite.test('Fault: bojāts DOCX ZIP (kļūda paliek preflight datos)',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_broken_zip')
        broken=root/'broken.docx'; broken.write_bytes(b'PK\x03\x04 not a real zip archive')
        good=make_docx(root/'good.docx','MARKERIS BROKEN ZIP')
        project=project_of([good,broken],root/'OUT.docx')
        broken_item=[x for x in project.items if x.filename=='broken.docx'][0]
        check(broken_item.errors,f'Bojāts DOCX netika atzīmēts: {broken_item.errors}')
        check(not broken_item.eligible_for_merge,'Bojāts DOCX nedrīkst būt eligible')
        if not word_available: skip('Word COM nav pieejams')
        result=service().run(project)
        check(result['merged']==1,f"Apvienoti {result['merged']}/1: {result['errors'][:2]}")
        check(len(project.items)==2,'Preflight dati tika pazaudēti')
    with suite.test('Fault: 0 B DOCX un ~$ Word fails',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_zero')
        zero=root/'zero.docx'; zero.write_bytes(b'')
        owner=root/'~$temp.docx'; owner.write_bytes(b'\x15Microsoft Word')
        scanned=scan_folder(str(root),recursive=True)
        check([x.filename for x in scanned]==['zero.docx'],f'~$ fails netika izlaists: {[x.filename for x in scanned]}')
        project=project_of([zero,owner],root/'OUT.docx')
        check(project.items[0].errors and 'tukšs' in project.items[0].errors[0],f'0 B kļūda: {project.items[0].errors}')
        check(project.items[1].errors,'~$ fails netika atzīmēts kā kļūda')
        check(all(not x.eligible_for_merge for x in project.items),'Bojātie faili nedrīkst būt eligible')
    with suite.test('Fault: aizņemts izvades fails -> E_OUTPUT_LOCKED',section=SECTION_FAULTS):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'fault_locked')
        good=make_docx(root/'good.docx','MARKERIS LOCKED OUT')
        out=root/'LOCKED.docx'; out.write_bytes(b'PK\x03\x04existing output')
        handle=open(out,'r+b')
        try:
            try: service().run(project_of([good],out))
            except MergeError as exc:
                check(exc.error_code==ErrorCode.E_OUTPUT_LOCKED,f'Kods: {exc.error_code}')
                check('aizņemts' in exc.message.lower() or 'Aizver' in exc.full_text,f'Kļūdas teksts: {exc.message}')
            else: raise AssertionError('Aizņemtam izvades failam bija jāmet MergeError')
        finally:
            handle.close()
    with suite.test('Fault: izvade nedrīkst būt viens no avotiem (datu drošība)',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_output_is_source')
        source=make_docx(root/'avots.docx','MARKERIS AVOTS')
        before=sha256(source)
        project=project_of([source],source)
        try:
            service().run(project)
        except MergeError as exc:
            check(exc.error_code==ErrorCode.E_OUTPUT_IS_SOURCE,f'Kods: {exc.error_code}')
            check('avota failiem' in exc.message or 'avota' in exc.message,f'Kļūdas teksts: {exc.message}')
        else:
            raise AssertionError('Izvadei, kas ir avots, bija jāmet MergeError')
        check(sha256(source)==before,'AVOTA FAILS TIKA PĀRRAKSTĪTS!')
        check(docx_text(source).strip()=='MARKERIS AVOTS','Avota saturs mainījās')
    with suite.test('Fault: avots pazūd pēc skenēšanas (kļūda tiek reportēta)',section=SECTION_FAULTS):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'fault_missing_after_scan')
        first=make_docx(root/'01_ok.docx','MARKERIS OK PIRMS')
        second=make_docx(root/'02_pazudis.docx','MARKERIS PAZUDIS')
        project=project_of([first,second],root/'OUT.docx')
        second.unlink()
        result=service().run(project)
        check(result['merged']==1,f"Apvienoti {result['merged']}/1: {result['errors'][:2]}")
        check(len(result['errors'])==1,f"Kļūdu skaits: {len(result['errors'])}")
        entry=result['errors'][0]
        check(entry['error_code']==ErrorCode.E_INVALID_PATH.value,f"Kods: {entry['error_code']}")
        check('neeksistē' in entry['error'],f"Teksts: {entry['error']}")
    with suite.test('Fault: dublikāti order numuri -> DUPLICATE + STRICT bloķē',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_dup_order')
        paths=[make_docx(root/'d1.docx','D1'),make_docx(root/'d2.docx','D2')]
        items=items_of(paths)
        entries,_=ol.parse_json_entries('[{"order":1,"filename":"d1.docx"},{"order":1,"filename":"d2.docx"}]')
        matches=ol.match_entries(entries,items)
        check(matches[1].status==OrderListStatus.DUPLICATE,f'Statuss: {matches[1].status}')
        try: ol.enforce_strict(matches,True)
        except Exception as exc: check('STRICT MODE' in str(exc),'STRICT nekļuva par bloķējošu')
        else: raise AssertionError('Atkārtots order numurs STRICT režīmā bija jābloķē')
    with suite.test('Fault: Word nav pieejams (mock) -> WordUnavailableError',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_no_word')
        good=make_docx(root/'good.docx','MARKERIS NO WORD')
        with fake_com_environment(client=word_unavailable_client(),winword_pids=set()):
            try: service().run(project_of([good],root/'OUT.docx'))
            except WordUnavailableError as exc:
                check(exc.error_code==ErrorCode.E_WORD_UNAVAILABLE,f'Kods: {exc.error_code}')
                check('Server execution failed' in exc.full_text,f'Kļūdas teksts: {exc.full_text[:200]}')
            else: raise AssertionError('Bez Word bija jāmet WordUnavailableError')
    with suite.test('Fault: InsertFile simulēta kļūda (pilns COM teksts)',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_insertfile')
        good=make_docx(root/'good.docx','MARKERIS INSERT FAIL')
        word=FakeWord(); word.fail_all_inserts=True
        with fake_com_environment(word=word,winword_pids=set()):
            result=service().run(project_of([good],root/'OUT.docx'))
        check(result['merged']==0,f"Apvienoti {result['merged']}")
        entry=result['errors'][0]
        check(entry['error_code']==ErrorCode.E_INSERT_FILE.value,f"Kods: {entry['error_code']}")
        check('24753' in (entry['repr'] or '') or 'not valid' in entry['error'],'COM kļūdas teksts pazuda')
        check(entry['traceback'],'Trūkst traceback')
        check(entry['error']!='None','Kļūdas teksts ir None')
    with suite.test('Fault: SaveAs2 simulēta kļūda -> E_SAVE_AS',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_saveas')
        good=make_docx(root/'good.docx','MARKERIS SAVEAS FAIL')
        word=FakeWord(); word.saveas_failures=99
        with fake_com_environment(word=word,winword_pids=set()):
            try: service().run(project_of([good],root/'OUT.docx'))
            except MergeError as exc:
                check(exc.error_code==ErrorCode.E_SAVE_AS,f'Kods: {exc.error_code}')
                check('locked for editing' in exc.full_text or 'WRD6ER32' in exc.full_text,f'Kļūdas teksts: {exc.message}')
            else: raise AssertionError('SaveAs2 kļūdai bija jāmet MergeError')
    with suite.test('Fault: normalizācijas kļūda reportē abas kļūdas',section=SECTION_FAULTS):
        root=clean_dir(work/'fault_normalize')
        good=make_docx(root/'good.docx','MARKERIS NORMALIZE FAIL')
        word=FakeWord(); word.fail_all_inserts=True; word.open_failures=99
        with fake_com_environment(word=word,winword_pids=set()):
            result=service().run(project_of([good],root/'OUT.docx'))
        check(result['merged']==0,f"Apvienoti {result['merged']}")
        entry=result['errors'][0]
        check('InsertFile' in entry['error'] and 'normalizācija' in entry['error'],f"Teksts: {entry['error']}")
        check(entry['used_normalized'] is True,'used_normalized nav atzīmēts')
        # Veiksmīgi normalizēto failu saraksts paliek tukšs (normalizācija neizdevās),
        # bet kļūda ir pilnībā reportā.
        check(result['normalized']==[],f"normalized: {result['normalized']}")


def _merge_subprocess(source_dir,out,force_dispatch=False,timeout=420):
    env=dict(os.environ)
    if force_dispatch: env[word_com.DISPATCH_METHODS_ENV]='Dispatch'
    return subprocess.run([sys.executable,'-m','docmerge.cli.main','--merge-all',str(source_dir),'--output',str(out)],
                          capture_output=True,text=True,encoding='utf-8',errors='replace',
                          cwd=str(PROJECT_ROOT),env=env,timeout=timeout)


def add_session_tests(suite:Suite,work,word_available):
    with suite.test('Word sesija: Dispatch fallback NEAIZVĒR lietotāja sesiju',section=SECTION_SESSION):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'session_dispatch_fallback')
        baseline=winword_pids()
        with WordFixtureMaker() as maker:
            user_path=maker.make(root/'user_doc.docx','LIETOTAJA SESIJAS DOKUMENTS')
            user_doc=maker.open_document(user_path)
            check(user_doc is not None,'Lietotāja dokuments netika atvērts')
            sources=root/'src'; sources.mkdir(parents=True,exist_ok=True)
            make_docx(sources/'01_doc.docx','MARKERIS SESSION 1')
            make_docx(sources/'02_doc.docx','MARKERIS SESSION 2')
            out=root/'SESSION_MERGE.docx'
            cp=_merge_subprocess(sources,out,force_dispatch=True)
            check(cp.returncode==0,f'Merge rc={cp.returncode}: {(cp.stdout or "")[-500:]} {(cp.stderr or "")[-300:]}')
            check(out.is_file(),'Merge izvade nav izveidota')
            check('MARKERIS SESSION 1' in docx_text(out),'Izvadē trūkst pirmā dokumenta satura')
            count=maker.word.Documents.Count
            names=[d.Name for d in maker.word.Documents]
            check(count>=1,f'Lietotāja sesijā vairs nav neviena dokumenta (Count={count})')
            check(any('user_doc' in (name or '') for name in names),
                  f'LIETOTĀJA DOKUMENTS TIKA AIZVĒRTS! Atlikušie: {names}')
            check('user_doc' in (user_doc.Name or ''),'Lietotāja dokumenta COM atsauce vairs nav derīga')
        no_leftover_word(baseline,'Session fallback:')
    with suite.test('Word sesija: DispatchEx ceļš lietotāja sesiju neskar',section=SECTION_SESSION):
        if not word_available: skip('Word COM nav pieejams')
        root=clean_dir(work/'session_dispatchex')
        baseline=winword_pids()
        with WordFixtureMaker() as maker:
            user_path=maker.make(root/'user_doc.docx','LIETOTAJA SESIJAS DOKUMENTS 2')
            user_doc=maker.open_document(user_path)
            sources=root/'src'; sources.mkdir(parents=True,exist_ok=True)
            make_docx(sources/'01_doc.docx','MARKERIS SESIJA A')
            out=root/'SESSION_MERGE2.docx'
            cp=_merge_subprocess(sources,out)
            check(cp.returncode==0,f'Merge rc={cp.returncode}: {(cp.stdout or "")[-500:]}')
            check(out.is_file(),'Merge izvade nav izveidota')
            names=[d.Name for d in maker.word.Documents]
            check(any('user_doc' in (name or '') for name in names),f'Lietotāja dokuments pazuda: {names}')
            check('user_doc' in (user_doc.Name or ''),'Lietotāja dokumenta COM atsauce nav derīga')
        no_leftover_word(baseline,'Session DispatchEx:')
    with suite.test('Word sesija: Quit tikai savai instancei (mock ownership)',section=SECTION_SESSION):
        from docmerge.testing.fakes import FakeWord, fake_com_environment
        import types
        saved_env=os.environ.get(word_com.DISPATCH_METHODS_ENV)
        os.environ[word_com.DISPATCH_METHODS_ENV]='Dispatch'
        try:
            user_word=FakeWord()
            client=types.SimpleNamespace(Dispatch=lambda name:user_word)
            with fake_com_environment(word=user_word,client=client,winword_pids={4321}):
                session=word_com.open_word_session(client)
                check(session.owned is False,'Pieslēgta lietotāja sesija tika atzīta par savu')
                check(session.close() is False,'Lietotāja sesijas close() nedrīkst atgriezt True')
                check(user_word.quit_calls==0,'LIETOTĀJA SESIJAI tika izsaukts Quit()')
            own_word=FakeWord()
            client2=types.SimpleNamespace(Dispatch=lambda name:own_word)
            with fake_com_environment(word=own_word,client=client2,winword_pids=set()):
                session=word_com.open_word_session(client2)
                check(session.owned is True,'Jauna Dispatch instance netika atzīta par savu')
                check(session.close() is True,'Savas instances close() jāatgriež True')
                check(own_word.quit_calls==1,f'Quit netika izsaukts savai instancei: {own_word.quit_calls}')
        finally:
            if saved_env is None: os.environ.pop(word_com.DISPATCH_METHODS_ENV,None)
            else: os.environ[word_com.DISPATCH_METHODS_ENV]=saved_env


def build_full_suite(report_dir='reports',work_dir=None,word_info=None):
    """Izveido un izpilda OFFLINE FULL komplektu (QUICK + apjomi + Word E2E + fault)."""
    word_info=word_info or {}
    word_available=bool(word_info.get('word_com'))
    suite=Suite('OFFLINE_FULL',report_dir=report_dir)
    work=clean_dir(work_dir or (RUNTIME/'_offline_full'))
    suite.meta.update({'work_dir':str(work),'word_available':word_available,
                       'word_version':word_info.get('word_version'),
                       'word_owned_probe':word_info.get('word_owned'),
                       'winword_leftover_probe':word_info.get('leftover_winword')})
    if not word_available:
        suite.note('Word COM nav pieejams — Word E2E testi tiks izlaisti (SKIP), pārējie testi darbojas.')
    # QUICK testi (imports, config, saraksti, matching, strict, manifest, atskaites)
    quick_tests.add_import_and_config_tests(suite)
    quick_tests.add_project_and_scanner_tests(suite,work)
    quick_tests.add_sort_tests(suite,work)
    quick_tests.add_csv_tests(suite,work)
    quick_tests.add_json_tests(suite,work)
    quick_tests.add_matching_tests(suite,work)
    quick_tests.add_strict_tests(suite,work)
    quick_tests.add_paths_tests(suite,work)
    quick_tests.add_unlock_tests(suite,work)
    quick_tests.add_manifest_tests(suite,work)
    quick_tests.add_report_tests(suite,work)
    # FULL papildinājumi
    add_order_list_tests(suite,work,word_available)
    add_scaling_tests(suite,work,word_available)
    add_name_tests(suite,work,word_available)
    add_fault_tests(suite,work,word_available)
    add_word_e2e_tests(suite,work,word_available)
    add_session_tests(suite,work,word_available)
    return suite
