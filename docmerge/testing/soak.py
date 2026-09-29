"""OFFLINE soak tests: 100–300 dokumenti, vairāki merge cikli, bez lietotāja iesaistes.

Mērķis ir reāla slodze, nevis gaidīšana: cikli tiek pabeigti, cik ātri vien iespējams.
Pēc katra cikla: izvade eksistē, dokumentu skaits, marķieri, avotu hash nemainīgs,
temp iztīrīts, faili nav aizņemti, atmiņas/process stāvoklis, Word notīrīts.
Viena cikla kļūda tiek precīzi ierakstīta, un pārējie cikli turpinās.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

from docmerge.core import order_list as ol
from docmerge.domain.models import Project
from docmerge.testing.full import (RUNTIME, assert_no_new_temp_files, assert_ordered, items_of, no_leftover_word,
                                   project_of, runtime_temp_leftovers, service, temp_leftover_baseline)
from docmerge.testing.fixtures import (can_open_for_write, clean_dir, hashes_of, make_docx, object_count,
                                       page_break_count, process_memory_mb, sha256, winword_pids)
from docmerge.testing.harness import Suite, check, skip, warn

SOAK_SETUP='soak-setup'; SOAK_LISTS='soak-lists'; SOAK_CYCLES='soak-cycles'
DEFAULT_DOCS=120
MAX_DOCS=300


def build_soak_suite(report_dir='reports',work_dir=None,word_info=None,doc_count=None):
    """Izveido un izpilda soak komplektu (OFFLINE, bez lietotāja iesaistes)."""
    word_info=word_info or {}
    word_available=bool(word_info.get('word_com'))
    raw=doc_count or os.environ.get('DOC_MERGE_SOAK_DOCS') or DEFAULT_DOCS
    count=max(1,min(int(raw),MAX_DOCS))
    suite=Suite('SOAK',report_dir=report_dir)
    work=clean_dir(work_dir or (RUNTIME/'_soak'))
    sources=clean_dir(work/'sources')
    suite.meta.update({'doc_count':count,'word_available':word_available,
                       'word_version':word_info.get('word_version'),'work_dir':str(work)})
    markers=[f'SOAK MARKER {i:04d}' for i in range(1,count+1)]
    paths=[sources/f'doc_{i:04d}.docx' for i in range(1,count+1)]

    with suite.test(f'soak: sagatavo {count} DOCX',section=SOAK_SETUP):
        for index,path in enumerate(paths):
            make_docx(path,markers[index])
        check(len(list(sources.glob('*.docx')))==count,f'Gaidīti {count} faili')
    before=hashes_of(paths)

    with suite.test('soak: liela saraksta CSV/JSON apstrāde un eksports',section=SOAK_LISTS):
        project=Project(name='SOAK',items=items_of(paths))
        reversed_paths=list(reversed(paths))
        csv_path=work/'soak_order.csv'
        rows=['order,filename']+[f'{i},{p.name}' for i,p in enumerate(reversed_paths,1)]
        csv_path.write_text('\n'.join(rows)+'\n',encoding='utf-8')
        entries,_=ol.load_order_list(csv_path)
        check(len(entries)==count,f'CSV ieraksti: {len(entries)}')
        summary=ol.summarize(ol.match_entries(entries,project.items))
        check(summary['found']==count,f'CSV matching: {summary}')
        json_path=work/'soak_order.json'
        json_path.write_text(json.dumps([p.name for p in reversed_paths],ensure_ascii=False),encoding='utf-8')
        check(len(ol.load_order_list(json_path)[0])==count,'JSON ierakstu skaits kļūdains')
        exported,_=ol.export_order(work/'soak_export.json',project.items)
        replay,_=ol.load_order_list(exported)
        replay_summary=ol.summarize(ol.match_entries(replay,project.items))
        check(replay_summary['found']==count,f'Eksporta roundtrip: {replay_summary}')
        check(hashes_of(paths)==before,'Avotu hash mainījās sagatavošanas laikā')

    if not word_available:
        suite.note('Word COM nav pieejams — Word merge cikli tiks izlaisti (SKIP).')
        for label in ('cikls 1: natural secība','cikls 2: saraksta secība (apgriezta)',
                      'cikls 3: atkārtojums (idempotence)','cikls 4: STRICT bloķē'):
            suite.record(f'soak {label}',status='SKIP',detail='Word COM nav pieejams',section=SOAK_CYCLES)
        return suite

    baseline=winword_pids(); temp_before=temp_leftover_baseline()
    memory_start=process_memory_mb(); objects_start=object_count()
    suite.meta['memory_mb_start']=memory_start; suite.meta['objects_start']=objects_start
    add_soak_cycles(suite,work,paths,markers,before,baseline,temp_before)
    return suite


def add_soak_cycles(suite,work,paths,markers,before,baseline,temp_before=None):
    count=len(paths)
    reversed_names=[p.name for p in reversed(paths)]

    with suite.test(f'soak cikls 1: {count} dokumenti natural secībā',section=SOAK_CYCLES):
        out=work/'SOAK_ALL.docx'
        result=service().run(project_of(paths,out))
        check(result['merged']==count,f"Apvienoti {result['merged']}/{count}: {result['errors'][:2]}")
        check(out.is_file(),'Izvade neeksistē')
        assert_ordered(suite,out,markers,'Soak 1:')
        check(page_break_count(out)==count-1,f"Lappuses pārtraukumi: {page_break_count(out)}")
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        assert_no_new_temp_files(temp_before,'Soak 1:')
        check(can_open_for_write(out),'Izvades fails ir aizņemts')
        no_leftover_word(baseline,'Soak 1:')
        suite.meta['cycle1_duration_s']=round(suite.results[-1].duration_s,1)

    with suite.test(f'soak cikls 2: saraksta secība ({count} apgriezti)',section=SOAK_CYCLES):
        out=work/'SOAK_REVERSED.docx'
        project=project_of(paths,out,order_entries=reversed_names,strict=False,order_path=work/'soak_order.json')
        check(project.items[0].filename==paths[-1].name,f'Saraksta secība nav piemērota: {project.items[0].filename}')
        result=service().run(project)
        check(result['merged']==count,f"Apvienoti {result['merged']}/{count}: {result['errors'][:2]}")
        check(result['order_list']['summary']['found']==count,f"Order saraksts: {result['order_list']['summary']}")
        assert_ordered(suite,out,markers[::-1],'Soak 2:')
        check(hashes_of(paths)==before,'Avotu hash mainījās')
        assert_no_new_temp_files(temp_before,'Soak 2:')
        no_leftover_word(baseline,'Soak 2:')
        suite.meta['cycle2_duration_s']=round(suite.results[-1].duration_s,1)

    with suite.test('soak cikls 3: atkārtojums (idempotence, backup, manifests)',section=SOAK_CYCLES):
        out=work/'SOAK_REVERSED.docx'
        result=service().run(project_of(paths,out,order_entries=reversed_names,strict=True,
                                       order_path=work/'soak_order.json'))
        check(result['merged']==count,f"Apvienoti {result['merged']}/{count}: {result['errors'][:2]}")
        check(result['manifest_diff']['unchanged']==count,f"UNCHANGED: {result['manifest_diff']['unchanged']}")
        check(Path(str(out)+'.bak').is_file(),'Backup .bak netika izveidots')
        assert_ordered(suite,out,markers[::-1],'Soak 3:')
        check(out.stat().st_size>0,'Izvade ir tukša')
        no_leftover_word(baseline,'Soak 3:')
        suite.meta['cycle3_duration_s']=round(suite.results[-1].duration_s,1)

    with suite.test('soak cikls 4: STRICT bloķē merge ar bojātu sarakstu',section=SOAK_CYCLES):
        out=work/'SOAK_STRICT.docx'
        entries=[p.name for p in paths[:5]]+['neeksiste_soak_999.docx']
        project=project_of(paths,out,order_entries=entries,strict=True,order_path=work/'soak_strict.csv')
        try:
            service().run(project)
        except Exception as exc:  # noqa: BLE001 - STRICT bloķē merge
            check('STRICT MODE' in str(exc),f'Nepareizs kļūdas teksts: {exc}')
            check(not out.exists(),'Izvade tika izveidota, lai gan STRICT bloķēja merge')
        else:
            raise AssertionError('STRICT režīmā ar trūkstošu ierakstu merge bija jābloķē')
        no_leftover_word(baseline,'Soak 4:')

    with suite.test('soak: atmiņas un procesu stāvoklis pēc cikliem',section=SOAK_CYCLES):
        memory_start=suite.meta.get('memory_mb_start'); objects_start=suite.meta.get('objects_start')
        memory_end=process_memory_mb(); objects_end=object_count()
        suite.meta['memory_mb_end']=memory_end; suite.meta['objects_end']=objects_end
        if memory_start and memory_end and memory_end>max(900,memory_start*3):
            suite.note(f'BRĪDINĀJUMS: atmiņas pieaugums liels: {memory_start} MB -> {memory_end} MB')
        if objects_start and objects_end>objects_start*4:
            suite.note(f'BRĪDINĀJUMS: Python objektu skaits pieaudzis: {objects_start} -> {objects_end}')
        check(hashes_of(paths)==before,'Avotu hash mainījās visos ciklos')
        assert_no_new_temp_files(temp_before,'Soak resursi:')
        check(list((RUNTIME/'unlocked').glob('*'))==[],'Palika atbloķētās kopijas')
        no_leftover_word(baseline,'Soak resursi:')
