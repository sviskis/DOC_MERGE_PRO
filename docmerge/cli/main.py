import argparse,json,os
from pathlib import Path
from docmerge.domain.enums import SortMode
from docmerge.domain.errors import MergeStage
from docmerge.domain.models import Project
from docmerge.core.scanner import scan_folder,filter_word_documents
from docmerge.core.preflight import preflight_all
from docmerge.core.order_list import (apply_order, enforce_strict, evaluate_order_list, export_order,
                                      load_order_list, match_entries, summarize)
from docmerge.persistence.project_store import load_project,save_project
from docmerge.engines.word_com import system_check
from docmerge.reporting.error_report import capture_exception,error_log_path,log_error_report,persist_error_log_safely
from docmerge.reporting.logger import create_logger
from docmerge.core.merge_service import MergeService


def _print(data):
    print(json.dumps(data,ensure_ascii=False,indent=2))


def _order_list_paths(project):
    """No projekta saglabātais order saraksts (path, ja tāds ir)."""
    return project.options.order_list_path, project.options.order_list_entries


def _run_offline(args):
    """Offline testi: --offline-test / --offline-test-full / --soak-test."""
    from docmerge.testing import run_offline_test
    full=bool(args.offline_test_full); soak=bool(args.soak_test)
    result=run_offline_test(full=full,soak=soak)
    _print({'suite':result.get('suite'),'status':result.get('status'),'passed':result.get('passed'),
            'failed':result.get('failed'),'warnings':result.get('warnings'),'skipped':result.get('skipped'),
            'duration_s':result.get('duration_s'),'report_md':result.get('report_md'),
            'result_json':result.get('result_json'),'offline_guard':result.get('offline_guard')})
    return 0 if result.get('status')=='PASS' else 1


def build_parser():
    p=argparse.ArgumentParser(prog='DOC_MERGE_PRO',description='DOC_MERGE_PRO — DOC/DOCX apvienošana ar CSV/JSON secību')
    p.add_argument('--system-check',action='store_true'); p.add_argument('--scan'); p.add_argument('--merge-all',dest='merge_all',help='Visi .doc/.docx no mapes -> viens DOCX (prasa --output)'); p.add_argument('--project'); p.add_argument('--preflight',action='store_true'); p.add_argument('--run',action='store_true'); p.add_argument('--save-project'); p.add_argument('--output'); p.add_argument('--no-titles',dest='titles',action='store_false',help='Neievietot faila nosaukumu kā virsrakstu'); p.add_argument('--no-unlock',dest='unlock',action='store_false',help='Nenoņemt Protected View (Zone.Identifier) marķējumu')
    p.add_argument('--order-list',dest='order_list',help='CSV/JSON secības saraksts (absolūtā merge secība)')
    p.add_argument('--strict-order-list',dest='strict_order_list',action='store_true',help='Saraksta MISSING/AMBIGUOUS/DUPLICATE bloķē merge')
    p.add_argument('--show-order-list',dest='show_order_list',action='store_true',help='Izdrukā saraksta matching tabulu (JSON)')
    p.add_argument('--export-order',dest='export_order',help='Saglabā pašreizējo secību CSV/JSON failā')
    p.add_argument('--open-output',dest='open_output',action='store_true',help='Atver izvades DOCX pēc veiksmīga merge (Windows)')
    p.add_argument('--offline-test',dest='offline_test',action='store_true',help='OFFLINE QUICK testi (bez tīkla)')
    p.add_argument('--offline-test-full',dest='offline_test_full',action='store_true',help='OFFLINE pilnie testi (Word E2E, fault injection)')
    p.add_argument('--soak-test',dest='soak_test',action='store_true',help='OFFLINE soak tests (100–300 DOCX, vairāki cikli)')
    return p


def main():
    p=build_parser(); a=p.parse_args()
    if a.offline_test or a.offline_test_full or a.soak_test: return _run_offline(a)
    if a.system_check: print(json.dumps(system_check(),ensure_ascii=False,indent=2)); return 0
    if a.merge_all and a.project: p.error('Norādi vai nu --merge-all, vai --project')
    # --output ir obligāts tikai tad, ja tiešām notiek merge (eksporta/priekšskatījuma
    # režīmos tas nav vajadzīgs).
    export_only=bool(a.export_order) and not a.run and not a.preflight and not a.order_list
    if a.merge_all and not a.output and not export_only: p.error('--merge-all prasa --output <fails.docx>')
    proj=load_project(a.project) if a.project else (Project(name=Path(a.merge_all or a.scan).name) if (a.merge_all or a.scan) else None)
    if proj is None:p.error('Norādi --project, --scan vai --merge-all')
    if a.merge_all: proj.items=filter_word_documents(scan_folder(a.merge_all,recursive=True))
    if a.scan:proj.items=scan_folder(a.scan)
    proj.options.insert_titles=bool(a.titles); proj.options.unlock_protected_view=bool(a.unlock)
    proj.options.open_output_when_finished=bool(a.open_output)
    if a.output:proj.output_path=a.output
    logger,_=create_logger(); elog=error_log_path()
    if a.order_list:
        # CSV/JSON secība ir absolūta: natural sort pēc tam netiek pielietots.
        try:
            entries,meta=load_order_list(a.order_list)
        except Exception as exc:  # noqa: BLE001
            report=capture_exception(exc,stage=MergeStage.ORDER_LIST,context={'path':a.order_list})
            log_error_report(logger,report,prefix='CLI ORDER LIST ERROR'); persist_error_log_safely(report,log_path=elog,logger=logger)
            _print({'ok':False,'error':report.to_dict()}); return 3
        matches=match_entries(entries,proj.items); summary=summarize(matches)
        if a.show_order_list:
            _print({'order_list':{'meta':meta,'summary':summary,'entries':[m.to_dict() for m in matches]}})
        apply_order(proj,entries,matches,strict=bool(a.strict_order_list),order_path=a.order_list)
        try:
            enforce_strict(matches,proj.options.strict_order_mode)
        except Exception as exc:  # noqa: BLE001 - STRICT MODE bloķē merge, kļūda netiek pazaudēta
            report=capture_exception(exc,stage=MergeStage.ORDER_LIST,context={'path':a.order_list,'blocking':summary['blocking_count']})
            log_error_report(logger,report,prefix='CLI STRICT ORDER ERROR'); persist_error_log_safely(report,log_path=elog,logger=logger)
            _print({'ok':False,'strict':True,'summary':summary,'error':report.to_dict()}); return 3
    if a.export_order:
        target,fmt=export_order(a.export_order,proj.items)
        _print({'ok':True,'exported':str(target),'format':fmt,'count':len(proj.items)})
    if a.preflight:
        preflight_all(proj.items)
        for x in proj.items:print(f'{x.health_status.value:8} {x.filename:35} {x.detected_format}')
    if a.save_project:save_project(proj,a.save_project)
    if (a.run or a.merge_all) and not export_only:
        if not proj.output_path:p.error('Nav output_path')
        s=MergeService(logger,error_log_path=elog); s.preflight(proj)
        try:
            _print(s.run(proj))
        except Exception as exc:  # noqa: BLE001 - pilna kļūda tiek saglabāta un izdrukāta
            report=capture_exception(exc,stage=MergeStage.SAVE_AS,context={'output_path':proj.output_path})
            log_error_report(logger,report,prefix='CLI ERROR'); persist_error_log_safely(report,log_path=elog,logger=logger)
            _print({'ok':False,'error':report.to_dict()})
            return 2
        if a.open_output:
            path=str(proj.output_path)
            if os.name=='nt' and Path(path).is_file():
                try: os.startfile(path)  # noqa: S606 - lietotāja norādīts izvades fails
                except OSError as exc:
                    report=capture_exception(exc,stage=MergeStage.INTERNAL,context={'output_path':path})
                    log_error_report(logger,report,prefix='CLI OUTPUT OPEN ERROR'); persist_error_log_safely(report,log_path=elog,logger=logger)
            else:
                print(json.dumps({'ok':True,'opened':False,'reason':'Tikai Windows/fails neeksistē'},ensure_ascii=False))
    return 0
if __name__=='__main__': raise SystemExit(main())
