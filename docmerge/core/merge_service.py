from pathlib import Path

from docmerge.core.order_list import enforce_strict, evaluate_order_list, order_list_report
from docmerge.core.preflight import preflight_all
from docmerge.core.sorter import sort_items
from docmerge.domain.errors import MergeStage
from docmerge.engines.word_com import WordComEngine
from docmerge.persistence.manifest_store import build_manifest, compare_manifest, load_manifest, save_manifest
from docmerge.reporting.error_report import (capture_exception,log_error_report,
                                             persist_error_log_safely,report_from_parts)
from docmerge.reporting.result_report import write_result

MANIFEST_PATH='runtime/merge_manifest.json'
MANIFEST_PREVIOUS_PATH='runtime/merge_manifest_previous.json'
MANIFEST_DIFF_PATH='runtime/manifest_diff.json'
RESULT_PATH='reports/result.json'


def _entry_to_report(entry):
    return report_from_parts(entry.get('error') or '(nav ziņas)',error_type=entry.get('error_type') or 'ComError',
                             stage=entry.get('stage'),error_code=entry.get('error_code'),
                             context={'path':entry.get('path'),'item_id':entry.get('id'),
                                      'used_normalized':entry.get('used_normalized')},
                             repr_text=entry.get('repr'),traceback_text=entry.get('traceback'))


class MergeService:
    def __init__(self,logger=None,error_log_path=None):
        self.logger=logger
        self.error_log_path=Path(error_log_path) if error_log_path else None

    def preflight(self,project):
        project.items=preflight_all(project.items); return project.items

    def unknown_error_log(self,report,prefix):
        """Pilnu kļūdu raksta standard logā; persistent JSONL raksta izsaucējs."""
        return log_error_report(self.logger,report,prefix=prefix)

    def order_list_state(self,project):
        """Saglabātā CSV/JSON saraksta atkārtota novērtēšana -> (matches, summary)."""
        return evaluate_order_list(project)

    def run(self,project,progress=None):
        # 1) Saraksta secība ir absolūta: STRICT režīmā problemātiski ieraksti
        #    (MISSING/AMBIGUOUS/DUPLICATE) bloķē merge PIRMS jebkādas Word darbības.
        order_matches,order_summary=self.order_list_state(project)
        enforce_strict(order_matches,project.options.strict_order_mode)
        items=sort_items(project.items,project.options.sort_mode)
        Path('runtime').mkdir(exist_ok=True); Path('reports').mkdir(exist_ok=True)
        # 2) Manifest (relative path / size / mtime / sha256) + salīdzinājums ar iepriekšējo palaišanu.
        manifest=build_manifest(project,items)
        previous=load_manifest(MANIFEST_PATH)
        manifest_diff=compare_manifest(previous,manifest)
        save_manifest(MANIFEST_PATH,manifest)
        if previous: save_manifest(MANIFEST_PREVIOUS_PATH,previous)
        save_manifest(MANIFEST_DIFF_PATH,manifest_diff)
        order_report=order_list_report(order_matches,{'path':project.options.order_list_path},
                                       project.options.strict_order_mode)
        skipped=sum(not x.eligible_for_merge for x in items)
        warnings=sum(bool(x.warnings) for x in items)
        try:
            result=WordComEngine(self.logger).merge(items,project.output_path,project.options.separator,
                                                    error_policy=project.options.error_policy,progress=progress,
                                                    insert_titles=project.options.insert_titles,
                                                    unlock_protected_view=project.options.unlock_protected_view)
        except Exception as exc:  # noqa: BLE001 - kļūda tiek pilnībā saglabāta un pārmesta tālāk
            report=self.unknown_error_log(capture_exception(exc,stage=MergeStage.SAVE_AS,
                                                            context={'output_path':project.output_path,'total':len(items)}),
                                          'MERGE FAILED')
            write_result(RESULT_PATH,status='ERROR',total=len(items),merged=0,skipped=skipped,
                         errors=1,warnings=warnings,output_path=project.output_path,error=report.to_dict(),
                         order_list=order_report,manifest_diff=manifest_diff)
            raise
        for entry in result.get('errors',[]):
            persist_error_log_safely(_entry_to_report(entry),log_path=self.error_log_path,logger=self.logger)
        errors=len(result['errors'])
        status='DONE' if not warnings and not errors else 'DONE_WITH_WARNINGS'
        result['order_list']=order_report
        result['manifest_diff']=manifest_diff
        write_result(RESULT_PATH,status=status,total=len(items),merged=result['merged'],skipped=skipped,
                     errors=errors,warnings=warnings,output_path=result['output_path'],
                     documents_add_attempts=result.get('documents_add_attempts'),normalized=result.get('normalized'),
                     unlocked=result.get('unlocked'),order_list=order_report,manifest_diff=manifest_diff,
                     error_details=[_entry_to_report(e).to_dict() for e in result.get('errors',[])])
        return result

