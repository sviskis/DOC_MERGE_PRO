"""Merge izpilde background thread — bez jebkādas Tk saskares.

GUI tikai padod callbackus (on_log/on_progress/on_success/on_error), kas paši
maršrutē izmaiņas uz Tk galveno thread caur UiBridge.

`on_error` vienmēr saņem `ErrorReport` ar pilnu tekstu (repr, type, traceback),
nekad None. Kļūda tiek arī ierakstīta persistent kļūdu logā.
"""
from docmerge.domain.errors import MergeStage
from docmerge.reporting.error_report import capture_exception,log_error_report,persist_error_log_safely


def run_merge_job(service,project,on_log=None,on_progress=None,on_success=None,on_error=None,
                  on_refresh=None,on_started=None,logger=None,error_log_path=None):
    """Izpilda preflight + merge. Atgriež rezultātu vai None (ja kļūda)."""

    def emit(callback,*args):
        if callback is None: return
        try: callback(*args)
        except Exception: pass  # callback kļūda nedrīkst nogāzt merge

    try:
        emit(on_started)
        emit(on_log,'Merge sākts...')
        service.preflight(project)
        emit(on_refresh)
        result=service.run(project,progress=lambda done,total,item: emit(on_progress,done,total,item))
        emit(on_log,f"Merge pabeigts: {result.get('merged')}/{result.get('total')} apvienoti, kļūdas: {len(result.get('errors',[]))}")
        emit(on_success,result)
        return result
    except BaseException as exc:  # noqa: BLE001 - kļūda tiek pilnībā materializēta šeit
        report=capture_exception(exc,stage=MergeStage.WORKER)
        try:
            log_error_report(logger,report,prefix='MERGE WORKER ERROR')
            persist_error_log_safely(report,log_path=error_log_path,logger=logger)
        except Exception: pass
        emit(on_log,'KĻŪDA: '+report.summary)
        emit(on_error,report)
        return None
