import json

from docmerge.domain.errors import ErrorCode, MergeStage
from docmerge.domain.models import DocumentItem
from docmerge.reporting.error_report import capture_exception
from docmerge.workers.merge_worker import run_merge_job

COM_ARGS=(-2147352567,'Exception occurred.',
          (0,'Microsoft Word','Word cannot start the converter WRD6ER32.CNV.','wdmain11.chm',24612,-2146823156),None)


class FakeComError(Exception):
    def __init__(self,args=COM_ARGS):
        super().__init__(args); self.args=args


class OkService:
    def __init__(self): self.preflight_calls=0
    def preflight(self,project): self.preflight_calls+=1; return None
    def run(self,project,progress=None):
        if progress is not None:
            progress(1,2,DocumentItem('a.docx')); progress(2,2,DocumentItem('b.docx'))
        return {'ok':True,'merged':2,'total':2,'errors':[],'output_path':'out.docx','normalized':[]}


class ExplodingService:
    def __init__(self): self.preflight_calls=0
    def preflight(self,project): self.preflight_calls+=1; return None
    def run(self,project,progress=None): raise FakeComError()


class PreflightExplodingService:
    def preflight(self,project): raise ValueError('preflight boom')
    def run(self,project,progress=None): raise AssertionError('run() nedrīkst tikt izsaukts')


def test_on_error_receives_real_error_text_not_none():
    errors=[]; logs=[]
    run_merge_job(ExplodingService(),object(),on_error=errors.append,on_log=logs.append)
    assert len(errors)==1
    report=errors[0]
    assert report is not None
    assert report.text and report.text.strip()
    assert 'WRD6ER32.CNV' in report.text
    assert report.error_type=='FakeComError'
    assert report.traceback_text and 'FakeComError' in report.traceback_text
    assert any('KĻŪDA' in line for line in logs)


def test_on_error_persists_jsonl_with_traceback(tmp_path):
    log_path=tmp_path/'errors.jsonl'; errors=[]
    run_merge_job(ExplodingService(),object(),on_error=errors.append,error_log_path=log_path)
    data=json.loads(log_path.read_text(encoding='utf-8').strip())
    assert 'WRD6ER32.CNV' in data['repr']
    assert 'FakeComError' in (data['traceback_text'] or '')
    assert errors and errors[0].error_type=='FakeComError'


def test_preflight_failure_goes_to_on_error_and_run_is_not_called():
    errors=[]; service=PreflightExplodingService()
    result=run_merge_job(service,object(),on_error=errors.append)
    assert result is None
    assert len(errors)==1 and 'preflight boom' in errors[0].text
    assert errors[0].stage==MergeStage.WORKER.value


def test_success_path_reports_progress_and_result():
    progress=[]; success=[]; logs=[]
    result=run_merge_job(OkService(),object(),on_progress=lambda done,total,item:progress.append((done,total)),
                         on_success=success.append,on_log=logs.append)
    assert result['merged']==2
    assert progress==[(1,2),(2,2)]
    assert success and success[0]['output_path']=='out.docx'
    assert any('Merge pabeigts' in line for line in logs)


def test_callback_exception_does_not_lose_error():
    def boom(report): raise RuntimeError('callback sabojāts')
    result=run_merge_job(ExplodingService(),object(),on_error=boom)
    assert result is None


# ---------------------------------------------------------------- GUI popup regresija
def _legacy_worker_handler():
    """Vecais (bojātais) GUI pattern: lambda izmanto 'e' pēc except bloka."""
    try:
        raise FakeComError()
    except Exception as e:  # noqa: BLE001
        return lambda: str(e)


def _fixed_worker_handler(on_error):
    """Jaunais pattern: teksts tiek saglabāts PIRMS atliktā callback."""
    try:
        raise FakeComError()
    except Exception as exc:  # noqa: BLE001
        text=capture_exception(exc,stage=MergeStage.WORKER).text
        on_error(lambda msg=text:msg)


def test_legacy_deferred_lambda_pattern_lost_the_error():
    callback=_legacy_worker_handler()
    try:
        callback()
    except NameError as exc:
        assert 'e' in str(exc)
    else:
        raise AssertionError('Vecais pattern nedrīkstēja nostrādāt — ja tas mainījās, regresijas tests jāpārskata')


def test_fixed_pattern_delivers_real_error_text_to_popup():
    shown={}
    def popup(text): shown['text']=text
    _fixed_worker_handler(lambda callback:popup(callback()))
    assert shown['text'] and 'WRD6ER32.CNV' in shown['text']
    assert shown['text']!='None'
