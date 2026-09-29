import json
import logging
import pytest

from docmerge.domain.errors import (DocMergeError, ErrorCode, MergeError, MergeStage,
                                   ValidationError, WordUnavailableError)
from docmerge.reporting.error_report import (capture_exception, log_error_report, report_from_parts,
                                             write_error_log)

COM_ARGS=(-2147352567,'Exception occurred.',
          (0,'Microsoft Word','Word cannot start the converter WRD6ER32.CNV.','wdmain11.chm',24612,-2146823156),None)


class FakeComError(Exception):
    """Atdarina pywintypes.com_error str() formātu."""
    def __init__(self,args=COM_ARGS):
        super().__init__(args); self.args=args


def test_capture_exception_keeps_repr_type_and_traceback():
    try:
        raise FakeComError()
    except Exception as exc:  # noqa: BLE001
        report=capture_exception(exc,stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE)
    assert report is not None
    assert report.error_type=='FakeComError'
    assert 'WRD6ER32.CNV' in report.repr
    assert 'WRD6ER32.CNV' in report.message
    assert 'WRD6ER32.CNV' in report.text
    assert report.traceback_text and 'FakeComError' in report.traceback_text
    assert report.stage==MergeStage.INSERT_FILE.value
    assert report.error_code==ErrorCode.E_INSERT_FILE.value
    assert report.text.strip()
    assert report.to_dict()['repr']==report.repr


def test_error_report_text_is_never_none_or_empty():
    report=capture_exception(None)
    assert report.text.strip()
    assert report.error_type=='None'
    empty=capture_exception(ValueError(''))
    assert empty.text.strip()
    assert empty.error_type=='ValueError'


def test_merge_error_preserves_original_com_exception_text():
    original=FakeComError()
    error=MergeError('InsertFile neizdevās',original=original,context={'path':'X:/a.docx'})
    assert error.stage==MergeStage.INSERT_FILE
    assert error.error_code==ErrorCode.E_INSERT_FILE
    assert error.original_type=='FakeComError'
    assert 'WRD6ER32.CNV' in error.original_message
    assert 'WRD6ER32.CNV' in error.original_repr
    assert error.traceback_text is None  # oriģinālam nav traceback (nav raise kontekstā)
    assert 'WRD6ER32.CNV' in error.full_text
    assert 'WRD6ER32.CNV' in str(error)
    assert error.to_dict()['original_repr']==error.original_repr


def test_merge_error_defaults_from_subclasses():
    assert WordUnavailableError('x').stage==MergeStage.WORD_DISPATCH
    assert WordUnavailableError('x').error_code==ErrorCode.E_WORD_UNAVAILABLE
    assert ValidationError('x').stage==MergeStage.PREFLIGHT
    assert DocMergeError('x').error_code==ErrorCode.E_INTERNAL


def test_capture_exception_wraps_docmerge_error_without_losing_text():
    wrapped=MergeError('apvalks',original=FakeComError(),stage=MergeStage.SAVE_AS,error_code=ErrorCode.E_SAVE_AS)
    report=capture_exception(wrapped)
    assert report.message=='apvalks'
    assert report.stage==MergeStage.SAVE_AS.value
    assert report.error_code==ErrorCode.E_SAVE_AS.value
    assert 'WRD6ER32.CNV' in report.text


def test_persistent_error_log_writes_jsonl_with_traceback(tmp_path):
    path=tmp_path/'errors.jsonl'
    try:
        raise FakeComError()
    except Exception as exc:  # noqa: BLE001
        report=capture_exception(exc,stage=MergeStage.INSERT_FILE)
    written=write_error_log(report,log_path=path)
    assert written==path
    lines=path.read_text(encoding='utf-8').strip().splitlines()
    assert len(lines)==1
    data=json.loads(lines[0])
    assert 'WRD6ER32.CNV' in data['repr']
    assert 'FakeComError' in (data['traceback_text'] or '')
    assert data['stage']==MergeStage.INSERT_FILE.value


def test_log_error_report_writes_full_text_to_logger():
    records=[]
    handler=logging.Handler()
    handler.emit=lambda record:records.append(record.getMessage())
    logger=logging.getLogger('docmerge.test.error_report'); logger.handlers=[handler]; logger.setLevel(logging.DEBUG)
    try:
        raise FakeComError()
    except Exception as exc:  # noqa: BLE001
        report=capture_exception(exc,stage=MergeStage.INSERT_FILE)
    log_error_report(logger,report,prefix='TEST')
    text='\n'.join(records)
    assert 'WRD6ER32.CNV' in text
    assert 'TEST' in text
    assert any('tb |' in line for line in records)


def test_report_from_parts_keeps_repr_and_stage():
    report=report_from_parts('kļūda','ComError',stage=MergeStage.SAVE_AS,error_code=ErrorCode.E_SAVE_AS,
                             context={'path':'P'},repr_text=repr(COM_ARGS))
    assert report.stage==MergeStage.SAVE_AS.value
    assert report.error_code==ErrorCode.E_SAVE_AS.value
    assert report.repr==repr(COM_ARGS)
    assert report.context=={'path':'P'}
