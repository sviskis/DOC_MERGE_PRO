"""Vienota kļūdu savākšana: repr + type + message + pilns traceback + persistent logs.

Kļūda tiek pilnībā materializēta šeit (nevis vēlāk lambda), tāpēc GUI nekad
nepazaudē īsto izņēmuma tekstu (piem. Word COM).
"""
import json
import traceback
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path

from docmerge.domain.errors import DocMergeError, ErrorCode, MergeStage, _coerce_code, _coerce_stage

ERROR_LOG_NAME='errors.jsonl'
ERROR_LOG_DIR='logs'


def error_log_path(project_root=None):
    root=Path(project_root) if project_root else Path.cwd()
    return root/ERROR_LOG_DIR/ERROR_LOG_NAME


@dataclass
class ErrorReport:
    message:str=''
    error_type:str='Exception'
    repr:str|None=None
    traceback_text:str|None=None
    stage:str=MergeStage.INTERNAL.value
    error_code:str=ErrorCode.E_INTERNAL.value
    context:dict=field(default_factory=dict)
    original_message:str|None=None
    original_type:str|None=None
    timestamp:str=field(default_factory=lambda:datetime.now().isoformat(timespec='seconds'))

    @property
    def text(self):
        """Cilvēkam lasāms teksts popup logam. Nekad nav tukšs un nekad nav 'None'."""
        if not self.message and not self.repr:
            return f'[{self.stage} | {self.error_code}] Nezināma kļūda bez vēstījuma'
        return self.full_text

    @property
    def full_text(self):
        lines=[f'[{self.stage} | {self.error_code}] {self.error_type}: {self.message}']
        if self.context: lines.append('Konteksts: '+', '.join(f'{k}={v}' for k,v in self.context.items()))
        if self.original_message and self.original_message!=self.message:
            lines.append(f'Sākotnējā kļūda: {self.original_message}')
        if self.repr and self.repr!=self.message: lines.append(f'repr: {self.repr}')
        if self.traceback_text: lines.append('Traceback:\n'+self.traceback_text.rstrip())
        return '\n'.join(lines)

    @property
    def summary(self):
        return f'{self.stage} | {self.error_code} | {self.error_type}: {self.message}'

    def to_dict(self): return asdict(self)


def capture_exception(error,stage=None,error_code=None,context=None):
    """Pārvērš izņēmumu pilnīgā ErrorReport (repr, type, message, traceback)."""
    if isinstance(error,DocMergeError):
        ctx=dict(error.context); ctx.update(context or {})
        return ErrorReport(
            message=error.message,error_type=type(error).__name__,repr=repr(error),
            traceback_text=error.traceback_text,stage=error.stage.value,
            error_code=error.error_code.value,context=ctx,
            original_message=error.original_message,original_type=error.original_type,
        )
    if error is None:
        return ErrorReport(message='(nav ziņas)',error_type='None',stage=_coerce_stage(stage).value,error_code=_coerce_code(error_code).value,context=dict(context or {}))
    try: message=str(error)
    except Exception: message=''
    try: repr_text=repr(error)
    except Exception: repr_text=None
    tb=None
    if getattr(error,'__traceback__',None) is not None:
        try: tb=''.join(traceback.format_exception(type(error),error,error.__traceback__)) or None
        except Exception: tb=None
    return ErrorReport(message=message,error_type=type(error).__name__,repr=repr_text,traceback_text=tb,
                       stage=_coerce_stage(stage).value,error_code=_coerce_code(error_code).value,
                       context=dict(context or {}))


def report_from_parts(message,error_type='Exception',stage=None,error_code=None,context=None,repr_text=None,traceback_text=None):
    return ErrorReport(message=str(message),error_type=error_type,repr=repr_text,traceback_text=traceback_text,
                       stage=_coerce_stage(stage).value,error_code=_coerce_code(error_code).value,context=dict(context or {}))


def log_error_report(logger,report,prefix='KĻŪDA'):
    """Raksta pilnu kļūdu standard loggerī (kuru jau izmanto logs/docmerge_*.log)."""
    if logger is None: return report
    logger.error('%s | %s',prefix,report.summary)
    if report.context: logger.error('%s | Konteksts: %s',prefix,report.context)
    if report.repr and report.repr!=report.message: logger.error('%s | repr: %s',prefix,report.repr)
    if report.traceback_text:
        for line in report.traceback_text.rstrip().splitlines(): logger.error('%s | tb | %s',prefix,line)
    return report


def write_error_log(report,log_path=None,project_root=None):
    """Persistent kļūdu fails (JSONL), saglabā arī pilnu traceback."""
    path=Path(log_path) if log_path else error_log_path(project_root)
    try:
        path.parent.mkdir(parents=True,exist_ok=True)
        with path.open('a',encoding='utf-8') as fh:
            fh.write(json.dumps(report.to_dict(),ensure_ascii=False)+'\n')
    except OSError:
        return None
    return path


def persist_error_log_safely(report,log_path=None,project_root=None,logger=None):
    path=write_error_log(report,log_path=log_path,project_root=project_root)
    if path is None and logger is not None:
        logger.error('Neizdevās ierakstīt %s',ERROR_LOG_NAME)
    return path
