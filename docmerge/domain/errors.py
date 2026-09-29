import traceback
from datetime import datetime
from enum import Enum


class MergeStage(str, Enum):
    """Kurā posmā kļūda radās (stage + error code visām kļūdām)."""
    INIT="INIT"
    SYSTEM_CHECK="SYSTEM_CHECK"
    PREFLIGHT="PREFLIGHT"
    WORD_DISPATCH="WORD_DISPATCH"
    DOCUMENTS_ADD="DOCUMENTS_ADD"
    DOCUMENT_OPEN="DOCUMENT_OPEN"
    DOCUMENT_CLOSE="DOCUMENT_CLOSE"
    RANGE="RANGE"
    SEPARATOR="SEPARATOR"
    INSERT_FILE="INSERT_FILE"
    NORMALIZE="NORMALIZE"
    SAVE_AS="SAVE_AS"
    OUTPUT_REPLACE="OUTPUT_REPLACE"
    CLEANUP="CLEANUP"
    ORDER_LIST="ORDER_LIST"
    WORKER="WORKER"
    INTERNAL="INTERNAL"


class ErrorCode(str, Enum):
    E_NO_DOCUMENTS="DM-MERGE-001"
    E_INVALID_PATH="DM-MERGE-002"
    E_PATH_TOO_LONG="DM-MERGE-003"
    E_OUTPUT_LOCKED="DM-MERGE-004"
    E_ORDER_LIST="DM-MERGE-005"
    E_STRICT_ORDER="DM-MERGE-006"
    E_OUTPUT_IS_SOURCE="DM-MERGE-007"
    E_WORD_UNAVAILABLE="DM-WORD-001"
    E_WORD_DISPATCH="DM-WORD-002"
    E_DOCUMENTS_ADD="DM-WORD-003"
    E_INSERT_FILE="DM-WORD-004"
    E_SAVE_AS="DM-WORD-005"
    E_NORMALIZE="DM-WORD-006"
    E_DOCUMENT_OPEN="DM-WORD-007"
    E_VALIDATION="DM-VALID-001"
    E_INTERNAL="DM-INT-999"


def _coerce_stage(value):
    if isinstance(value,MergeStage): return value
    if value is None: return MergeStage.INIT
    try: return MergeStage(value)
    except ValueError: return MergeStage.INIT


def _coerce_code(value):
    if isinstance(value,ErrorCode): return value
    if value is None: return ErrorCode.E_INTERNAL
    try: return ErrorCode(value)
    except ValueError: return ErrorCode.E_INTERNAL


def describe_exception(error):
    """Saglabā precīzu kļūdas tekstu: type name, repr, str un pilnu traceback."""
    if error is None:
        return {'type':None,'message':'','repr':None,'traceback':None}
    try: message=str(error)
    except Exception: message=''
    try: repr_text=repr(error)
    except Exception: repr_text=None
    tb=None
    if getattr(error,'__traceback__',None) is not None:
        try: tb=''.join(traceback.format_exception(type(error),error,error.__traceback__)) or None
        except Exception: tb=None
    return {'type':type(error).__name__,'message':message,'repr':repr_text,'traceback':tb}


class DocMergeError(Exception):
    """Bāzes kļūda ar stage + error code + sākotnējā izņēmuma tekstu."""

    stage=MergeStage.INTERNAL
    error_code=ErrorCode.E_INTERNAL

    def __init__(self,message='',*,stage=None,error_code=None,original=None,context=None):
        self.message=str(message)
        super().__init__(self.message)
        self.stage=_coerce_stage(stage if stage is not None else self.__class__.stage)
        self.error_code=_coerce_code(error_code if error_code is not None else self.__class__.error_code)
        self.context={k:v for k,v in (context or {}).items() if v is not None}
        self.raised_at=datetime.now().isoformat(timespec='seconds')
        detail=describe_exception(original)
        self.original=original
        self.original_type=detail['type']
        self.original_message=detail['message']
        self.original_repr=detail['repr']
        self.traceback_text=detail['traceback']

    def __str__(self): return self.full_text

    @property
    def full_text(self):
        """Pilns teksts: stage, kods, vēstījums, tipi, repr un traceback."""
        lines=[f'[{self.stage.value} | {self.error_code.value}] {self.message}']
        if self.context: lines.append('Konteksts: '+', '.join(f'{k}={v}' for k,v in self.context.items()))
        if self.original_repr is not None:
            lines.append(f'Sākotnējā kļūda ({self.original_type}): {self.original_repr}')
        elif self.original_type:
            lines.append(f'Sākotnējā kļūda: {self.original_type}')
        if self.traceback_text: lines.append('Traceback:\n'+self.traceback_text.rstrip())
        return '\n'.join(lines)

    def to_dict(self):
        return {
            'message':self.message,'stage':self.stage.value,'error_code':self.error_code.value,
            'context':dict(self.context),'raised_at':self.raised_at,
            'original_type':self.original_type,'original_message':self.original_message,
            'original_repr':self.original_repr,'traceback':self.traceback_text,
        }


class WordUnavailableError(DocMergeError):
    stage=MergeStage.WORD_DISPATCH; error_code=ErrorCode.E_WORD_UNAVAILABLE

class ValidationError(DocMergeError):
    stage=MergeStage.PREFLIGHT; error_code=ErrorCode.E_VALIDATION

class MergeError(DocMergeError):
    stage=MergeStage.INSERT_FILE; error_code=ErrorCode.E_INSERT_FILE

class OrderListError(DocMergeError):
    """CSV/JSON secības saraksta lasīšanas, parsēšanas vai matching kļūda."""
    stage=MergeStage.ORDER_LIST; error_code=ErrorCode.E_ORDER_LIST

class StrictOrderError(OrderListError):
    """STRICT MODE bloķēja merge (MISSING / AMBIGUOUS / DUPLICATE saraksta ieraksts)."""
    error_code=ErrorCode.E_STRICT_ORDER

