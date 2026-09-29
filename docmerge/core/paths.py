"""Ceļu palīgfunkcijas: absolūtie ceļi, OneDrive/cloud placeholders, garuma kontrole.

Word COM (Range.InsertFile, SaveAs2, Documents.Open) relatīvos ceļus risina pret
Word procesa darba mapi (C:\\WINDOWS\\system32), tāpēc VISI ceļi, kas tiek nodoti
COM, ir jāpadara absolūti.
"""
import os
from pathlib import Path

from docmerge.domain.errors import ErrorCode, MergeStage, ValidationError

WORD_MAX_PATH=255
FILE_ATTRIBUTE_OFFLINE=0x1000
FILE_ATTRIBUTE_RECALL_ON_OPEN=0x40000
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS=0x400000
_CLOUD_ATTRIBUTES=FILE_ATTRIBUTE_OFFLINE|FILE_ATTRIBUTE_RECALL_ON_OPEN|FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS


def absolute_path(path):
    """Atgriež absolūtu (resolve) ceļu kā str, nemainot oriģinālo failu."""
    try:
        return str(Path(path).expanduser().resolve())
    except OSError:
        return os.path.abspath(str(path))


def path_too_long(path,max_length=WORD_MAX_PATH):
    return len(str(path))>max_length


def ensure_word_compatible_path(path,stage=MergeStage.INIT,label='ceļš'):
    """Absolūts ceļš, kas ir derīgs Word COM. Met ValidationError, ja nav derīgs."""
    text=str(path or '').strip()
    if not text:
        raise ValidationError(f'Nav norādīts {label}',stage=stage,error_code=ErrorCode.E_INVALID_PATH)
    resolved=absolute_path(text)
    if path_too_long(resolved):
        raise ValidationError(
            f'{label} ir pārāk garš Word COM ({len(resolved)} > {WORD_MAX_PATH}): {resolved}',
            stage=stage,error_code=ErrorCode.E_PATH_TOO_LONG,context={'path':resolved})
    return resolved


def file_attributes(path):
    try:
        return int(getattr(os.stat(path),'st_file_attributes',0))
    except OSError:
        return 0


def is_cloud_placeholder(path):
    """True, ja fails ir OneDrive/Dropbox 'tikai tiešsaistē' vietturis."""
    if os.name!='nt': return False
    return bool(file_attributes(path)&_CLOUD_ATTRIBUTES)


def describe_path_risks(path):
    """Atgriež brīdinājumu sarakstu par ceļu (relatīvs, garš, OneDrive vietturis)."""
    warnings=[]
    text=str(path or '')
    if not text: return ['Nav norādīts ceļš']
    if not Path(text).is_absolute(): warnings.append('Ceļš nav absolūts — Word COM to var neatrast')
    if path_too_long(absolute_path(text)):
        warnings.append(f'Ceļš garāks par {WORD_MAX_PATH} rakstzīmēm — Word COM to var noraidīt')
    if is_cloud_placeholder(text):
        warnings.append('OneDrive cloud vietturis (tikai tiešsaistē) — lejupielādē failu pirms apvienošanas')
    return warnings
