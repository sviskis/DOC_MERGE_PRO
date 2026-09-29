"""Windows Protected View atbloķēšana: Mark-of-the-Web (Zone.Identifier) + Read-only.

Šī loģika ir pārņemta no pārbaudītā skripta (docx_atbloketajs), kas pirms
apvienošanas noņem Zone.Identifier plūsmu. Bez tā Word COM `Range.InsertFile`
var nokrist ar konvertera/kļūdas paziņojumu, un izvades DOCX atveras
Protected View režīmā.

Svarīgi: avota faili paliek nemainīgi (immutable). Ja avotam ir
Zone.Identifier, Word COM saņem atbloķētu KOPIJU (runtime/unlocked), bet vietā
tiek atbloķēta tikai izvades DOCX.
"""
from __future__ import annotations

import os
import shutil
import stat
from pathlib import Path

ZONE_STREAM_SUFFIX = ':Zone.Identifier'
# Windows dažkārt neesošai ADS plūsmai atgriež winerror 2 vai 3.
MISSING_STREAM_WINERRORS = (2, 3)


def zone_stream_path(file_path) -> str:
    """Alternatīvās datu plūsmas ceļš: fail.docx:Zone.Identifier."""
    return f'{Path(file_path)}{ZONE_STREAM_SUFFIX}'


def has_zone_identifier(file_path) -> bool:
    """True, ja failam ir Mark-of-the-Web (Protected View avots).

    Ārpus Windows vienmēr False. Nekāda kļūda netiek pārmesta.
    """
    if os.name != 'nt':
        return False
    try:
        os.stat(zone_stream_path(file_path))
        return True
    except OSError:
        return False


def remove_zone_identifier(file_path) -> bool:
    """Noņem Zone.Identifier ADS. Atgriež True, ja plūsma tika atrasta un noņemta.

    Uzmanību: no Read-only faila Windows atgriež PermissionError (winerror 5) —
    tāpēc pirms šī izsaukuma jānoņem Read-only atribūts (`unlock_file` to dara).
    """
    if os.name != 'nt':
        return False
    try:
        os.remove(zone_stream_path(file_path))
        return True
    except FileNotFoundError:
        return False
    except OSError as exc:
        if getattr(exc, 'winerror', None) in MISSING_STREAM_WINERRORS:
            return False
        raise


def make_writable(file_path) -> bool:
    """Noņem faila Read-only atribūtu. Atgriež True, ja tas bija uzlikts."""
    current_mode = Path(file_path).stat().st_mode
    was_read_only = not bool(current_mode & stat.S_IWRITE)
    if was_read_only:
        os.chmod(file_path, current_mode | stat.S_IWRITE)
    return was_read_only


def unlock_file(file_path) -> tuple[bool, bool]:
    """Atbloķē failu vietā. Atgriež (zone_removed, read_only_removed).

    Secība ir svarīga: VISPIRMS jānoņem Read-only atribūts, jo
    Zone.Identifier plūsmas dzēšana no Read-only faila Windows atgriež
    PermissionError (winerror 5) — tāpēc oriģinālajā skriptā šādi faili varēja
    palikt ar Protected View marķējumu.

    Nekad nemet izņēmumu — atbloķēšana ir papildu drošība, nevis kritiskais ceļš.
    """
    read_only_removed = False
    try:
        read_only_removed = make_writable(file_path)
    except OSError:
        read_only_removed = False
    zone_removed = False
    try:
        zone_removed = remove_zone_identifier(file_path)
    except OSError:
        zone_removed = False
    return zone_removed, read_only_removed


def copy_unlocked(source, target) -> Path:
    """Nokopē failu bez Mark-of-the-Web un Read-only. Avots netiek modificēts.

    Uzmanību: Windows `shutil.copy2` (CopyFile2) pārnes arī alternatīvās datu
    plūsmas, tāpēc Zone.Identifier ir jānoņem arī no KOPIJAS — to dara
    `unlock_file` (vispirms Read-only, tad plūsma).
    """
    source_path = Path(source)
    target_path = Path(target)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, target_path)
    unlock_file(target_path)
    return target_path
