"""Protected View / Zone.Identifier atbloķēšanas testi.

Mark-of-the-Web plūsmas ir Windows NTFS funkcija, tāpēc ADS testi tiek
izlaisti citās OS; pārējā loģika (Read-only, kopija) ir pārnesama.
"""
import os
import stat

import pytest

from docmerge.core.unlock import (copy_unlocked, has_zone_identifier, make_writable,
                                  remove_zone_identifier, unlock_file, zone_stream_path)

ZONE_STREAM='[ZoneTransfer]\r\nZoneId=3\r\n'
windows_only=pytest.mark.skipif(os.name!='nt',reason='Zone.Identifier ir Windows NTFS funkcija')


def write_zone_stream(path):
    with open(zone_stream_path(path),'w',encoding='utf-8') as fh: fh.write(ZONE_STREAM)


def make_read_only(path): os.chmod(path,stat.S_IREAD)


def make_writable_always(path):
    """Atgriež rakstīšanas tiesības, lai pytest varētu izdzēst temp failus."""
    mode=os.stat(path).st_mode
    os.chmod(path,mode|stat.S_IWRITE)


def test_zone_stream_path_points_to_ads(tmp_path):
    assert zone_stream_path(tmp_path/'a.docx').endswith('a.docx:Zone.Identifier')


@windows_only
def test_missing_zone_identifier_is_reported_as_false(tmp_path):
    p=tmp_path/'clean.docx'; p.write_bytes(b'PK\x03\x04')
    assert has_zone_identifier(p) is False
    assert remove_zone_identifier(p) is False
    assert has_zone_identifier(tmp_path/'neeksiste.docx') is False


@windows_only
def test_zone_identifier_is_detected_and_removed(tmp_path):
    p=tmp_path/'from_internet.docx'; p.write_bytes(b'PK\x03\x04')
    write_zone_stream(p)
    assert has_zone_identifier(p) is True
    assert remove_zone_identifier(p) is True
    assert has_zone_identifier(p) is False
    assert p.read_bytes()==b'PK\x03\x04'  # saturs nav mainīts


@windows_only
def test_unlock_file_removes_zone_and_read_only(tmp_path):
    p=tmp_path/'readonly.docx'; p.write_bytes(b'PK\x03\x04')
    write_zone_stream(p); make_read_only(p)
    zone_removed,read_only_removed=unlock_file(p)
    assert zone_removed is True
    assert read_only_removed is True
    assert has_zone_identifier(p) is False
    assert bool(os.stat(p).st_mode & stat.S_IWRITE)
    assert unlock_file(p)==(False,False)  # otrais izsaukums neko nemaina


def test_make_writable_reports_change(tmp_path):
    p=tmp_path/'file.docx'; p.write_bytes(b'x')
    assert make_writable(p) is False
    make_read_only(p)
    assert make_writable(p) is True
    make_writable_always(p)


@windows_only
def test_copy_unlocked_keeps_source_untouched(tmp_path):
    source=tmp_path/'source.docx'; source.write_bytes(b'PK\x03\x04')
    write_zone_stream(source); make_read_only(source)
    target=tmp_path/'unlocked'/'source.docx'
    copy_unlocked(source,target)
    assert target.exists()
    assert has_zone_identifier(target) is False
    assert bool(os.stat(target).st_mode & stat.S_IWRITE)
    # Avots paliek nemainīgs (immutable): gan marķējums, gan Read-only
    assert has_zone_identifier(source) is True
    assert not bool(os.stat(source).st_mode & stat.S_IWRITE)
    make_writable_always(source)
