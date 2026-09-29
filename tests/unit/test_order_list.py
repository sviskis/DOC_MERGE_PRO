"""CSV/JSON secības saraksta testi: parsēšana, matching, STRICT, eksports."""
import json

import pytest

from docmerge.core import order_list as ol
from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.preflight import preflight_all
from docmerge.core.sorter import sort_items
from docmerge.domain.enums import OrderListStatus, SortMode
from docmerge.domain.errors import ErrorCode, OrderListError, StrictOrderError
from docmerge.domain.models import DocumentItem, Project


def make_items(tmp_path,names):
    items=[]
    for name in names:
        path=tmp_path/name; path.parent.mkdir(parents=True,exist_ok=True); write_minimal_docx(path,(name,))
        item=DocumentItem(str(path)); item.hydrate_from_path(); items.append(item)
    return preflight_all(items)


# --------------------------------------------------------------------------- CSV
def test_csv_single_column_with_header():
    entries,problems=ol.parse_csv_entries('filename\n01_ievads.docx\n05_stasts.doc\n')
    assert [e.requested for e in entries]==['01_ievads.docx','05_stasts.doc']
    assert [e.order for e in entries]==[1,2]
    assert problems==[]


def test_csv_without_header():
    entries,_=ol.parse_csv_entries('a.docx\nb.docx\n')
    assert [e.requested for e in entries]==['a.docx','b.docx']


def test_csv_order_filename_header():
    entries,_=ol.parse_csv_entries('order,filename\n3,c.docx\n1,a.docx\n2,b.docx\n')
    assert [(e.order,e.requested) for e in entries]==[(3,'c.docx'),(1,'a.docx'),(2,'b.docx')]


def test_csv_path_column_and_backslashes():
    entries,_=ol.parse_csv_entries('path\napaksmape\\01_ievads.docx\n')
    assert entries[0].requested=='apaksmape\\01_ievads.docx'


def test_csv_semicolon_delimiter():
    entries,_=ol.parse_csv_entries('order;path\n2;b\\x.docx\n1;a\\y.docx\n')
    assert [(e.order,e.requested) for e in entries]==[(2,'b\\x.docx'),(1,'a\\y.docx')]


def test_csv_bom_and_cp1257(tmp_path):
    bom=tmp_path/'bom.csv'; bom.write_text('filename\nāčēģī.docx\n',encoding='utf-8-sig')
    entries,meta=ol.load_order_list(bom)
    assert meta['encoding']=='utf-8-sig' and entries[0].requested=='āčēģī.docx'
    legacy=tmp_path/'cp1257.csv'; legacy.write_bytes('filename\nāčēģī.docx\n'.encode('cp1257'))
    entries,meta=ol.load_order_list(legacy)
    assert meta['encoding']=='cp1257' and entries[0].requested=='āčēģī.docx'


def test_csv_comments_and_problems():
    entries,_=ol.parse_csv_entries('# komentars\n\nfilename\na.docx\n\n')
    assert [e.requested for e in entries]==['a.docx']
    entries,problems=ol.parse_csv_entries('order,filename\n1,a.docx\n1,b.docx\nx,c.docx\n')
    assert [e.order for e in entries]==[1,1,2]
    assert any('Atkārtots order' in p for p in problems)
    assert any('nederīgs order' in p for p in problems)


def test_csv_empty_text_is_reported():
    entries,problems=ol.parse_csv_entries('')
    assert entries==[] and problems and 'tukšs' in problems[0].lower()


def test_load_order_list_missing_file(tmp_path):
    with pytest.raises(OrderListError) as info:
        ol.load_order_list(tmp_path/'nav.csv')
    assert info.value.error_code==ErrorCode.E_ORDER_LIST
    assert info.value.stage.value=='ORDER_LIST'


# -------------------------------------------------------------------------- JSON
def test_json_array_of_strings():
    entries,_=ol.parse_json_entries('["a.docx","b.docx"]')
    assert [e.requested for e in entries]==['a.docx','b.docx']


def test_json_files_object():
    entries,_=ol.parse_json_entries('{"files":["a.docx","b.docx"]}')
    assert [e.requested for e in entries]==['a.docx','b.docx']


def test_json_files_object_with_order():
    raw='{"files":[{"order":2,"filename":"b.docx"},{"order":1,"filename":"a.docx"}]}'
    entries,_=ol.parse_json_entries(raw)
    assert [(e.order,e.requested) for e in entries]==[(2,'b.docx'),(1,'a.docx')]


def test_json_path_preferred_over_filename():
    entries,_=ol.parse_json_entries('[{"filename":"a.docx","path":"sub\\\\a.docx"}]')
    assert entries[0].requested=='sub\\a.docx'


def test_json_broken_raises_with_text():
    with pytest.raises(OrderListError) as info:
        ol.parse_json_entries('{"files":[{"order":1,')
    assert 'Bojāts JSON' in info.value.message


def test_json_invalid_root_and_bad_entries():
    with pytest.raises(OrderListError):
        ol.parse_json_entries('"tikai virkne"')
    entries,problems=ol.parse_json_entries('{"files":[{"filename":"a.docx"},5,{"order":2,"filename":"b.docx"}]}')
    assert [e.requested for e in entries]==['a.docx','b.docx']
    assert any('nav ne virkne' in p for p in problems)


def test_json_empty_files_reported():
    entries,problems=ol.parse_json_entries('{"files":[]}')
    assert entries==[] and any('tukšs' in p.lower() for p in problems)


def test_format_detection_by_extension(tmp_path):
    assert ol.format_for_path(tmp_path/'a.json')=='JSON'
    assert ol.format_for_path(tmp_path/'a.csv')=='CSV'
    assert ol.format_for_path(tmp_path/'a.txt')=='CSV'


# ----------------------------------------------------------------------- matching
def test_matching_priorities(tmp_path):
    items=make_items(tmp_path,['01_a.docx','sub/02_b.docx','sub2/01_a.docx'])
    entries,_=ol.parse_csv_entries('filename\n01_a.docx\n')
    assert ol.match_entries(entries,items)[0].status==OrderListStatus.AMBIGUOUS
    entries,_=ol.parse_csv_entries('path\nsub\\02_b.docx\n')
    match=ol.match_entries(entries,items)[0]
    assert match.status==OrderListStatus.FOUND and match.rule==ol.RULE_RELATIVE_PATH
    entries,_=ol.parse_csv_entries('filename\nsub2\\01_A.DOCX\n')
    match=ol.match_entries(entries,items)[0]
    assert match.status==OrderListStatus.FOUND
    assert match.rule in (ol.RULE_RELATIVE_PATH_CI,ol.RULE_FILENAME_CI)


def test_matching_missing_ignored_duplicate(tmp_path):
    (tmp_path/'notes.txt').write_text('x',encoding='utf-8')
    items=make_items(tmp_path,['a.docx','notes.txt'])
    assert ol.match_entries(ol.parse_csv_entries('filename\nnav.docx\n')[0],items)[0].status==OrderListStatus.MISSING
    assert ol.match_entries(ol.parse_csv_entries('filename\nnotes.txt\n')[0],items)[0].status==OrderListStatus.IGNORED
    matches=ol.match_entries(ol.parse_csv_entries('filename\na.docx\na.docx\n')[0],items)
    assert [m.status for m in matches]==[OrderListStatus.FOUND,OrderListStatus.DUPLICATE]


def test_matching_duplicate_order_numbers(tmp_path):
    items=make_items(tmp_path,['a.docx','b.docx'])
    entries,_=ol.parse_json_entries('[{"order":1,"filename":"a.docx"},{"order":1,"filename":"b.docx"}]')
    matches=ol.match_entries(entries,items)
    assert [m.status for m in matches]==[OrderListStatus.FOUND,OrderListStatus.DUPLICATE]


def test_ambiguous_does_not_pick_a_file(tmp_path):
    items=make_items(tmp_path,['a/dup.docx','b/dup.docx'])
    matches=ol.match_entries(ol.parse_csv_entries('filename\ndup.docx\n')[0],items)
    assert matches[0].status==OrderListStatus.AMBIGUOUS
    assert matches[0].resolved_path==''


def test_summarize_counts_and_blocking(tmp_path):
    items=make_items(tmp_path,['a.docx','b.docx'])
    entries,_=ol.parse_csv_entries('filename\na.docx\nnav.docx\n')
    summary=ol.summarize(ol.match_entries(entries,items))
    assert (summary['found'],summary['missing'],summary['blocking_count'])==(1,1,1)
    assert 'nav.docx' in summary['problems'][0]


# ------------------------------------------------------------------------- strict
def test_enforce_strict_raises_and_is_noop_when_off(tmp_path):
    items=make_items(tmp_path,['a.docx'])
    matches=ol.match_entries(ol.parse_csv_entries('filename\nnav.docx\n')[0],items)
    ol.enforce_strict(matches,False)
    with pytest.raises(StrictOrderError) as info:
        ol.enforce_strict(matches,True)
    assert info.value.error_code==ErrorCode.E_STRICT_ORDER
    assert 'STRICT MODE' in info.value.message


def test_normalize_parts_handles_slashes_and_quotes():
    assert ol.normalize_parts('a\\b\\c.docx')==('a','b','c.docx')
    assert ol.normalize_parts('a/b/c.docx')==('a','b','c.docx')
    assert ol.normalize_parts('"./a.docx"')==('a.docx',)
    assert ol.normalize_parts('')==()


# --------------------------------------------------------------- apply / evaluate
def test_apply_order_sets_absolute_order_and_disables_others(tmp_path):
    items=make_items(tmp_path,['a.docx','b.docx','c.docx'])
    project=Project(name='t',items=list(items))
    entries,_=ol.parse_csv_entries('filename\nc.docx\na.docx\n')
    matches=ol.match_entries(entries,project.items)
    ordered,others=ol.apply_order(project,entries,matches,strict=False,order_path='order.csv')
    assert [x.filename for x in ordered]==['c.docx','a.docx']
    assert [x.filename for x in project.items[:2]]==['c.docx','a.docx']
    assert all(not x.enabled for x in others)
    assert project.options.sort_mode==SortMode.ORDER_LIST
    assert project.options.order_list_entries==['c.docx','a.docx']
    assert [x.filename for x in sort_items(project.items,SortMode.ORDER_LIST)][:2]==['c.docx','a.docx']
    matches,summary=ol.evaluate_order_list(project)
    assert len(matches)==2 and summary['found']==2


def test_clear_order_list_reenables_items(tmp_path):
    items=make_items(tmp_path,['a.docx','b.docx'])
    project=Project(name='t',items=list(items))
    entries,_=ol.parse_csv_entries('filename\na.docx\n')
    ol.apply_order(project,entries,ol.match_entries(entries,project.items))
    assert any(not x.enabled for x in project.items)
    ol.clear_order_list(project)
    assert all(x.enabled for x in project.items)
    assert project.options.order_list_enabled is False


# ------------------------------------------------------------------------- export
def test_export_roundtrip_csv_and_json(tmp_path):
    items=make_items(tmp_path,['a.docx','sub/b.docx'])
    project=Project(name='t',items=list(reversed(items)))
    for suffix in ('csv','json'):
        target,fmt=ol.export_order(tmp_path/f'order.{suffix}',project.items)
        assert fmt==('JSON' if suffix=='json' else 'CSV')
        entries,meta=ol.load_order_list(target)
        assert len(entries)==2
        matches=ol.match_entries(entries,project.items)
        assert [m.status for m in matches]==[OrderListStatus.FOUND]*2
        assert [m.resolved_name for m in matches]==[x.filename for x in project.items]


def test_export_unsupported_format(tmp_path):
    with pytest.raises(OrderListError):
        ol.export_order(tmp_path/'order.docx',[],fmt='XML')


def test_entries_from_project_and_report(tmp_path):
    project=Project(name='t',items=[])
    project.options.order_list_entries=['a.docx','b.docx']
    entries=ol.entries_from_project(project)
    assert [e.requested for e in entries]==['a.docx','b.docx']
    assert [e.order for e in entries]==[1,2]
    items=make_items(tmp_path,['a.docx'])
    matches=ol.match_entries(ol.parse_csv_entries('filename\nnav.docx\n')[0],items)
    report=ol.order_list_report(matches,{'path':'order.csv'},strict=True)
    assert report['strict'] is True
    assert report['summary']['missing']==1
    assert report['entries'][0]['status']=='MISSING'
    json.dumps(report)  # reportam jābūt JSON serializējamam
