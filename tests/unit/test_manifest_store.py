"""Manifesta un manifests-diff testi (relative path / size / mtime / SHA-256)."""
from docmerge.core.hashing import sha256_file
from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.preflight import preflight_all
from docmerge.domain.models import DocumentItem, Project
from docmerge.persistence.manifest_store import (DIFF_ADDED, DIFF_CHANGED, DIFF_MISSING, DIFF_UNCHANGED,
                                                build_manifest, compare_manifest, load_manifest, manifest_key,
                                                save_manifest)


def make_project(tmp_path,names,write=True):
    items=[]
    for name in names:
        path=tmp_path/name; path.parent.mkdir(parents=True,exist_ok=True)
        if write or not path.exists(): write_minimal_docx(path,(name,))
        item=DocumentItem(str(path)); item.hydrate_from_path(); items.append(item)
    items=preflight_all(items)
    return Project(name='m',items=items,output_path=str(tmp_path/'out.docx')), items


def test_manifest_contains_relative_path_size_mtime_sha256(tmp_path):
    project,items=make_project(tmp_path,['a.docx','sub/b.docx'])
    manifest=build_manifest(project,items)
    first=manifest['documents'][0]
    assert manifest['document_count']==2
    assert first['relative_path']=='a.docx'
    assert manifest['documents'][1]['relative_path']=='sub\\b.docx'
    assert first['size_bytes']==(tmp_path/'a.docx').stat().st_size
    assert first['modified_time']>0
    assert first['sha256']==sha256_file(str(tmp_path/'a.docx'))
    assert manifest['order_list']['enabled'] is False
    assert first['status']=='READY'


def test_manifest_order_list_section_is_recorded(tmp_path):
    project,items=make_project(tmp_path,['a.docx'])
    project.options.order_list_enabled=True
    project.options.order_list_path='order.json'
    project.options.order_list_entries=['a.docx']
    project.options.strict_order_mode=True
    manifest=build_manifest(project,items)
    assert manifest['order_list']=={'enabled':True,'path':'order.json','strict':True,'entries':['a.docx']}


def test_compare_manifest_detects_all_statuses(tmp_path):
    project,items=make_project(tmp_path,['first.docx','second.docx'])
    previous=build_manifest(project,items)
    write_minimal_docx(tmp_path/'second.docx',('second.docx','mainits saturs'))
    write_minimal_docx(tmp_path/'third.docx',('third.docx',))
    project2,items2=make_project(tmp_path,['first.docx','second.docx','third.docx'],write=False)
    diff=compare_manifest(previous,build_manifest(project2,items2))
    assert (diff['unchanged'],diff['changed'],diff['added'],diff['missing'])==(1,1,1,0)
    assert diff['status']=='WARN'
    assert {entry['status'] for entry in diff['documents']}=={DIFF_UNCHANGED,DIFF_CHANGED,DIFF_ADDED}


def test_compare_manifest_reports_missing_and_first_run(tmp_path):
    project,items=make_project(tmp_path,['only.docx'])
    previous=build_manifest(project,items)
    empty_project=Project(name='m',items=[])
    diff=compare_manifest(previous,build_manifest(empty_project,[]))
    assert diff['missing']==1 and diff['status']=='WARN'
    assert diff['documents'][0]['status']==DIFF_MISSING
    first=compare_manifest(None,previous)
    assert first['status']=='FIRST_RUN' and first['added']==1


def test_load_manifest_handles_missing_and_broken(tmp_path):
    assert load_manifest(tmp_path/'nav.json') is None
    broken=tmp_path/'broken.json'; broken.write_text('{nav',encoding='utf-8')
    assert load_manifest(broken) is None


def test_save_and_load_manifest_roundtrip(tmp_path):
    data={'job_id':'x','document_count':1,'documents':[{'relative_path':'a.docx','size_bytes':1,'sha256':'abc'}]}
    path=save_manifest(tmp_path/'m.json',data)
    assert path.is_file()
    assert load_manifest(path)['document_count']==1


def test_manifest_key_prefers_relative_path():
    assert manifest_key({'relative_path':'Sub\\A.docx'})=='sub/a.docx'
    assert manifest_key({'source':'C:/tmp/B.DOCX'})=='c:/tmp/b.docx'
    assert manifest_key({'relative_path':''})==''
    assert manifest_key(None)==''
