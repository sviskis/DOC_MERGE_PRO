"""MergeService + order saraksta integrācija: STRICT bloķē PIRMS Word."""
import json
from pathlib import Path
from unittest.mock import patch

import pytest

from docmerge.core import order_list as ol
from docmerge.core.merge_service import MergeService
from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.preflight import preflight_all
from docmerge.domain.errors import StrictOrderError
from docmerge.domain.models import DocumentItem, Project


def make_project(tmp_path,names,order):
    items=[]
    for name in names:
        path=tmp_path/name; write_minimal_docx(path,(name,))
        item=DocumentItem(str(path)); item.hydrate_from_path(); items.append(item)
    items=preflight_all(items)
    project=Project(name='svc',items=items,output_path=str(tmp_path/'out.docx'))
    entries=[ol.OrderEntry(order=i,requested=r) for i,r in enumerate(order,1)]
    matches=ol.match_entries(entries,project.items)
    ol.apply_order(project,entries,matches,strict=True,order_path=str(tmp_path/'order.csv'))
    return project


def test_strict_order_list_blocks_merge_before_word(tmp_path):
    project=make_project(tmp_path,['a.docx'],['a.docx','neeksiste.docx'])
    calls=[]
    with patch('docmerge.core.merge_service.WordComEngine') as engine:
        engine.side_effect=lambda *a,**k: calls.append(a)
        with pytest.raises(StrictOrderError) as info:
            MergeService().run(project)
    assert 'STRICT MODE' in info.value.message
    assert calls==[],'Word dzinējs tika izsaukts, lai gan STRICT bloķēja'
    assert not Path(project.output_path).exists()


def test_non_strict_order_list_merges_matched_items_only(tmp_path,monkeypatch):
    monkeypatch.chdir(tmp_path)  # runtime/ un reports/ tiek izolēti šim testam
    project=make_project(tmp_path,['a.docx','b.docx'],['b.docx','neeksiste.docx'])
    project.options.strict_order_mode=False
    captured={}

    class FakeEngine:
        def __init__(self,logger=None): pass
        def merge(self,items,output_path,*args,**kwargs):
            # Produkcijā WordComEngine pats izfiltrē tikai eligible failus.
            captured['all']=[Path(x.source_path).name for x in items]
            captured['eligible']=[Path(x.source_path).name for x in items if x.eligible_for_merge]
            Path(output_path).write_text('PK\x03\x04fake',encoding='utf-8')
            return {'ok':True,'output_path':str(output_path),'merged':len(captured['eligible']),'errors':[],
                    'normalized':[],'unlocked':[],'documents_add_attempts':1,'warning':None,
                    'total':len(items),'word_owned':True,'word_dispatch_method':None}

    with patch('docmerge.core.merge_service.WordComEngine',FakeEngine):
        result=MergeService().run(project)
    assert captured['eligible']==['b.docx'],f"Eligible: {captured['eligible']}"
    assert captured['all']==['b.docx','a.docx'],f"Redzamie faili: {captured['all']}"
    assert project.items[1].enabled is False,'Fails ārpus saraksta palika ieslēgts'
    assert any('Nav order sarakstā' in warning for warning in project.items[1].warnings)
    assert result['order_list']['summary']['found']==1
    assert result['order_list']['summary']['missing']==1
    assert result['manifest_diff']['status'] in ('FIRST_RUN','OK','WARN')
    assert Path('runtime/merge_manifest.json').is_file()
    assert Path('reports/result.json').is_file()
    data=json.loads(Path('reports/result.json').read_text(encoding='utf-8'))
    assert data['order_list']['summary']['missing']==1
    diff=data['manifest_diff']
    assert diff['status']=='FIRST_RUN',f'Manifesta diff: {diff}'
    # Manifestā ir visi projekta faili (arī ārpus saraksta esošie, kas tiek izlaisti).
    assert diff['added']==2 and diff['missing']==0,f'Manifesta diff: {diff}'


def test_order_list_state_helper(tmp_path):
    project=make_project(tmp_path,['a.docx'],['a.docx'])
    matches,summary=MergeService().order_list_state(project)
    assert len(matches)==1 and summary['found']==1
