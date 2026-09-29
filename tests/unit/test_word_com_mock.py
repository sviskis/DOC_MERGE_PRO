"""Mock COM testi WordComEngine — bez reāla Word.

Pārbauda: Documents.Add() retry, InsertFile kļūdas saglabāšanu, normalizācijas
fallback un to, ka avota faili netiek modificēti.
"""
import hashlib
import os
import sys
import types
from pathlib import Path

import pytest

from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.paths import absolute_path
from docmerge.core.preflight import preflight_all
from docmerge.domain.enums import ErrorPolicy
from docmerge.domain.errors import ErrorCode, MergeError, MergeStage
from docmerge.domain.models import DocumentItem
from docmerge.engines import word_com
from docmerge.engines.word_com import WordComEngine

COM_ARGS=(-2147352567,'Exception occurred.',
          (0,'Microsoft Word','Word cannot start the converter WRD6ER32.CNV.','wdmain11.chm',24612,-2146823156),None)
INSERT_ARGS=(-2147352567,'Exception occurred.',
             (0,'Microsoft Word','The document name or path is not valid.','wdmain11.chm',24753,-2146823015),None)


class FakeComError(Exception):
    def __init__(self,args=COM_ARGS):
        super().__init__(args); self.args=args


class FakeFont:
    def __init__(self): self.Bold=None; self.Italic=None


class FakeParagraphFormat:
    def __init__(self): self.SpaceBefore=None; self.SpaceAfter=None; self.KeepWithNext=None


class FakeRange:
    def __init__(self,doc):
        self.doc=doc; self._text=''; self.Style=None; self.Font=FakeFont(); self.ParagraphFormat=FakeParagraphFormat()
        self.collapse_calls=0
    @property
    def Text(self): return self._text
    @Text.setter
    def Text(self,value):
        self._text=value; self.doc.word.titles.append(value.strip()); self.doc.end+=len(value)
    def Collapse(self,direction): self.collapse_calls+=1
    def InsertFile(self,FileName=None,ConfirmConversions=None,Link=None,Attachment=None):
        self.doc.word.insert_calls.append(FileName)
        if self.doc.word.should_fail(FileName): raise FakeComError(INSERT_ARGS)
        self.doc.end+=10
    def InsertBreak(self,code): self.doc.word.breaks.append(code); self.doc.end+=1


class FakeContent:
    def __init__(self,doc): self.doc=doc
    @property
    def End(self): return self.doc.end


class FakeDocument:
    def __init__(self,word,name='Doc1'):
        self.word=word; self.name=name; self.end=1; self.closed=False; self.saved=None
    @property
    def Content(self): return FakeContent(self)
    def Range(self,start,end): return FakeRange(self)
    def SaveAs2(self,path,FileFormat=None):
        if self.word.saveas_failures>0:
            self.word.saveas_failures-=1; raise FakeComError()
        self.saved=(path,FileFormat); self.word.saved_paths.append(path)
        with open(path,'wb') as fh: fh.write(b'PK\x03\x04fake-docx')
    def Close(self,save=False):
        if self.word.close_failures>0:
            self.word.close_failures-=1; raise FakeComError()
        self.closed=True


class FakeDocuments:
    def __init__(self,word): self.word=word
    def Add(self):
        self.word.add_calls+=1
        if self.word.add_failures>0:
            self.word.add_failures-=1; raise FakeComError()
        doc=FakeDocument(self.word,f'Document{self.word.add_calls}'); self.word.docs.append(doc); return doc
    def Open(self,path,**kwargs):
        self.word.open_calls.append(path)
        if self.word.open_failures>0:
            self.word.open_failures-=1
            raise FakeComError((0,'Microsoft Word',"Sorry, we couldn't find your file.",'wdmain11.chm',24654,-2146823114))
        doc=FakeDocument(self.word,str(path)); self.word.docs.append(doc); return doc


class FakeWord:
    def __init__(self):
        self.Version='16.0'; self.Visible=True; self.DisplayAlerts=-1
        self.add_calls=0; self.add_failures=0; self.open_failures=0
        self.saveas_failures=0; self.close_failures=0
        self.fail_insert_paths=set(); self.fail_all_inserts=False
        self.insert_calls=[]; self.breaks=[]; self.saved_paths=[]; self.open_calls=[]; self.docs=[]; self.titles=[]
        self.quit_calls=0; self.Documents=FakeDocuments(self)
    def should_fail(self,path):
        if self.fail_all_inserts: return True
        return str(path) in self.fail_insert_paths
    def Quit(self): self.quit_calls+=1


@pytest.fixture
def fake_word(monkeypatch):
    word=FakeWord()
    pythoncom=types.SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None)
    client=types.SimpleNamespace(DispatchEx=lambda name:word)
    win32com=types.SimpleNamespace(client=client)
    monkeypatch.setitem(sys.modules,'pythoncom',pythoncom)
    monkeypatch.setitem(sys.modules,'win32com',win32com)
    monkeypatch.setitem(sys.modules,'win32com.client',client)
    return word


def make_sources(tmp_path,count=2):
    items=[]
    for i in range(count):
        path=tmp_path/f'source_{i}.docx'
        write_minimal_docx(path,(f'Tests {i}',))
        item=DocumentItem(str(path)); item.hydrate_from_path(); items.append(item)
    return preflight_all(items)


def sha(path):
    with open(path,'rb') as fh: return hashlib.sha256(fh.read()).hexdigest()


def test_documents_add_retry_makes_merge_succeed(fake_word,tmp_path):
    fake_word.add_failures=1  # pirmais Add() nokrīt ar WRD6ER32.CNV, otrais izdodas
    items=make_sources(tmp_path,2)
    out=tmp_path/'out.docx'
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(out))
    assert result['ok'] is True
    assert result['merged']==2
    assert result['documents_add_attempts']==2
    assert result['warning'] is None
    assert out.exists()
    assert fake_word.quit_calls==1
    assert fake_word.breaks==[7]  # atdalītājs tikai otrajam dokumentam


def test_documents_add_failure_is_reported_with_com_text(fake_word,tmp_path):
    fake_word.add_failures=99; fake_word.open_failures=99
    items=make_sources(tmp_path,1)
    with pytest.raises(MergeError) as info:
        WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    error=info.value
    assert error.stage==MergeStage.DOCUMENTS_ADD
    assert error.error_code==ErrorCode.E_DOCUMENTS_ADD
    assert 'WRD6ER32.CNV' in error.original_message
    assert 'WRD6ER32.CNV' in error.full_text
    assert error.context.get('fallback_error')


def test_documents_add_fallback_uses_blank_docx(fake_word,tmp_path):
    fake_word.add_failures=99  # fallback: Documents.Open(tukšs DOCX)
    items=make_sources(tmp_path,1)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==1
    assert result['warning'] and 'Documents.Add()' in result['warning']
    assert (tmp_path/'normalized'/'_blank_master.docx').exists()


def test_insert_file_error_keeps_full_com_text(fake_word,tmp_path):
    fake_word.fail_all_inserts=True
    items=make_sources(tmp_path,1)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['ok'] is False
    assert result['merged']==0
    assert len(result['errors'])==1
    entry=result['errors'][0]
    assert entry['error_code']==ErrorCode.E_INSERT_FILE.value
    assert entry['stage']==MergeStage.INSERT_FILE.value
    assert '24753' in entry['repr'] or 'not valid' in entry['repr']
    assert entry['traceback']
    assert entry['error']!='None'
    assert entry['used_normalized'] is True


def test_normalization_fallback_merges_item_and_keeps_source_untouched(fake_word,tmp_path):
    items=make_sources(tmp_path,1)
    source=items[0].source_path
    before=sha(source)
    fake_word.fail_insert_paths={source}  # tiešais InsertFile nokrīt, normalizētais strādā
    # cleanup_normalized=False: normalizētais DOCX paliek diskā (noklusēti tas ir temp un tiek dzēsts).
    result=WordComEngine(normalized_dir=tmp_path/'normalized',cleanup_normalized=False).merge(items,str(tmp_path/'out.docx'))
    assert result['ok'] is True
    assert result['merged']==1
    assert result['normalized']==[source]
    assert sha(source)==before  # avots nav modificēts
    normalized=list((tmp_path/'normalized').glob('*.normalized.docx'))
    assert len(normalized)==1
    assert str(normalized[0]) in fake_word.insert_calls
    assert source in fake_word.insert_calls
    assert fake_word.docs[0].closed is True


def test_normalized_temp_files_are_cleaned_by_default(fake_word,tmp_path):
    """Normalizētie DOCX ir pārģenerējami temp faili — pēc merge tiem jāpazūd."""
    items=make_sources(tmp_path,1)
    source=items[0].source_path
    fake_word.fail_insert_paths={source}
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==1
    assert result['normalized']==[source]
    assert list((tmp_path/'normalized').glob('*.normalized.docx'))==[]
    assert result['word_owned'] is True


def test_save_as_error_is_reported_with_com_text(fake_word,tmp_path):
    fake_word.saveas_failures=1
    items=make_sources(tmp_path,1)
    with pytest.raises(MergeError) as info:
        WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert info.value.stage==MergeStage.SAVE_AS
    assert info.value.error_code==ErrorCode.E_SAVE_AS
    assert 'WRD6ER32.CNV' in info.value.full_text


def test_missing_source_is_skipped_with_report(fake_word,tmp_path):
    items=make_sources(tmp_path,1)
    items[0].source_path=str(tmp_path/'pazudis.docx'); items[0].hydrate_from_path()
    items[0].errors.clear(); items[0].health_status=items[0].health_status.OK
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==0
    entry=result['errors'][0]
    assert entry['error_code']==ErrorCode.E_INVALID_PATH.value
    assert 'neeksistē' in entry['error']


def test_no_eligible_documents_raises_merge_error(tmp_path):
    with pytest.raises(MergeError) as info:
        WordComEngine().merge([],str(tmp_path/'out.docx'))
    assert info.value.error_code==ErrorCode.E_NO_DOCUMENTS
    assert info.value.stage==MergeStage.INIT


# ------------------------------------- virsraksti / atdalītāji / Protected View
def test_titles_are_inserted_for_every_document(fake_word,tmp_path):
    items=make_sources(tmp_path,3)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'),insert_titles=True)
    assert result['merged']==3
    assert fake_word.titles==['source_0','source_1','source_2']
    assert fake_word.breaks==[7,7]  # atdalītājs tikai 2. un 3. dokumentam


def test_titles_are_not_inserted_when_option_is_off(fake_word,tmp_path):
    items=make_sources(tmp_path,2)
    WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert fake_word.titles==[]


def test_separator_is_kept_when_normalization_fallback_is_used(fake_word,tmp_path):
    items=make_sources(tmp_path,2)
    second=items[1].source_path
    fake_word.fail_insert_paths={second}  # tiešais InsertFile nokrīt -> normalizācija
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==2
    assert result['normalized']==[second]
    assert fake_word.breaks==[7]  # atdalītājs netiek pazaudēts


def test_separator_and_title_are_not_duplicated_on_retry(fake_word,tmp_path):
    items=make_sources(tmp_path,3)
    fake_word.fail_insert_paths={items[2].source_path}
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'),insert_titles=True)
    assert result['merged']==3
    assert fake_word.breaks==[7,7]  # nav dubultota lappuses pārtraukuma
    assert fake_word.titles==['source_0','source_1','source_2']  # nav dubultota virsraksta


def test_zone_identifier_source_uses_unlocked_copy_and_cleans_it(fake_word,tmp_path,monkeypatch):
    items=make_sources(tmp_path,2)
    source=items[1].source_path
    before=sha(source)
    monkeypatch.setattr(word_com,'has_zone_identifier',lambda path: str(path)==source)
    unlocked_dir=tmp_path/'unlocked'
    result=WordComEngine(normalized_dir=tmp_path/'normalized',unlocked_dir=unlocked_dir).merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==2
    assert result['unlocked']==[source]
    assert source not in fake_word.insert_calls                 # oriģināls netiek padots Word
    assert any(Path(p).parent==unlocked_dir for p in fake_word.insert_calls)
    assert sha(source)==before                                   # avots nav modificēts
    assert list(unlocked_dir.glob('*'))==[]                      # kopija izdzēsta pēc darba


def test_unlock_copy_is_not_used_when_option_is_off(fake_word,tmp_path,monkeypatch):
    items=make_sources(tmp_path,1)
    source=items[0].source_path
    monkeypatch.setattr(word_com,'has_zone_identifier',lambda path: True)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'),unlock_protected_view=False)
    assert result['merged']==1
    assert result['unlocked']==[]
    assert fake_word.insert_calls==[source]


def test_output_is_unlocked_after_successful_merge(fake_word,tmp_path,monkeypatch):
    calls=[]
    monkeypatch.setattr(word_com,'unlock_file',lambda path: calls.append(Path(path)) or (True,False))
    items=make_sources(tmp_path,1)
    out=tmp_path/'out.docx'
    WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(out))
    assert calls==[out]


def test_output_unlock_is_skipped_when_option_is_off(fake_word,tmp_path,monkeypatch):
    calls=[]
    monkeypatch.setattr(word_com,'unlock_file',lambda path: calls.append(Path(path)) or (False,False))
    items=make_sources(tmp_path,1)
    WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'),unlock_protected_view=False)
    assert calls==[]


# --------------------------------------- STOP_ON_ERROR / normalizācijas kļūdas
def test_stop_on_error_policy_raises_with_original_com_error(fake_word,tmp_path):
    fake_word.fail_all_inserts=True
    items=make_sources(tmp_path,1)
    with pytest.raises(MergeError) as info:
        WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'),
                                                                  error_policy=ErrorPolicy.STOP_ON_ERROR)
    assert info.value.error_code==ErrorCode.E_INSERT_FILE
    assert info.value.stage==MergeStage.INSERT_FILE
    assert '24753' in (info.value.original_repr or '') or 'not valid' in info.value.original_message


def test_normalization_failure_reports_both_errors(fake_word,tmp_path):
    fake_word.fail_all_inserts=True; fake_word.open_failures=99
    items=make_sources(tmp_path,1)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==0
    entry=result['errors'][0]
    assert 'InsertFile un normalizācija neizdevās' in entry['error']
    assert entry['used_normalized'] is True


def test_all_paths_handed_to_word_are_absolute(fake_word,tmp_path,monkeypatch):
    # Relatīvos ceļus Word COM risina pret Word darba mapi (C:\WINDOWS\system32),
    # tāpēc normalizētie/atbloķētie faili Word'am jāpadod ABSOLŪTI.
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(word_com,'has_zone_identifier',lambda path: True)
    unlocked_dir=Path('runtime')/word_com.UNLOCKED_DIR_NAME
    normalized_dir=Path('runtime')/'normalized'
    original_should_fail=fake_word.should_fail
    unlock_parent=os.path.normcase(str(Path(absolute_path(unlocked_dir))))
    fake_word.should_fail=lambda path: original_should_fail(path) or os.path.normcase(str(Path(path).parent))==unlock_parent
    items=make_sources(tmp_path,2)
    result=WordComEngine(normalized_dir=normalized_dir,unlocked_dir=unlocked_dir).merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==2
    assert result['unlocked']==[x.source_path for x in items]
    assert fake_word.insert_calls and fake_word.open_calls
    assert all(os.path.isabs(str(p)) for p in fake_word.insert_calls)
    assert all(os.path.isabs(str(p)) for p in fake_word.open_calls)
    assert any(os.path.normcase(str(Path(p).parent))==unlock_parent for p in fake_word.insert_calls)
    assert any(word_com.UNLOCKED_DIR_NAME in str(p) for p in fake_word.open_calls)  # atbloķētā kopija tika normalizēta


# ------------------------------------------------- Word dispatch (COM aktivācija)
SERVER_ARGS=(-2146959355,'Server execution failed',None,None)


def _install_com(monkeypatch,client):
    pythoncom=types.SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None)
    win32com=types.SimpleNamespace(client=client)
    monkeypatch.setitem(sys.modules,'pythoncom',pythoncom)
    monkeypatch.setitem(sys.modules,'win32com',win32com)
    monkeypatch.setitem(sys.modules,'win32com.client',client)
    monkeypatch.setattr(word_com,'WORD_DISPATCH_RETRY_DELAY',0)


def test_dispatch_falls_back_to_dispatch_when_dispatchex_fails(monkeypatch,tmp_path):
    word=FakeWord(); calls=[]
    def dispatch_ex(name):
        calls.append('DispatchEx'); raise FakeComError(SERVER_ARGS)
    def dispatch(name):
        calls.append('Dispatch'); return word
    _install_com(monkeypatch,types.SimpleNamespace(DispatchEx=dispatch_ex,Dispatch=dispatch))
    items=make_sources(tmp_path,1)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==1
    assert calls==['DispatchEx','Dispatch']


def test_dispatch_retries_when_both_methods_fail_temporarily(monkeypatch,tmp_path):
    word=FakeWord(); calls=[]
    def flaky_dispatchex(application):
        calls.append('DispatchEx')
        if len(calls)==1: raise FakeComError(SERVER_ARGS)  # pirmais DispatchEx nokrīt
        return word
    def always_fail(application):
        calls.append('Dispatch'); raise FakeComError(SERVER_ARGS)
    _install_com(monkeypatch,types.SimpleNamespace(DispatchEx=flaky_dispatchex,Dispatch=always_fail))
    items=make_sources(tmp_path,1)
    result=WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert result['merged']==1
    assert calls==['DispatchEx','Dispatch','DispatchEx']  # 2. kārtā DispatchEx izdevās


def test_dispatch_total_failure_raises_word_unavailable(monkeypatch,tmp_path):
    def boom(name): raise FakeComError(SERVER_ARGS)
    _install_com(monkeypatch,types.SimpleNamespace(DispatchEx=boom,Dispatch=boom))
    items=make_sources(tmp_path,1)
    with pytest.raises(word_com.WordUnavailableError) as info:
        WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,str(tmp_path/'out.docx'))
    assert info.value.error_code==ErrorCode.E_WORD_UNAVAILABLE
    assert 'Server execution failed' in info.value.full_text


# --------------------------------------------- datu drošība: izvade nav avots
def test_output_equal_to_source_is_rejected_before_word(fake_word,tmp_path):
    """Merge nedrīkst pārrakstīt avota failu (source ir immutable)."""
    items=make_sources(tmp_path,1)
    source=items[0].source_path
    before=sha(source)
    with pytest.raises(MergeError) as info:
        WordComEngine(normalized_dir=tmp_path/'normalized').merge(items,source)
    assert info.value.error_code==ErrorCode.E_OUTPUT_IS_SOURCE
    assert 'avota' in info.value.message
    assert fake_word.insert_calls==[],'Word tika izsaukts, lai gan izvade ir avots'
    assert fake_word.quit_calls==0,'Word sesija netika aizvērta drošā ceļā'
    assert sha(source)==before,'AVOTA FAILS TIKA PĀRRAKSTĪTS'


def test_output_equal_to_source_ignores_case(tmp_path):
    items=make_sources(tmp_path,1)
    source=items[0].source_path
    upper=str(source).upper()
    engine=WordComEngine(normalized_dir=tmp_path/'normalized')
    with pytest.raises(MergeError):
        engine._assert_output_is_not_source(upper,items)

