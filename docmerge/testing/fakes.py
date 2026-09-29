"""COM fakes fault injection testiem (Word nav nepieciešams).

Atļauj simulēt: Word nav pieejams, Documents.Add kļūda, InsertFile kļūda,
SaveAs2 kļūda, normalizācijas kļūda, aizņemts izvades fails.
"""
from __future__ import annotations

import os
import sys
import types
from contextlib import contextmanager

from docmerge.engines import word_com

COM_ARGS=(-2147352567,'Exception occurred.',
          (0,'Microsoft Word','Word cannot start the converter WRD6ER32.CNV.','wdmain11.chm',24612,-2146823156),None)
INSERT_ARGS=(-2147352567,'Exception occurred.',
             (0,'Microsoft Word','The document name or path is not valid.','wdmain11.chm',24753,-2146823015),None)
SERVER_ARGS=(-2146959355,'Server execution failed',None,None)
SAVEAS_ARGS=(-2147352567,'Exception occurred.',
             (0,'Microsoft Word','This file is locked for editing.','wdmain11.chm',0,-2146823114),None)


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
    @property
    def Text(self): return self._text
    @Text.setter
    def Text(self,value):
        self._text=value; self.doc.word.titles.append(str(value).strip()); self.doc.end+=len(value)
    def Collapse(self,direction): self.doc.word.collapse_calls+=1
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
            self.word.saveas_failures-=1; raise FakeComError(SAVEAS_ARGS)
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
    @property
    def Count(self): return len(self.word.docs)


class FakeWord:
    """Pietiekams Word.Application aizvietotājs engine testiem."""

    def __init__(self):
        self.Version='16.0'; self.Visible=True; self.DisplayAlerts=-1
        self.add_calls=0; self.add_failures=0; self.open_failures=0
        self.saveas_failures=0; self.close_failures=0
        self.fail_insert_paths=set(); self.fail_all_inserts=False
        self.insert_calls=[]; self.breaks=[]; self.saved_paths=[]; self.open_calls=[]; self.docs=[]; self.titles=[]
        self.quit_calls=0; self.collapse_calls=0; self.Documents=FakeDocuments(self)

    def should_fail(self,path):
        if self.fail_all_inserts: return True
        return str(path) in self.fail_insert_paths

    def Quit(self): self.quit_calls+=1


@contextmanager
def fake_com_environment(word=None,client=None,winword_pids=None):
    """Uzstāda fake COM (pythoncom/win32com) un atjauno iepriekšējo stāvokli.

    `winword_pids` nosaka, vai priekšā jau darbojās Word (īpašumtiesību tests).
    """
    word=word or FakeWord()
    if client is None:
        client=types.SimpleNamespace(DispatchEx=lambda name:word,Dispatch=lambda name:word)
    pythoncom=types.SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None,
                                    CoFreeUnusedLibraries=lambda:None)
    win32com=types.SimpleNamespace(client=client)
    saved_modules={key:sys.modules.get(key) for key in ('pythoncom','win32com','win32com.client')}
    saved={}
    for name,value in (('WORD_DISPATCH_RETRY_DELAY',0),('word_process_ids',lambda:winword_pids)):
        saved[name]=getattr(word_com,name)
        setattr(word_com,name,value)
    saved_env=os.environ.pop(word_com.DISPATCH_METHODS_ENV,None)
    sys.modules['pythoncom']=pythoncom; sys.modules['win32com']=win32com; sys.modules['win32com.client']=client
    try:
        yield word
    finally:
        for key,value in saved_modules.items():
            if value is None: sys.modules.pop(key,None)
            else: sys.modules[key]=value
        for name,value in saved.items(): setattr(word_com,name,value)
        if saved_env is not None: os.environ[word_com.DISPATCH_METHODS_ENV]=saved_env


def word_unavailable_client():
    """Klients, kuram abas dispatch metodes nokrīt (Word nav pieejams)."""
    def boom(name): raise FakeComError(SERVER_ARGS)
    return types.SimpleNamespace(DispatchEx=boom,Dispatch=boom)
