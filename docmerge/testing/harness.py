"""Offline testu ietvars (bez pytest): PASS/FAIL/WARN/SKIP + reporti.

Katrs tests tiek izpildīts atsevišķi; izņēmums NEMET tālāk, bet tiek pilnībā
materializēts (message + repr + traceback) un ierakstīts rezultātos, tāpēc
neviena kļūda nepazūd un viena neizdevusies pārbaude neaptur pārējās.
"""
from __future__ import annotations

import json
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

PASS='PASS'; FAIL='FAIL'; WARN='WARN'; SKIP='SKIP'


class TestSkipped(Exception):
    """Testa izlaišana (piem. Word nav pieejams)."""


class TestWarning(Exception):
    """Brīdinājums: rezultāts WARN, nevis FAIL."""


def check(condition,message='Pārbaude neizdevās'):
    if not condition: raise AssertionError(message)
    return True


def warn(message): raise TestWarning(message)


def skip(message): raise TestSkipped(message)


@dataclass
class TestResult:
    name:str
    section:str=''
    status:str=PASS
    detail:str=''
    duration_s:float=0.0
    error:str|None=None
    error_type:str|None=None
    traceback_text:str|None=None

    def to_dict(self):
        return {'name':self.name,'section':self.section,'status':self.status,'detail':self.detail,
                'duration_s':round(self.duration_s,4),'error':self.error,'error_type':self.error_type,
                'traceback':self.traceback_text}


class Suite:
    """Testu komplekts ar rezultātu uzskaiti, meta datiem un reportiem."""

    def __init__(self,name,report_dir='reports'):
        self.name=name; self.report_dir=Path(report_dir)
        self.results:list[TestResult]=[]; self.meta:dict={}; self.notes:list[str]=[]
        self.started_at=datetime.now(); self._t0=time.time()

    # ------------------------------------------------------------------ recording
    def add(self,result): self.results.append(result); return result

    def record(self,name,status=PASS,detail='',section='',error=None,error_type=None,traceback_text=None,duration_s=0.0):
        return self.add(TestResult(name=name,section=section,status=status,detail=detail,duration_s=duration_s,
                                   error=error,error_type=error_type,traceback_text=traceback_text))

    def note(self,message):
        """Ārējs brīdinājums (nav atsevišķs tests)."""
        self.notes.append(str(message)); return message

    def test(self,name,section='',warn_only=False):
        """Ar `with suite.test('nosaukums'):` — izņēmumi netiek izmesti tālāk."""
        suite=self; started=time.time()
        class _Case:
            def __enter__(self): return self
            def __exit__(self,exc_type,exc,tb):
                duration=time.time()-started
                status=PASS; detail=''; error=None; error_type=None; tb_text=None
                if exc_type is None:
                    status=PASS
                elif issubclass(exc_type,TestSkipped):
                    status=SKIP; detail=str(exc)
                elif issubclass(exc_type,TestWarning):
                    status=WARN; detail=str(exc)
                else:
                    status=WARN if warn_only else FAIL
                    error_type=getattr(exc_type,'__name__',None)
                    try: error=str(exc)
                    except Exception: error=''
                    try: detail=repr(exc)
                    except Exception: detail=''
                    try: tb_text=''.join(traceback.format_exception(exc_type,exc,tb)) or None
                    except Exception: tb_text=None
                suite.record(name,status=status,detail=detail,section=section,error=error,
                             error_type=error_type,traceback_text=tb_text,duration_s=duration)
                return True
        return _Case()

    def run(self,func,name=None,section='',warn_only=False):
        """Izpilda funkciju kā vienu testu (alternatīva `with` blokam)."""
        with self.test(name or getattr(func,'__name__','tests'),section=section,warn_only=warn_only):
            return func()

    # ------------------------------------------------------------------ summary
    @property
    def duration_s(self): return time.time()-self._t0

    def counts(self):
        counts={PASS:0,FAIL:0,WARN:0,SKIP:0}
        for result in self.results: counts[result.status]=counts.get(result.status,0)+1
        return counts

    @property
    def status(self): return FAIL if self.counts()[FAIL] else PASS

    def failures(self): return [r for r in self.results if r.status==FAIL]

    def summary(self):
        counts=self.counts()
        return {'suite':self.name,'status':self.status,'passed':counts[PASS],'failed':counts[FAIL],
                'warnings':counts[WARN],'skipped':counts[SKIP],'tests_run':len(self.results),
                'duration_s':round(self.duration_s,2),'started_at':self.started_at.isoformat(timespec='seconds'),
                'finished_at':datetime.now().isoformat(timespec='seconds'),'meta':dict(self.meta),
                'notes':list(self.notes),'results':[r.to_dict() for r in self.results]}

    # ------------------------------------------------------------------ reports
    def write_reports(self,kind=None):
        """Raksta reports/<KIND>_TEST_REPORT.md un reports/<KIND>_TEST_RESULT.json."""
        kind=(kind or self.name).upper()
        self.report_dir.mkdir(parents=True,exist_ok=True)
        md_path=self.report_dir/f'{kind}_TEST_REPORT.md'
        json_path=self.report_dir/f'{kind}_TEST_RESULT.json'
        data=self.summary()
        md_path.write_text(self.render_markdown(data),encoding='utf-8')
        json_path.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
        return md_path,json_path

    def render_markdown(self,data=None):
        data=data or self.summary()
        lines=[f'# {self.name} test report','','## Kopsavilkums','',
               '| Lauks | Vērtība |','| --- | --- |',
               f'| Statuss | **{data["status"]}** |',
               f'| PASS | {data["passed"]} |',f'| FAIL | {data["failed"]} |',
               f'| WARN | {data["warnings"]} |',f'| SKIP | {data["skipped"]} |',
               f'| Testi kopā | {data["tests_run"]} |',
               f'| Ilgums | {data["duration_s"]} s |',
               f'| Sākts | {data["started_at"]} |',f'| Pabeigts | {data["finished_at"]} |','',
               '## Vide','','| Lauks | Vērtība |','| --- | --- |']
        for key,value in (data.get('meta') or {}).items():
            lines.append(f'| {key} | {self._cell(value)} |')
        lines.append('')
        if data.get('notes'):
            lines+=['## Piezīmes','']+[f'- {self._cell(note)}' for note in data['notes']]+['']
        lines+=['## Rezultāti','','| # | Sadaļa | Tests | Statuss | Ilgums | Detaļas |','| --- | --- | --- | --- | --- | --- |']
        for i,result in enumerate(data['results'],1):
            lines.append(f"| {i} | {self._cell(result['section'])} | {self._cell(result['name'])} | "
                         f"**{result['status']}** | {result['duration_s']} s | {self._cell(result['detail'])} |")
        failures=[r for r in data['results'] if r['status']==FAIL]
        if failures:
            lines+=['','## Kļūdas (pilns teksts)','']
            for result in failures:
                lines+=[f"### {result['section']} / {result['name']}",'',
                        f"- tipa kļūda: `{result.get('error_type')}`",
                        f"- ziņojums: {self._cell(result.get('error'))}",
                        f"- detaļas: {self._cell(result.get('detail'))}"]
                if result.get('traceback'):
                    lines+=['','```',str(result['traceback']).rstrip(),'```']
                lines.append('')
        return '\n'.join(lines)+'\n'

    @staticmethod
    def _cell(value):
        text='' if value is None else str(value)
        return text.replace('|','\\|').replace('\n',' ').strip()


def word_version_short(info):
    """Word versija īsam reportam no `system_check()` rezultāta."""
    if not isinstance(info,dict): return 'nav'
    return str(info.get('word_version') or 'nav')
