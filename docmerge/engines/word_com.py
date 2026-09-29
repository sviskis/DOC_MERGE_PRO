from pathlib import Path
import gc,os,shutil,subprocess,time,uuid

from docmerge.core.ooxml import write_minimal_docx
from docmerge.core.paths import absolute_path, ensure_word_compatible_path
from docmerge.core.unlock import copy_unlocked, has_zone_identifier, unlock_file
from docmerge.domain.enums import DocumentStatus, ErrorPolicy, SeparatorMode
from docmerge.domain.errors import ErrorCode, MergeError, MergeStage, WordUnavailableError
from docmerge.reporting.error_report import capture_exception, log_error_report

WD_PAGE_BREAK=7; WD_SECTION_BREAK_NEXT_PAGE=2; WD_SECTION_BREAK_CONTINUOUS=3; WD_FORMAT_XML_DOCUMENT=12
WD_STYLE_NORMAL=-1; WD_COLLAPSE_END=0

# Word dažreiz pirmo Documents.Add() noraida ar vienreizēju converter kļūdu
# (piem. 'Word cannot start the converter WRD6ER32.CNV.'), bet atkārtots
# izsaukums izdodas -> tāpēc Add() tiek atkārtots.
DOCUMENTS_ADD_ATTEMPTS=3
DOCUMENTS_ADD_RETRY_DELAY=0.25
BLANK_MASTER_NAME='_blank_master.docx'
UNLOCKED_DIR_NAME='unlocked'
# Pirmā DispatchEx pēc Word aizvēršanas var nokrist ar 'Server execution failed'
# (COM aktivācijas sacensība), tāpēc tiek mēģināts vairākas reizes un ar
# fallback uz Dispatch — tāpat kā pārbaudītais skripts
# (New-Object -ComObject Word.Application).
WORD_DISPATCH_ATTEMPTS=3
WORD_DISPATCH_RETRY_DELAY=0.5
WORD_DISPATCH_METHODS=('DispatchEx','Dispatch')
# Process vārds īpašumtiesību noteikšanai (vai Word jau darbojas lietotājam).
WORD_PROCESS_NAME='WINWORD.EXE'
# Testu/debug hook: ar šo env mainīgo var piespiest tikai Dispatch (fallback ceļu),
# lai pārbaudītu, ka lietotāja esošā Word sesija netiek aizvērta.
DISPATCH_METHODS_ENV='DOC_MERGE_WORD_DISPATCH_METHODS'

_SEPARATOR_BREAK_CODES={SeparatorMode.PAGE_BREAK:WD_PAGE_BREAK,SeparatorMode.SECTION_BREAK_NEXT_PAGE:WD_SECTION_BREAK_NEXT_PAGE,SeparatorMode.SECTION_BREAK_CONTINUOUS:WD_SECTION_BREAK_CONTINUOUS}


def _import_com():
    import pythoncom, win32com.client
    return pythoncom, win32com


def _log_warn(log,message,*args):
    if log is None: return
    try: log.warning(message,*args)
    except Exception: pass


def _client_module(win32com):
    """Atgriež moduli ar Dispatch/DispatchEx (pieņem gan win32com, gan win32com.client).

    Ja `win32com.client` apakšmodulis vēl nav importēts, tas tiek importēts
    (citādi `getattr(win32com,'client')` var atgriezt nepareizo moduli).
    """
    client=getattr(win32com,'client',None)
    if client is not None: return client
    try:
        import win32com.client as client_module
        return client_module
    except Exception:  # noqa: BLE001 - bez pywin32 paliek dotais objekts
        return win32com


class WordSession:
    """Word COM sesija ar īpašumtiesību uzskaiti.

    owned=True  -> instanci izveidojām MĒS (DispatchEx vai Dispatch, ja Word
                   iepriekš nedarbojās) un drīkstam izsaukt `Quit()`.
    owned=False -> pieslēgāmies LIETOTĀJA jau atvērtai Word sesijai; `Quit()`
                   NEDRĪKST izsaukt, citādi tiktu aizvērts lietotāja darbs
                   (DispatchEx -> Dispatch fallback regresija).
    """

    def __init__(self,word,method,owned,process_ids_before=None):
        self.word=word
        self.method=method
        self.owned=bool(owned)
        self.process_ids_before=process_ids_before
        self.display_alerts_before=None
        self.closed=False

    @property
    def safe_to_quit(self): return self.owned and not self.closed

    def configure(self,log=None):
        """Sagatavo sesiju darbam. LIETOTĀJA sesijai Visible NETIEK mainīts."""
        if self.owned:
            try: self.word.Visible=False
            except Exception as exc:  # noqa: BLE001
                _log_warn(log,'Nevar iestatīt Word.Visible=False: %r',exc)
        try: self.display_alerts_before=self.word.DisplayAlerts
        except Exception: self.display_alerts_before=None
        try: self.word.DisplayAlerts=0
        except Exception as exc:  # noqa: BLE001
            _log_warn(log,'Nevar iestatīt Word.DisplayAlerts=0: %r',exc)
        return self

    def close(self,log=None):
        """Aizver sesiju. LIETOTĀJA sesiju NEKAD neaizver (tikai atjauno DisplayAlerts)."""
        if self.closed: return False
        self.closed=True
        if self.owned:
            try:
                self.word.Quit()
                _log_warn(log,'Word sesija aizvērta (owned, %s)',self.method)
            except Exception as exc:  # noqa: BLE001
                _log_warn(log,'Word.Quit kļūda (owned, %s): %r',self.method,exc)
            finally:
                self.word=None
            return True
        try:
            if self.display_alerts_before is not None: self.word.DisplayAlerts=self.display_alerts_before
        except Exception as exc:  # noqa: BLE001
            _log_warn(log,'DisplayAlerts atjaunošana neizdevās: %r',exc)
        _log_warn(log,'Word instanci izveidoja LIETOTĀJS (%s) — Quit() netiek izsaukts',self.method)
        self.word=None
        return False


def dispatch_word(win32com,log=None):
    """Izveido Word COM instanci (atpakaļsaderība) — skat. `open_word_session`.

    BRĪDINĀJUMS: atgriež tikai `.word`, tāpēc īpašumtiesības (drīkst vai nedrīkst
    izsaukt `Quit()`) tiek zaudētas. Jaunam kodam jāizmanto `open_word_session`.
    """
    return open_word_session(win32com,log=log).word


def configured_dispatch_methods():
    """Dispatch metodes; `DOC_MERGE_WORD_DISPATCH_METHODS` ļauj piespiest fallback."""
    raw=os.environ.get(DISPATCH_METHODS_ENV,'')
    methods=tuple(x.strip() for x in raw.split(',') if x.strip())
    return methods or WORD_DISPATCH_METHODS


def word_process_ids():
    """WINWORD.EXE procesu ID (lokāli, bez tīkla). None, ja noteikt nevar."""
    if os.name!='nt': return None
    try:
        cp=subprocess.run(['tasklist','/FI',f'IMAGENAME eq {WORD_PROCESS_NAME}','/NH'],
                          capture_output=True,text=True,timeout=20)
    except (OSError,subprocess.SubprocessError): return None
    if cp.returncode!=0: return None
    pids=set()
    for line in (cp.stdout or '').splitlines():
        parts=line.split()
        if len(parts)>=2 and parts[0].upper()==WORD_PROCESS_NAME:
            try: pids.add(int(parts[1].replace(',','')))
            except ValueError: pass
    return pids


def word_is_running(process_ids=None):
    """True/False, ja zināms; None, ja noteikt nevar (tad drošāk neuzskatīt par savu)."""
    pids=word_process_ids() if process_ids is None else process_ids
    if pids is None: return None
    return bool(pids)


def word_rot_registered(client=None):
    """Vai Word ir reģistrēts Running Object Table (papildu signāls diagnosticēšanai)."""
    try:
        client=client if client is not None else _client_module(__import__('win32com'))
        client.GetActiveObject('Word.Application')
        return True
    except Exception:  # noqa: BLE001 - ROT nav pieejams vai Word nav reģistrēts
        return None


def open_word_session(win32com,log=None,attempts=None):
    """Atver Word COM sesiju ar atkārtojumu, Dispatch fallback un īpašumtiesībām.

    DispatchEx atgriež jaunu instanci (vienmēr mūsu). Dispatch var pieslēgties
    LIETOTĀJA jau atvērtai sesijai — tad `owned=False` un `Quit()` netiek izsaukts.
    """
    client=_client_module(win32com)
    methods=configured_dispatch_methods()
    attempts=WORD_DISPATCH_ATTEMPTS if attempts is None else max(1,int(attempts))
    pids_before=word_process_ids() if 'Dispatch' in methods else None
    last=None
    for attempt in range(1,attempts+1):
        for method in methods:
            try: factory=getattr(client,method,None)
            except Exception: factory=None
            if factory is None: continue
            try:
                word=factory('Word.Application')
            except Exception as exc:  # noqa: BLE001 - COM aktivācija var būt īslaicīgi nepieejama
                last=exc
                _log_warn(log,'Word nav pieejams (mēģinājums %d, %s): %r',attempt,method,exc)
                continue
            if method=='Dispatch':
                running=word_is_running(pids_before)
                # Nezināms -> drošāk uzskatīt par lietotāja sesiju (Quit nenotiks).
                owned=False if running is None else not bool(running)
            else:
                owned=True
            session=WordSession(word,method,owned,process_ids_before=pids_before).configure(log=log)
            if attempt>1 or method!=methods[0] or not session.owned:
                _log_warn(log,'Word instance izveidota ar %s (%d. mēģinājums, owned=%s)',method,attempt,session.owned)
            return session
        time.sleep(WORD_DISPATCH_RETRY_DELAY)
    if last is None:
        raise WordUnavailableError('Word COM nav pieejams (nav pieejama neviena dispatch metode)')
    raise last


def system_check():
    """Pārbauda Word COM pieejamību. Atgriež dict ar pilnu kļūdas informāciju."""
    r={'windows':os.name=='nt','pywin32':False,'word_com':False,'word_version':None,'error':None,
       'error_type':None,'error_repr':None,'error_traceback':None,'stage':None,'error_code':None,
       'word_owned':None,'word_dispatch_method':None,'winword_pids_before':None,
       'winword_pids_after':None,'leftover_winword':None,'cleanup_warnings':[]}
    if os.name!='nt':
        r['error']='Word COM tikai Windows'; r['stage']=MergeStage.SYSTEM_CHECK.value; r['error_code']=ErrorCode.E_WORD_UNAVAILABLE.value
        return r
    session=None; pythoncom=None
    try:
        pythoncom,win32com=_import_com(); r['pywin32']=True
        pids_before=word_process_ids()
        r['winword_pids_before']=sorted(pids_before) if pids_before is not None else None
        pythoncom.CoInitialize()
        session=open_word_session(win32com)
        r['word_com']=True; r['word_version']=str(session.word.Version)
        r['word_owned']=session.owned; r['word_dispatch_method']=session.method
    except Exception as exc:  # noqa: BLE001 - ziņojums tiek pilnībā saglabāts
        report=capture_exception(exc,stage=MergeStage.SYSTEM_CHECK,error_code=ErrorCode.E_WORD_DISPATCH)
        r.update({'error':report.message or '(nav ziņas)','error_type':report.error_type,
                  'error_repr':report.repr,'error_traceback':report.traceback_text,
                  'stage':report.stage,'error_code':report.error_code})
    finally:
        try:
            if session is not None: session.close()
        except Exception as exc:  # noqa: BLE001
            r['cleanup_warnings'].append(repr(exc))
        try:
            if pythoncom is not None: pythoncom.CoUninitialize()
        except Exception as exc:  # noqa: BLE001
            r['cleanup_warnings'].append(repr(exc))
        pids_after=word_process_ids()
        r['winword_pids_after']=sorted(pids_after) if pids_after is not None else None
        if pids_after is not None:
            before=set(r['winword_pids_before'] or [])
            r['leftover_winword']=sorted(set(pids_after)-before)
    return r


class WordComEngine:
    """Production merge engine: Microsoft Word COM + Range.InsertFile.

    Source faili ir immutable. Ja tiešais InsertFile neizdodas, avots tiek
    atvērts ar Word COM (ReadOnly=True), normalizēts uz runtime/normalized/
    DOCX, un tad InsertFile tiek izsaukts ar normalizēto failu.

    Ja avotam ir Mark-of-the-Web (Protected View), Word COM saņem atbloķētu
    kopiju runtime/unlocked (oriģināls netiek mainīts), un izvades DOCX tiek
    atbloķēts (Zone.Identifier + Read-only noņemts).
    """

    def __init__(self,logger=None,normalized_dir=None,add_attempts=DOCUMENTS_ADD_ATTEMPTS,unlocked_dir=None,
                 cleanup_normalized=True):
        self.logger=logger
        self._normalized_dir_arg=Path(normalized_dir) if normalized_dir else None
        self._unlocked_dir_arg=Path(unlocked_dir) if unlocked_dir else None
        self.add_attempts=max(1,int(add_attempts))
        # Normalizētie DOCX ir pārģenerējami temp faili: pēc merge tie tiek izdzēsti.
        self.cleanup_normalized=bool(cleanup_normalized)
        self._created_normalized=[]

    # ------------------------------------------------------------------ logging
    def _log(self,level,message,*args):
        if self.logger is None: return
        try: getattr(self.logger,level)(message,*args)
        except Exception: pass

    def _warn(self,message,*args): self._log('warning',message,*args)
    def _info(self,message,*args): self._log('info',message,*args)
    def _error(self,message,*args): self._log('error',message,*args)

    def _log_report(self,prefix,report):
        if self.logger is not None: log_error_report(self.logger,report,prefix=prefix)
        return report

    # ------------------------------------------------------------------- paths
    def _work_dir(self,base):
        """Absolūta darba mape. Word COM relatīvu ceļu risina pret Word darba mapi."""
        path=Path(base); path.mkdir(parents=True,exist_ok=True)
        return Path(absolute_path(path))

    def _normalized_dir(self):
        base=self._normalized_dir_arg or (Path('runtime')/'normalized')
        return self._work_dir(base)

    def _ensure_blank_docx(self):
        blank=self._normalized_dir()/BLANK_MASTER_NAME
        if not blank.exists(): write_minimal_docx(blank,())
        return blank

    def _unlocked_dir(self):
        """Absolūta darba mape atbloķētām avota kopijām (avots paliek nemainīgs)."""
        base=self._unlocked_dir_arg or (Path('runtime')/UNLOCKED_DIR_NAME)
        return self._work_dir(base)

    def _collect_com_garbage(self):
        """GC + CoFreeUnusedLibraries (skripta [GC]::Collect() ekvivalents)."""
        try: gc.collect()
        except Exception: pass
        try:
            import pythoncom
            pythoncom.CoFreeUnusedLibraries()
        except Exception: pass

    # ------------------------------------------------------------------- merge
    def merge(self,items,output_path,separator=SeparatorMode.PAGE_BREAK,error_policy=None,progress=None,
              insert_titles=False,unlock_protected_view=True):
        eligible=[x for x in items if x.eligible_for_merge]
        if not eligible:
            raise MergeError('Nav apvienojamu failu',stage=MergeStage.INIT,error_code=ErrorCode.E_NO_DOCUMENTS)
        policy=error_policy or ErrorPolicy.RETRY_ONCE_THEN_SKIP
        out=Path(ensure_word_compatible_path(output_path,stage=MergeStage.OUTPUT_REPLACE,label='Izvades ceļš'))
        try: out.parent.mkdir(parents=True,exist_ok=True)
        except OSError as exc:
            raise MergeError(f'Nevar izveidot izvades mapi: {out.parent}',stage=MergeStage.OUTPUT_REPLACE,
                             error_code=ErrorCode.E_OUTPUT_LOCKED,original=exc,context={'output_path':str(out)}) from exc
        # DATU DROŠĪBA: izvades fails nedrīkst būt viens no avotiem (source ir immutable),
        # citādi merge pārrakstītu lietotāja dokumentu.
        self._assert_output_is_not_source(out,eligible)
        temp=out.with_name(f'{out.stem}.{uuid.uuid4().hex[:8]}.tmp{out.suffix or ".docx"}')
        pythoncom=win32com=None; session=word=master=None
        merged=0; errors=[]; normalized=[]; unlocked=[]; add_attempts=0; add_warning=None
        created_unlocked=[]
        self._created_normalized=[]
        total=len(eligible)
        try:
            pythoncom,win32com=_import_com()
            pythoncom.CoInitialize()
            self._info('Word COM inicializēts; apvienojamie dokumenti: %d',total)
            session=self._open_session(win32com)
            word=session.word
            master,add_attempts,add_report=self._create_master(word)
            if add_report is not None:
                add_warning=f'Documents.Add() kļūda tika kompensēta: {add_report.summary}'
                self._warn('%s',add_warning)
            for i,item in enumerate(eligible,1):
                item.merge_status=DocumentStatus.MERGING
                ok,report,used_norm,used_unlock=self._merge_item(word,master,item,i,separator,policy,
                                                                 insert_titles=insert_titles,
                                                                 unlock_protected_view=unlock_protected_view,
                                                                 created_unlocked=created_unlocked)
                if ok:
                    merged+=1; item.merge_status=DocumentStatus.MERGED
                    if used_norm: normalized.append(item.source_path)
                    if used_unlock: unlocked.append(item.source_path)
                else:
                    item.merge_status=DocumentStatus.ERROR
                    errors.append(self._error_entry(item,report,used_normalized=used_norm,used_unlocked=used_unlock))
                if progress is not None:
                    try: progress(i,total,item)
                    except Exception as exc:  # noqa: BLE001 - progress nedrīkst nogāzt merge
                        self._warn('Progress callback kļūda: %r',exc)
            self._save_master(master,temp,out)
            master=None
            if unlock_protected_view:
                self._unlock_output(out)
        finally:
            self._cleanup(master,session,pythoncom,temp,created_unlocked)
        result={'ok':not errors,'output_path':str(out),'merged':merged,'errors':errors,'normalized':normalized,
                'unlocked':unlocked,'documents_add_attempts':add_attempts,'warning':add_warning,'total':total,
                'word_owned':(session.owned if session is not None else None),
                'word_dispatch_method':(session.method if session is not None else None)}
        self._info('Merge pabeigts: %d/%d apvienoti, kļūdas: %d',merged,total,len(errors))
        return result

    def _assert_output_is_not_source(self,out,items):
        """Bloķē merge, ja izvades fails ir viens no avota failiem (datu drošība)."""
        try:
            output_key=os.path.normcase(absolute_path(out))
            for item in items:
                if os.path.normcase(absolute_path(item.source_path))==output_key:
                    raise MergeError(
                        f'Izvades fails ir viens no avota failiem: {out}. Izvēlies citu izvades ceļu — '
                        'avota faili netiek pārrakstīti.',
                        stage=MergeStage.OUTPUT_REPLACE,error_code=ErrorCode.E_OUTPUT_IS_SOURCE,
                        context={'output_path':str(out),'source_path':item.source_path})
        except MergeError:
            raise
        except Exception as exc:  # noqa: BLE001 - ceļu salīdzināšana nedrīkst nogāzt merge
            self._warn('Izvades/avota ceļu salīdzināšana neizdevās: %r',exc)

    def _unlock_output(self,out):
        """Noņem Protected View marķējumu izvades DOCX (kā pārbaudītajā skriptā)."""
        zone_removed,read_only_removed=unlock_file(out)
        if zone_removed or read_only_removed:
            self._info('Izvade atbloķēta | %s (Zone.Identifier=%s, Read-only=%s)',out,zone_removed,read_only_removed)
        return zone_removed,read_only_removed

    def _cleanup(self,master,session,pythoncom,temp,created_unlocked=None):
        """Aizver Word un atbrīvo COM (FinalReleaseComObject + [GC]::Collect ekvivalents).

        LIETOTĀJA jau atvērtā Word sesija (session.owned=False) NETIEK aizvērta —
        bez šīs aizsardzības `DispatchEx -> Dispatch` fallback laikā varēja tikt
        aizvērts lietotāja atvērtais dokuments (datu zudums).
        """
        try:
            if master is not None: master.Close(False)
        except Exception as exc:  # noqa: BLE001
            self._warn('Master.Close kļūda (%s): %r',MergeStage.CLEANUP.value,exc)
        try:
            if session is not None: session.close(self.logger)
        except Exception as exc:  # noqa: BLE001
            self._warn('Word sesijas aizvēršana neizdevās (%s): %r',MergeStage.CLEANUP.value,exc)
        try:
            if pythoncom is not None: pythoncom.CoUninitialize()
        except Exception as exc:  # noqa: BLE001
            self._warn('CoUninitialize kļūda (%s): %r',MergeStage.CLEANUP.value,exc)
        try:
            if temp is not None and Path(temp).exists(): Path(temp).unlink()
        except OSError as exc:
            self._warn('Temp faila dzēšana neizdevās (%s): %r',MergeStage.CLEANUP.value,exc)
        session=None; master=None; pythoncom=None
        self._collect_com_garbage()
        for path in list(created_unlocked or []):
            self._remove_temp_file(path,'Atbloķētās kopijas')
        if self.cleanup_normalized:
            for path in list(getattr(self,'_created_normalized',[]) or []):
                self._remove_temp_file(path,'Normalizētā faila')
        self._created_normalized=[]

    def _remove_temp_file(self,path,label):
        """Dzēš tikai šī darba gaitā radīto temp failu; kļūda tiek tikai reģistrēta."""
        try:
            candidate=Path(path)
            if candidate.exists(): candidate.unlink()
        except OSError as exc:
            self._warn('%s dzēšana neizdevās (%s): %r',label,MergeStage.CLEANUP.value,exc)

    def _open_session(self,win32com):
        """Word sesija ar atkārtojumu, Dispatch fallback un īpašumtiesībām."""
        try:
            return open_word_session(win32com,log=self.logger)
        except Exception as exc:  # noqa: BLE001
            report=self._log_report('WORD DISPATCH ERROR',
                                    capture_exception(exc,stage=MergeStage.WORD_DISPATCH,error_code=ErrorCode.E_WORD_DISPATCH,
                                                      context={'attempts':WORD_DISPATCH_ATTEMPTS,
                                                               'methods':list(configured_dispatch_methods())}))
            raise WordUnavailableError(f'Word COM nav pieejams: {report.message}',original=exc) from exc

    def _create_master(self,word):
        """Izveido tukšu master dokumentu. Atgriež (master, mēģinājumi, diagnozes reports|None).

        Pirmā Documents.Add() Word pusē var atgriezt vienreizēju converter kļūdu
        (WRD6ER32.CNV), tāpēc izsaukums tiek atkārtots. Ja tas nepalīdz, master
        tiek atvērts no ģenerēta tukša DOCX (Documents.Open).
        """
        last=None
        for attempt in range(1,self.add_attempts+1):
            try:
                master=word.Documents.Add()
                if attempt>1:
                    self._warn('Documents.Add() izdevās %d. mēģinājumā; iepriekšējā kļūda: %s',attempt,
                               capture_exception(last).summary if last is not None else '-')
                return master,attempt,None
            except Exception as exc:  # noqa: BLE001 - COM kļūda, mēģinām vēlreiz
                last=exc
                self._warn('Documents.Add() mēģinājums %d/%d neizdevās: %r',attempt,self.add_attempts,exc)
                time.sleep(DOCUMENTS_ADD_RETRY_DELAY)
        report=self._log_report('DOCUMENTS ADD ERROR',
                                capture_exception(last,stage=MergeStage.DOCUMENTS_ADD,error_code=ErrorCode.E_DOCUMENTS_ADD))
        try:
            blank=self._ensure_blank_docx()
            master=word.Documents.Open(str(blank),ReadOnly=False,AddToRecentFiles=False,ConfirmConversions=False)
            self._warn('Master izveidots no tukša DOCX (fallback): %s',blank)
            return master,self.add_attempts,report
        except Exception as fallback_exc:  # noqa: BLE001
            raise MergeError(
                f'Nevar izveidot jaunu Word dokumentu pēc {self.add_attempts} mēģinājumiem: {report.message}',
                stage=MergeStage.DOCUMENTS_ADD,error_code=ErrorCode.E_DOCUMENTS_ADD,original=last,
                context={'fallback_error':repr(fallback_exc)}) from last

    def _merge_item(self,word,master,item,index,separator,policy,insert_titles=False,
                    unlock_protected_view=True,created_unlocked=None):
        """Atgriež (ok, report|None, used_normalized, used_unlocked). Nekad nepazaudē kļūdu tekstu."""
        source,path_error=self._source_path(item)
        if path_error is not None:
            self._log_report('MERGE SKIP',path_error)
            return False,path_error,False,False
        used_unlocked=False
        if unlock_protected_view:
            source,used_unlocked=self._word_safe_source(source,item,created_unlocked)
        needs_separator=index>1 and separator is not None
        separator_done=False; title_done=False
        first_exception=None
        try:
            separator_done,title_done=self._open_document_block(master,item,separator,needs_separator,
                                                                insert_titles,separator_done,title_done)
            self._insert_source(master,source)
            self._info('MERGE OK | %s',source)
            return True,None,False,used_unlocked
        except Exception as exc:  # noqa: BLE001
            # except-variable tiek izdzēsta bloka beigās, tāpēc izņēmums tiek
            # saglabāts atsevišķā mainīgajā (citādi vēlāk UnboundLocalError).
            first_exception=exc
            first=self._log_report(
                f'INSERTFILE ERROR | {source}',
                capture_exception(exc,stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE,
                                  context={'path':source,'item_id':item.id}))
        try:
            normalized_path=self._normalize_source(word,source,item)
        except Exception as norm_exc:  # noqa: BLE001
            norm_report=self._log_report(
                f'NORMALIZE ERROR | {source}',
                capture_exception(norm_exc,stage=MergeStage.NORMALIZE,error_code=ErrorCode.E_NORMALIZE,
                                  context={'path':source,'item_id':item.id}))
            combined=capture_exception(MergeError(
                f'InsertFile un normalizācija neizdevās: {first.message} | {norm_report.summary}',
                stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE,original=first_exception,
                context={'normalization_error':norm_report.summary,'path':source}))
            return False,combined,True,used_unlocked
        if policy==ErrorPolicy.STOP_ON_ERROR:
            raise MergeError(f'InsertFile neizdevās (STOP_ON_ERROR): {first.message}',
                             stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE,
                             original=first_exception,context={'path':source})
        try:
            # Atdalītājs/virsraksts tiek ievietots TIKAI vienu reizi: ja pirmais
            # mēģinājums tos jau ievietoja, atkārtoti tos nedublē.
            separator_done,title_done=self._open_document_block(master,item,separator,needs_separator,
                                                                insert_titles,separator_done,title_done)
            self._insert_source(master,str(normalized_path))
            self._info('MERGE OK (normalizēts) | %s -> %s',source,normalized_path.name)
            return True,None,True,used_unlocked
        except Exception as second_exc:  # noqa: BLE001
            self._log_report(f'INSERTFILE ERROR (normalizēts) | {normalized_path}',
                             capture_exception(second_exc,stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE,
                                               context={'path':str(normalized_path),'item_id':item.id}))
            combined=capture_exception(MergeError(
                f'InsertFile neizdevās arī ar normalizēto failu: {str(second_exc)}',
                stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INSERT_FILE,original=second_exc,
                context={'first_error':first.summary,'path':source,'normalized_path':str(normalized_path)}))
            return False,combined,True,used_unlocked

    def _open_document_block(self,master,item,separator,needs_separator,insert_titles,
                             separator_done=False,title_done=False):
        """Ievieto atdalītāju un virsrakstu pirms dokumenta.

        Atgriež (separator_done, title_done), lai pēc neizdevušā InsertFile
        mēģinājuma tie netiktu ievietoti otrreiz.
        """
        if needs_separator and not separator_done:
            self._insert_separator(master,separator); separator_done=True
        if insert_titles and not title_done:
            title_done=self._insert_title(master,item)
        return separator_done,title_done

    def _word_safe_source(self,source,item,created_unlocked=None):
        """Word COM drošs avota ceļš.

        Ja avotam ir Mark-of-the-Web (Protected View), Word COM saņem atbloķētu
        KOPIJU — oriģinālais avots netiek mainīts (immutable) un pēc darba
        kopija tiek izdzēsta.
        """
        if not has_zone_identifier(source):
            return source,False
        stem=Path(item.filename or source).stem or 'documents'
        target=self._unlocked_dir()/f'{stem}__{str(item.id)[:8]}{Path(source).suffix or ".docx"}'
        try:
            copy_unlocked(source,target)
        except OSError as exc:
            self._warn('Atbloķētas kopijas izveide neizdevās (%s): %r; izmanto oriģinālu',source,exc)
            return source,False
        if created_unlocked is not None: created_unlocked.append(target)
        self._warn('Avotam ir Protected View marķējums; izmanto atbloķētu kopiju | %s -> %s',source,target)
        return str(target),True

    def _source_path(self,item):
        """Absolūts, Word COM derīgs avota ceļš. Atgriež (path, error_report|None)."""
        raw=str(item.source_path or '').strip()
        try:
            resolved=ensure_word_compatible_path(raw,stage=MergeStage.INSERT_FILE,label='Avota ceļš')
        except Exception as exc:  # noqa: BLE001
            return raw,capture_exception(exc,stage=MergeStage.INSERT_FILE,error_code=ErrorCode.E_INVALID_PATH,
                                        context={'item_id':item.id})
        if not Path(resolved).is_file():
            return resolved,capture_exception(MergeError(
                f'Avota fails neeksistē: {resolved}',stage=MergeStage.INSERT_FILE,
                error_code=ErrorCode.E_INVALID_PATH,context={'item_id':item.id}))
        return resolved,None

    def _insert_source(self,master,source):
        """InsertFile uz dokumenta beigām. COM atsauce tiek atbrīvota uzreiz."""
        rng=None
        try:
            try:
                rng=master.Range(master.Content.End-1,master.Content.End-1)
            except Exception as exc:  # noqa: BLE001
                raise MergeError(f'Nevar izveidot Range dokumentā: {exc}',stage=MergeStage.RANGE,
                                 error_code=ErrorCode.E_INSERT_FILE,original=exc,context={'path':source}) from exc
            rng.InsertFile(FileName=source,ConfirmConversions=False,Link=False,Attachment=False)
        finally:
            rng=None  # atbrīvo COM atsauci (FinalReleaseComObject ekvivalents Python pusē)

    def _insert_title(self,master,item):
        """Faila nosaukums kā trekna virsraksta rindkopa (kā pārbaudītajā skriptā).

        Virsraksts nav kritisks: ja tas neizdodas, dokumentu joprojām apvieno.
        Atgriež True, ja virsraksts tika ievietots.
        """
        title=str(Path(item.source_path or '').stem or item.filename or '').strip()
        if not title: return False
        rng=None
        try:
            try:
                rng=master.Range(master.Content.End-1,master.Content.End-1)
            except Exception as exc:  # noqa: BLE001
                raise MergeError(f'Nevar izveidot Range virsrakstam: {exc}',stage=MergeStage.RANGE,
                                 error_code=ErrorCode.E_INSERT_FILE,original=exc,
                                 context={'path':item.source_path}) from exc
            rng.Text=f'{title}\r'
            rng.Style=WD_STYLE_NORMAL
            rng.Font.Bold=True
            rng.Font.Italic=False
            rng.ParagraphFormat.SpaceBefore=0
            rng.ParagraphFormat.SpaceAfter=6
            rng.ParagraphFormat.KeepWithNext=True
            rng.Collapse(WD_COLLAPSE_END)
            rng.Font.Bold=False
            rng.Font.Italic=False
            rng.ParagraphFormat.KeepWithNext=False
            return True
        except Exception as exc:  # noqa: BLE001 - virsraksts nav kritisks
            self._warn('Virsraksta ievietošana neizdevās (%s): %r',MergeStage.INSERT_FILE.value,exc)
            return False
        finally:
            rng=None

    def _insert_separator(self,master,separator):
        code=_SEPARATOR_BREAK_CODES.get(separator)
        if code is None: return
        try:
            master.Range(master.Content.End-1,master.Content.End-1).InsertBreak(code)
        except Exception as exc:  # noqa: BLE001
            raise MergeError(f'InsertBreak neizdevās: {exc}',stage=MergeStage.SEPARATOR,
                             error_code=ErrorCode.E_INSERT_FILE,original=exc) from exc

    def _normalize_source(self,word,source,item):
        """Atver avotu ar Word COM ReadOnly un saglabā normalizētu DOCX.

        Oriģinālais fails netiek modificēts (ReadOnly=True, AddToRecentFiles=False).
        """
        stem=Path(item.filename or source).stem or 'documents'
        target=self._normalized_dir()/f'{stem}__{item.id[:8]}.normalized.docx'
        doc=None
        try:
            try:
                doc=word.Documents.Open(source,ReadOnly=True,AddToRecentFiles=False,
                                        ConfirmConversions=False,NoEncodingDialog=True)
            except Exception as exc:  # noqa: BLE001
                raise MergeError(f'Nevar atvērt avotu ar Word COM: {exc}',stage=MergeStage.DOCUMENT_OPEN,
                                 error_code=ErrorCode.E_DOCUMENT_OPEN,original=exc,context={'path':source}) from exc
            try:
                doc.SaveAs2(str(target),FileFormat=WD_FORMAT_XML_DOCUMENT)
            except Exception as exc:  # noqa: BLE001
                raise MergeError(f'Nevar saglabāt normalizēto DOCX: {exc}',stage=MergeStage.NORMALIZE,
                                 error_code=ErrorCode.E_NORMALIZE,original=exc,
                                 context={'path':source,'normalized_path':str(target)}) from exc
            self._info('NORMALIZED | %s -> %s',source,target)
            self._created_normalized.append(target)
            return target
        finally:
            if doc is not None:
                try: doc.Close(False)
                except Exception as exc:  # noqa: BLE001
                    self._warn('Normalizētā dokumenta Close kļūda: %r',exc)

    def _save_master(self,master,temp,out):
        try:
            master.SaveAs2(str(temp),FileFormat=WD_FORMAT_XML_DOCUMENT)
        except Exception as exc:  # noqa: BLE001
            report=self._log_report('SAVEAS ERROR',
                                    capture_exception(exc,stage=MergeStage.SAVE_AS,error_code=ErrorCode.E_SAVE_AS,
                                                      context={'temp_path':str(temp),'output_path':str(out)}))
            raise MergeError(f'Izvades failu nevar saglabāt: {report.message}',stage=MergeStage.SAVE_AS,
                             error_code=ErrorCode.E_SAVE_AS,original=exc,
                             context={'temp_path':str(temp),'output_path':str(out)}) from exc
        try:
            master.Close(False)
        except Exception as exc:  # noqa: BLE001
            self._warn('Master.Close pēc SaveAs2 kļūda: %r',exc)
        if out.exists():
            try: shutil.copy2(out,out.with_suffix(out.suffix+'.bak'))
            except OSError as exc: self._warn('Backup kopija neizdevās: %r',exc)
        try:
            os.replace(temp,out)
        except OSError as exc:
            locked=bool(getattr(exc,'winerror',0) in (5,32,33))
            report=self._log_report('OUTPUT REPLACE ERROR',
                                    capture_exception(exc,stage=MergeStage.OUTPUT_REPLACE,
                                                      error_code=ErrorCode.E_OUTPUT_LOCKED if locked else ErrorCode.E_INTERNAL,
                                                      context={'temp_path':str(temp),'output_path':str(out)}))
            hint='Izvades fails ir aizņemts (atvērts Word/OneDrive). Aizver to un mēģini vēlreiz.' if locked else 'Nevar pārvietot izvades failu.'
            raise MergeError(f'{hint} {report.message}',stage=MergeStage.OUTPUT_REPLACE,
                             error_code=ErrorCode.E_OUTPUT_LOCKED if locked else ErrorCode.E_INTERNAL,
                             original=exc,context={'temp_path':str(temp),'output_path':str(out)}) from exc

    @staticmethod
    def _error_entry(item,report,used_normalized=False,used_unlocked=False):
        return {'id':item.id,'path':item.source_path,'error':(report.message or '(nav ziņas)'),
                'stage':report.stage,'error_code':report.error_code,'error_type':report.error_type,
                'repr':report.repr,'traceback':report.traceback_text,'used_normalized':used_normalized,
                'used_unlocked':used_unlocked}

