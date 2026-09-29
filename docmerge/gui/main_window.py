import os
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from pathlib import Path
import threading,traceback

from docmerge.domain.models import Project
from docmerge.domain.enums import SortMode,SeparatorMode
from docmerge.domain.errors import ErrorCode,MergeStage
from docmerge.core.scanner import scan_folder,add_files,filter_word_documents
from docmerge.core.sorter import sort_items
from docmerge.core.preflight import preflight_all
from docmerge.core.order_list import (apply_order, clear_order_list, entries_from_project, evaluate_order_list,
                                     export_order, load_order_list, match_entries, summarize)
from docmerge.persistence.project_store import save_project,load_project
from docmerge.engines.word_com import system_check
from docmerge.gui.order_list_dialog import OrderListPreview
from docmerge.gui.ui_bridge import UiBridge
from docmerge.reporting.error_report import (capture_exception,error_log_path,log_error_report,
                                             persist_error_log_safely,report_from_parts)
from docmerge.reporting.logger import create_logger
from docmerge.core.merge_service import MergeService
from docmerge.workers.merge_worker import run_merge_job

APP_TITLE='DOC_MERGE_PRO'
SORT_CHOICES=(('Pēc nosaukuma (dabiskā)',SortMode.NATURAL_NAME),
              ('Manuālā secība (↑/↓)',SortMode.MANUAL),
              ('Pēc saraksta (CSV/JSON)',SortMode.ORDER_LIST),
              ('Pēc izmaiņu laika ↑',SortMode.MODIFIED_DATE_ASC),
              ('Pēc izmaiņu laika ↓',SortMode.MODIFIED_DATE_DESC),
              ('Pēc mapes',SortMode.FOLDER_ORDER))
SORT_LABELS=[label for label,_ in SORT_CHOICES]


def sort_label(mode):
    for label,value in SORT_CHOICES:
        if value==mode: return label
    return SORT_LABELS[0]


def sort_mode_from_label(label):
    for text,value in SORT_CHOICES:
        if text==label: return value
    return SortMode.NATURAL_NAME


class MainWindow:
    """Tkinter GUI. Visas UI izmaiņas notiek TIKAI galvenajā thread (UiBridge)."""

    def __init__(self,root):
        self.root=root; root.title(APP_TITLE); root.geometry('1500x950'); root.minsize(1150,720)
        self.logger,self.log_path=create_logger(); self.error_log=error_log_path()
        self.project=Project()
        self.service=MergeService(self.logger,error_log_path=self.error_log)
        self.bridge=UiBridge(root,interval_ms=50,on_error=self._on_ui_callback_error)
        root.report_callback_exception=self._report_callback_exception
        root.protocol('WM_DELETE_WINDOW',self.on_close)
        self._merge_running=False
        self.order_matches=[]; self.order_summary=summarize([]); self.order_meta={}
        self._ui(); self.bridge.start(); self.refresh(); self.check_word()
        self._bind_shortcuts()

    def _bind_shortcuts(self):
        """Klaviatūras tausiņi manuālai secībai (drag/drop vietā, bez ārējām atkarībām)."""
        for seq,handler in (('<Alt-Up>',lambda e:self.move(-1)),('<Alt-Down>',lambda e:self.move(1)),
                            ('<Alt-Home>',lambda e:self.move_to(0)),('<Alt-End>',lambda e:self.move_to(len(self.project.items)-1))):
            try: self.root.bind(seq,handler)
            except Exception: pass

    # --------------------------------------------------------------- UI threading
    def post(self,func,*args,**kwargs):
        """Nodod UI izmaiņas galvenajam thread. Droši no jebkura thread."""
        return self.bridge.post(func,*args,**kwargs)

    def on_close(self):
        try: self.bridge.stop()
        finally: self.root.destroy()

    def _on_ui_callback_error(self,exc):
        report=capture_exception(exc,stage=MergeStage.INTERNAL,error_code=ErrorCode.E_INTERNAL)
        self._record_error(report,'UI CALLBACK ERROR')
        return report

    def _report_callback_exception(self,exc_type,exc_value,exc_tb):
        """Nekad nepazaudē Tk callback izņēmumu (nevis klusē kā noklusētais)."""
        report=capture_exception(exc_value,stage=MergeStage.INTERNAL,error_code=ErrorCode.E_INTERNAL)
        if report.traceback_text is None and exc_type is not None:
            report.traceback_text=''.join(traceback.format_exception(exc_type,exc_value,exc_tb))
        self._record_error(report,'TK CALLBACK ERROR')

    def _record_error(self,report,prefix):
        log_error_report(self.logger,report,prefix=prefix)
        persist_error_log_safely(report,log_path=self.error_log,logger=self.logger)
        try: self._append_log(report.full_text)
        except Exception: pass
        return report

    def _ui(self):
        s=ttk.Frame(self.root,padding=12); s.pack(fill='both',expand=True)
        h=ttk.Frame(s); h.pack(fill='x',pady=(0,10)); ttk.Label(h,text=APP_TITLE,font=('Segoe UI',22,'bold')).pack(side='left'); self.sys=tk.StringVar(value='Word: pārbauda...'); ttk.Label(h,textvariable=self.sys).pack(side='right')
        tb=ttk.Frame(s); tb.pack(fill='x',pady=(0,8))
        for t,c in [('Jauns',self.new),('Atvērt projektu',self.open_project),('Saglabāt projektu',self.save_project),('Pievienot failus',self.pick_files),('Pievienot mapi',self.pick_folder),('Izvades fails',self.pick_output),('Visus DOC/DOCX → viens DOCX',self.merge_all_word),('KĀRTOT PĒC SARAKSTA...',self.pick_order_list),('EKSPORTĒT SECĪBU',self.export_current_order)]: ttk.Button(tb,text=t,command=c).pack(side='left',padx=(0,6))
        opt=ttk.LabelFrame(s,text='Iestatījumi',padding=8); opt.pack(fill='x',pady=(0,8)); ttk.Label(opt,text='Secība:').grid(row=0,column=0,sticky='w'); self.sort=tk.StringVar(value=sort_label(SortMode.NATURAL_NAME)); cb=ttk.Combobox(opt,textvariable=self.sort,state='readonly',values=SORT_LABELS,width=28); cb.grid(row=0,column=1,padx=6); cb.bind('<<ComboboxSelected>>',lambda e:self.apply_sort()); ttk.Label(opt,text='Sadalītājs:').grid(row=0,column=2,sticky='w'); self.sep=tk.StringVar(value=SeparatorMode.PAGE_BREAK.value); ttk.Combobox(opt,textvariable=self.sep,state='readonly',values=[x.value for x in SeparatorMode],width=28).grid(row=0,column=3,padx=6); self.output=tk.StringVar(); ttk.Label(opt,text='Izvade:').grid(row=1,column=0,sticky='w',pady=(8,0)); ttk.Entry(opt,textvariable=self.output).grid(row=1,column=1,columnspan=3,sticky='ew',pady=(8,0)); opt.columnconfigure(3,weight=1); self.titles=tk.BooleanVar(value=self.project.options.insert_titles); ttk.Checkbutton(opt,text='Faila nosaukums kā trekns virsraksts',variable=self.titles).grid(row=2,column=0,columnspan=2,sticky='w',pady=(8,0)); self.unlock=tk.BooleanVar(value=self.project.options.unlock_protected_view); ttk.Checkbutton(opt,text='Noņemt Protected View marķējumu (Zone.Identifier)',variable=self.unlock).grid(row=2,column=2,columnspan=2,sticky='w',pady=(8,0)); self.strict=tk.BooleanVar(value=self.project.options.strict_order_mode); ttk.Checkbutton(opt,text='STRICT MODE (saraksta kļūdas bloķē merge)',variable=self.strict).grid(row=3,column=0,columnspan=2,sticky='w',pady=(8,0)); self.open_output=tk.BooleanVar(value=self.project.options.open_output_when_finished); ttk.Checkbutton(opt,text='OPEN OUTPUT WHEN FINISHED (atvērt izvadi Word)',variable=self.open_output).grid(row=3,column=2,columnspan=2,sticky='w',pady=(8,0))
        pane=ttk.Panedwindow(s,orient='horizontal'); pane.pack(fill='both',expand=True); left=ttk.Frame(pane); right=ttk.Frame(pane); pane.add(left,weight=4); pane.add(right,weight=1)
        box=ttk.LabelFrame(left,text='Failu saraksts',padding=5); box.pack(fill='both',expand=True); cols=('n','file','type','det','size','status','notes'); self.tree=ttk.Treeview(box,columns=cols,show='headings',selectmode='extended'); heads=['#','Fails','Tips','Atpazītais formāts','Izmērs','Statuss','Piezīmes']; widths=[45,300,70,135,90,100,320]
        for c,t,w in zip(cols,heads,widths): self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor='w')
        self.tree.tag_configure('disabled',background='#ededed')
        ys=ttk.Scrollbar(box,orient='vertical',command=self.tree.yview); self.tree.configure(yscrollcommand=ys.set); self.tree.pack(side='left',fill='both',expand=True); ys.pack(side='right',fill='y')
        r=ttk.Frame(left); r.pack(fill='x',pady=7); ttk.Button(r,text='↑ Augšup',command=lambda:self.move(-1)).pack(side='left',padx=3); ttk.Button(r,text='↓ Lejup',command=lambda:self.move(1)).pack(side='left',padx=3); ttk.Button(r,text='⤒ Sākumā',command=lambda:self.move_to(0)).pack(side='left',padx=3); ttk.Button(r,text='⤓ Beigās',command=lambda:self.move_to(len(self.project.items)-1)).pack(side='left',padx=3); ttk.Button(r,text='Noņemt',command=self.remove).pack(side='left',padx=3); ttk.Label(r,text='(Alt+↑/↓, Alt+Home/End)').pack(side='left',padx=(10,0))
        a=ttk.Frame(left); a.pack(fill='x'); ttk.Button(a,text='Pārbaudīt failus',command=self.preflight).pack(side='left',padx=3); ttk.Button(a,text='Dry Run',command=self.dry).pack(side='left',padx=3); self.merge_btn=ttk.Button(a,text='SĀKT APVIENOŠANU',command=self.merge); self.merge_btn.pack(side='left',padx=3); ttk.Button(a,text='Word tests',command=self.check_word).pack(side='left',padx=3); ttk.Button(a,text='Saraksta kopsavilkums',command=self.show_order_summary).pack(side='left',padx=3)
        pr=ttk.Frame(left); pr.pack(fill='x',pady=(7,0)); self.status=tk.StringVar(value='Gatavs'); ttk.Label(pr,textvariable=self.status,width=52,anchor='w').pack(side='left'); self.progress=tk.DoubleVar(value=0); ttk.Progressbar(pr,variable=self.progress,maximum=100).pack(side='left',fill='x',expand=True,padx=(6,0))
        sm=ttk.LabelFrame(right,text='Kopsavilkums',padding=8); sm.pack(fill='x',padx=(8,0)); self.summary=tk.StringVar(); ttk.Label(sm,textvariable=self.summary,justify='left').pack(anchor='w'); lg=ttk.LabelFrame(right,text='Žurnāls',padding=5); lg.pack(fill='both',expand=True,padx=(8,0),pady=(8,0)); self.log=tk.Text(lg,wrap='word'); self.log.pack(fill='both',expand=True)

    # ------------------------------------------------------------- UI (main thread)
    def _append_log(self,text):
        self.log.insert('end',str(text)+'\n'); self.log.see('end'); self.logger.info('%s',text)

    def ulog(self,text):
        """Droši izsaucams arī no background thread — raksta tikai galvenajā thread."""
        self.post(self._append_log,text)

    def _set_status(self,text): self.status.set(str(text))

    def _set_merge_running(self,running):
        self._merge_running=bool(running)
        try: self.merge_btn.configure(state='disabled' if running else 'normal')
        except Exception: pass

    def _error_text(self,report):
        text=getattr(report,'text',None)
        if not text or not str(text).strip(): text=str(report) or '(nav ziņas)'
        return str(text)

    def _show_background_error(self,report,title='Kļūda',popup=True):
        """Galvenajā thread: saglabā pilnu kļūdu logā + persistent failā + popup."""
        text=self._error_text(report)
        self._record_error(report,'GUI ERROR')
        self._set_status('Kļūda: '+getattr(report,'summary',str(report)))
        if popup:
            # Kļūdas teksts ir saglabāts mainīgajā PIRMS lambda (nekad None).
            self.root.after(0,lambda msg=text,t=title:messagebox.showerror(t,msg))

    # ------------------------------------------------------------------ project ops
    def sync(self):
        self.project.output_path=self.output.get().strip(); self.project.options.sort_mode=sort_mode_from_label(self.sort.get()); self.project.options.separator=SeparatorMode(self.sep.get())
        self.project.options.insert_titles=bool(self.titles.get()); self.project.options.unlock_protected_view=bool(self.unlock.get())
        self.project.options.strict_order_mode=bool(self.strict.get()); self.project.options.open_output_when_finished=bool(self.open_output.get())

    def _apply_options_to_ui(self):
        options=self.project.options
        self.output.set(self.project.output_path); self.sort.set(sort_label(options.sort_mode)); self.sep.set(options.separator.value)
        self.titles.set(options.insert_titles); self.unlock.set(options.unlock_protected_view)
        self.strict.set(options.strict_order_mode); self.open_output.set(options.open_output_when_finished)

    def new(self):
        self.project=Project(); self.output.set(''); self.progress.set(0); self.order_matches=[]; self.order_summary=summarize([]); self.order_meta={}
        self._apply_options_to_ui(); self.refresh()

    def open_project(self):
        p=filedialog.askopenfilename(filetypes=[('DOC Merge Project','*.docmerge.json'),('JSON','*.json')]);
        if p:
            self.project=load_project(p); self._apply_options_to_ui()
            if self.project.options.order_list_enabled: self._refresh_order_state()
            self.refresh(); self._append_log('Atvērts: '+p)

    def save_project(self):
        self.sync(); p=filedialog.asksaveasfilename(defaultextension='.docmerge.json',filetypes=[('DOC Merge Project','*.docmerge.json')]);
        if p: save_project(self.project,p); self._append_log('Saglabāts: '+p)

    def pick_folder(self):
        p=filedialog.askdirectory();
        if not p:return
        have={str(Path(x.source_path).resolve()) for x in self.project.items}
        for x in scan_folder(p):
            if str(Path(x.source_path).resolve()) not in have:self.project.items.append(x)
        self.apply_sort(); self._reapply_order_list_if_active(); self._append_log('Mape: '+p)

    def pick_files(self):
        ps=filedialog.askopenfilenames(filetypes=[('Dokumenti','*.doc *.docx *.docm *.dot *.dotx *.dotm *.rtf *.odt *.txt'),('Visi','*.*')]);
        if not ps:return
        have={str(Path(x.source_path).resolve()) for x in self.project.items}
        for x in add_files(list(ps)):
            if str(Path(x.source_path).resolve()) not in have:self.project.items.append(x)
        self.apply_sort(); self._reapply_order_list_if_active()

    def pick_output(self):
        p=filedialog.asksaveasfilename(defaultextension='.docx',filetypes=[('Word DOCX','*.docx')]);
        if p:self.output.set(p); self.project.output_path=p

    def merge_all_word(self):
        """Visi .doc/.docx no mapes -> viens DOCX (viena klikšķa plūsma)."""
        folder=filedialog.askdirectory(title='Mape ar DOC/DOCX failiem')
        if not folder: return
        items=filter_word_documents(scan_folder(folder,recursive=True))
        if not items:
            return messagebox.showwarning('Apvienošana','Mapē nav .doc/.docx failu.')
        if self.project.items and not messagebox.askyesno('Ātrā apvienošana','Pašreizējais failu saraksts tiks aizstāts ar mapes saturu. Turpināt?'):
            return
        out=filedialog.asksaveasfilename(title='Kur saglabāt apvienoto DOCX',defaultextension='.docx',
                                        initialfile=f'{Path(folder).name}_APVIENOTS.docx',filetypes=[('Word DOCX','*.docx')])
        if not out: return
        self.project=Project(items=items,output_path=out)
        self.project.options.sort_mode=SortMode.NATURAL_NAME; self.project.options.separator=SeparatorMode.PAGE_BREAK
        self.project.options.insert_titles=True; self.project.options.unlock_protected_view=True
        self.order_matches=[]; self.order_summary=summarize([]); self.order_meta={}
        self._apply_options_to_ui(); self.refresh()
        self._append_log(f'Ātrā apvienošana: {len(items)} DOC/DOCX -> {out}')
        self.merge()

    # ------------------------------------------------------------- order list (CSV/JSON)
    def pick_order_list(self):
        """'Kārtot pēc saraksta…' — CSV/JSON izvēle, matching un priekšskatījums."""
        p=filedialog.askopenfilename(title='Izvēlies secības sarakstu (CSV vai JSON)',
                                     filetypes=[('CSV/JSON saraksts','*.csv *.json'),('CSV','*.csv'),('JSON','*.json'),('Visi','*.*')])
        if not p: return
        try:
            entries,meta=load_order_list(p)
        except Exception as exc:  # noqa: BLE001 - kļūda tiek saglabāta un parādīta
            self._show_background_error(capture_exception(exc,stage=MergeStage.ORDER_LIST),'Saraksta kļūda'); return
        if not self.project.items:
            return messagebox.showwarning('Saraksts','Vispirms pievieno failus (Pievienot mapi/failus).')
        matches=match_entries(entries,self.project.items)
        summary=summarize(matches)
        self.order_matches=matches; self.order_summary=summary; self.order_meta=meta
        for problem in meta.get('problems') or []: self._append_log('Saraksts: '+problem)
        self._append_log(f"Saraksts {Path(p).name}: {summary['found']} atrasti, {summary['missing']} trūkst, "
                         f"{summary['duplicate']} dublikāti, {summary['ambiguous']} neskaidri, {summary['ignored']} izlaisti")
        self._open_order_preview(matches,summary,meta,entries,p)

    def _open_order_preview(self,matches,summary,meta,entries,order_path):
        def on_apply(applied_matches,strict):
            self._apply_order_list(entries,applied_matches,strict,order_path)
        OrderListPreview(self.root,matches,summary,meta,on_apply=on_apply,strict=bool(self.strict.get()))

    def _apply_order_list(self,entries,matches,strict,order_path):
        """Piemēro saraksta secību projektam (absolūtā secība, bez natural sort)."""
        self.strict.set(bool(strict))
        ordered,others=apply_order(self.project,entries,matches,strict=strict,order_path=order_path)
        self.sort.set(sort_label(SortMode.ORDER_LIST)); self.refresh()
        summary=summarize(matches)
        self._append_log(f"Saraksta secība piemērota: {len(ordered)} faili apvienošanai, {len(others)} ārpus saraksta "
                         f"(izlaisti). Trūkst: {summary['missing']}, dublikāti: {summary['duplicate']}, neskaidri: {summary['ambiguous']}")
        if strict and summary['blocking_count']:
            messagebox.showwarning('STRICT MODE',
                f"STRICT MODE ir ieslēgts un {summary['blocking_count']} saraksta ieraksti ir problemātiski — "
                'merge tiks bloķēts, līdz tie tiek izlaboti (vai STRICT izslēgts).\n\n'
                +'\n'.join(summary['problems'][:10]))

    def _refresh_order_state(self):
        """Pārrēķina saglabātā saraksta matching (pēc projekta atvēršanas/failu maiņas)."""
        matches,summary=evaluate_order_list(self.project)
        self.order_matches=matches; self.order_summary=summary
        self.order_meta={'path':self.project.options.order_list_path,'format':'-','count':len(matches)}
        return matches,summary

    def show_order_summary(self):
        """Parāda pēdējo/pārrēķināto saraksta kopsavilkumu (Atrasti/Trūkst/Dublikāti/Neskaidri)."""
        self.sync()
        if self.project.options.order_list_enabled:
            matches,summary=self._refresh_order_state()
            entries=None
        else:
            matches,summary,entries=self.order_matches,self.order_summary,None
        if not matches:
            return messagebox.showinfo('Saraksts','Order saraksts vēl nav izmantots.')
        self._open_order_preview(matches,summary,self.order_meta,entries,self.project.options.order_list_path)

    def export_current_order(self):
        """'Eksportēt secību' — saglabā pašreizējo GUI secību CSV vai JSON."""
        if not self.project.items:
            return messagebox.showwarning('Eksports','Nav failu, ko eksportēt.')
        p=filedialog.asksaveasfilename(title='Kur saglabāt secību',defaultextension='.csv',
                                       initialfile='order.csv',filetypes=[('CSV','*.csv'),('JSON','*.json')])
        if not p: return
        try:
            target,fmt=export_order(p,self.project.items)
        except Exception as exc:  # noqa: BLE001
            self._show_background_error(capture_exception(exc,stage=MergeStage.ORDER_LIST),'Eksporta kļūda'); return
        self._append_log(f'Secība eksportēta ({fmt}): {target}')

    def apply_sort(self):
        """Secības izvēle. 'Pēc saraksta' prasa CSV/JSON un NEveic natural sort."""
        mode=sort_mode_from_label(self.sort.get())
        if mode==SortMode.ORDER_LIST:
            if self.project.options.order_list_enabled or self.order_matches:
                self.project.items=sort_items(self.project.items,SortMode.ORDER_LIST); self.refresh()
                self._append_log('Secība no saglabātā saraksta (natural sort netiek pielietots).')
            else:
                self.pick_order_list()
            return
        self.project.options.sort_mode=mode
        self.project.items=sort_items(self.project.items,mode); self.refresh()

    def _selected_indexes(self):
        return sorted(int(x) for x in self.tree.selection())

    def _mark_manual(self):
        for n,x in enumerate(self.project.items,1): x.manual_order=n
        self.project.options.sort_mode=SortMode.MANUAL; self.sort.set(sort_label(SortMode.MANUAL))

    def move(self,d):
        """Pārvieto atzīmētos failus par d pozīcijām (atbalsta bloku ar vairākām atzīmēm)."""
        idx=self._selected_indexes()
        if not idx or d==0: return
        items=self.project.items; new=[]
        if d<0:
            idx=[i for i in idx if i>0]
            if not idx: return
            for i in idx: items[i-1],items[i]=items[i],items[i-1]
            new=[i-1 for i in idx]
        else:
            idx=[i for i in idx if i<len(items)-1]
            if not idx: return
            for i in reversed(idx): items[i],items[i+1]=items[i+1],items[i]
            new=[i+1 for i in idx]
        self._mark_manual(); self.refresh(); self.tree.selection_set([str(i) for i in new])

    def move_to(self,index):
        """Pārvieto atzīmētos uz saraksta sākumu (0) vai beigām (pēdējais indekss)."""
        idx=self._selected_indexes()
        if not idx: return
        items=list(self.project.items); block=[items[i] for i in idx]; chosen=set(idx)
        rest=[item for i,item in enumerate(items) if i not in chosen]
        to_top=index<=0
        self.project.items=(block+rest) if to_top else (rest+block)
        start=0 if to_top else len(rest)
        self._mark_manual(); self.refresh()
        self.tree.selection_set([str(i) for i in range(start,start+len(block))])

    def remove(self):
        for i in sorted([int(x) for x in self.tree.selection()],reverse=True): self.project.items.pop(i)
        self._reapply_order_list_if_active(); self.refresh()

    def _reapply_order_list_if_active(self):
        """Ja order saraksts ir aktīvs, pēc failu izmaiņām to pārrēķina (absolūtā secība)."""
        if not self.project.options.order_list_enabled: return False
        matches,summary=self._refresh_order_state()
        apply_order(self.project,entries_from_project(self.project),matches,
                    strict=self.project.options.strict_order_mode,order_path=self.project.options.order_list_path)
        self._append_log(f"Saraksts pārrēķināts: {summary['found']} atrasti, {summary['missing']} trūkst")
        return True

    def preflight(self):
        self._append_log('Preflight sākts...'); self._set_status('Preflight...')
        def w():
            try:
                preflight_all(self.project.items)
                self.post(self.refresh); self.post(self._append_log,'Preflight pabeigts.'); self.post(self._set_status,'Preflight pabeigts')
            except Exception as exc:  # noqa: BLE001 - pilna kļūda tiek saglabāta
                self.post(self._show_background_error,
                          capture_exception(exc,stage=MergeStage.PREFLIGHT,error_code=ErrorCode.E_VALIDATION),'Preflight kļūda')
        threading.Thread(target=w,daemon=True).start()

    def dry(self):
        self.sync(); data=sort_items(self.project.items,self.project.options.sort_mode); ready=sum(x.eligible_for_merge for x in data); warn=sum(bool(x.warnings) for x in data); err=sum(bool(x.errors) for x in data); messagebox.showinfo('Dry Run',f'Kopā: {len(data)}\nGatavi: {ready}\nBrīdinājumi: {warn}\nKļūdas: {err}\nSecība: {self.project.options.sort_mode.value}')

    # ---------------------------------------------------------------------- merge
    def merge(self):
        self.sync()
        if not self.project.output_path:
            return messagebox.showerror('Kļūda','Izvēlies izvades DOCX.')
        if self._merge_running:
            return messagebox.showinfo('Merge','Apvienošana jau notiek...')
        # STRICT MODE tiek pārbaudīts arī GUI pusē, lai lietotājs uzreiz redzētu iemeslu.
        if self.project.options.order_list_enabled and self.project.options.strict_order_mode:
            _,summary=self._refresh_order_state()
            if summary['blocking_count']:
                return messagebox.showerror('STRICT MODE',
                    f"Merge bloķēts: {summary['blocking_count']} problemātiski saraksta ieraksti.\n\n"+
                    '\n'.join(summary['problems'][:15]))
        if self.project.options.order_list_enabled:
            summary=self.order_summary
            self._append_log(f"Order saraksts: {summary.get('found',0)} atrasti, {summary.get('missing',0)} trūkst, "
                             f"{summary.get('duplicate',0)} dublikāti, {summary.get('ambiguous',0)} neskaidri")
        self._set_merge_running(True); self.progress.set(0); self._set_status('Merge notiek...'); self._append_log('Merge sākts...')
        def worker():
            run_merge_job(self.service,self.project,
                          on_log=lambda msg:self.post(self._append_log,msg),
                          on_progress=self._on_merge_progress,
                          on_success=self._on_merge_success,
                          on_error=self._on_merge_error,
                          on_refresh=lambda:self.post(self.refresh),
                          logger=self.logger,error_log_path=self.error_log)
        threading.Thread(target=worker,daemon=True).start()

    def _on_merge_progress(self,done,total,item):
        """Background thread -> tikai post(); Tk widgetus neskar."""
        self.post(self._apply_progress,done,total,getattr(item,'filename',''))

    def _apply_progress(self,done,total,name):
        self.progress.set(0.0 if not total else (done/total)*100.0)
        self._set_status(f'Merge {done}/{total}: {name}')

    def _on_merge_success(self,result):
        self.post(self._show_merge_success,result)

    def _show_merge_success(self,result):
        self._set_merge_running(False); self.progress.set(100.0)
        errors=list(result.get('errors') or []); normalized=list(result.get('normalized') or []); unlocked=list(result.get('unlocked') or [])
        order=list((result.get('order_list') or {}).get('summary') or {})
        diff=result.get('manifest_diff') or {}
        lines=[f"Apvienots: {result.get('merged')}/{result.get('total')}",f"Izvade: {result.get('output_path')}"]
        if normalized: lines.append(f'Normalizēti faili: {len(normalized)}')
        if unlocked: lines.append(f'Atbloķētas avota kopijas (Protected View): {len(unlocked)}')
        if order: lines.append(f"Order saraksts: {order.get('found',0)} atrasti, {order.get('missing',0)} trūkst, "
                               f"{order.get('duplicate',0)} dublikāti, {order.get('ambiguous',0)} neskaidri")
        if diff: lines.append(f"Manifests: {diff.get('unchanged',0)} UNCHANGED, {diff.get('changed',0)} CHANGED, "
                              f"{diff.get('missing',0)} MISSING, {diff.get('added',0)} ADDED")
        warning=result.get('warning')
        if warning: lines.append('Brīdinājums: '+str(warning))
        if errors:
            lines.append(f'Kļūdas: {len(errors)}')
            for e in errors[:5]: lines.append(f"  [{e.get('stage')} | {e.get('error_code')}] {e.get('path')}\n    {str(e.get('error'))[:300]}")
        text='\n'.join(lines); self._append_log(text); self._set_status('Gatavs' if not errors else f'Pabeigts ar {len(errors)} kļūdām')
        self._open_output_if_requested(result)
        if errors: self.root.after(0,lambda msg=text:messagebox.showwarning('Merge pabeigts ar kļūdām',msg))
        else: self.root.after(0,lambda msg=text:messagebox.showinfo('Gatavs',msg))

    def _open_output_if_requested(self,result):
        """OPEN OUTPUT WHEN FINISHED: atver izvades DOCX (tikai ja tas eksistē un nav kļūdu)."""
        if not bool(self.open_output.get()): return None
        if result.get('errors'): return None
        path=str(result.get('output_path') or '')
        if not path or not Path(path).is_file(): return None
        try:
            if os.name=='nt': os.startfile(path)  # noqa: S606 - lietotāja izvēlēts fails
            else: self._append_log('Automātiska atvēršana atbalstīta tikai Windows.')
        except OSError as exc:
            self._record_error(capture_exception(exc,stage=MergeStage.INTERNAL),'OUTPUT OPEN ERROR')
            self._append_log('Izvadi neizdevās atvērt: '+repr(exc)); return None
        self._append_log('Izvade atvērta: '+path); return path

    def _on_merge_error(self,report):
        """Background thread -> ErrorReport jau ir materializēts (nekad None)."""
        self.post(self._merge_failed,report)

    def _merge_failed(self,report):
        self._set_merge_running(False); self._set_status('Merge neizdevās')
        self._show_background_error(report,'Merge kļūda')

    # ----------------------------------------------------------------- word check
    def check_word(self):
        def w():
            try:
                r=system_check()
            except Exception as exc:  # noqa: BLE001
                self.post(self._show_background_error,capture_exception(exc,stage=MergeStage.SYSTEM_CHECK),'Word pārbaude'); return
            ok=bool(r.get('word_com'))
            text=f"Word: OK ({r.get('word_version')})" if ok else f"Word COM: ERROR — {r.get('error')}"
            self.post(self.sys.set,text); self.post(self._append_log,text)
            if not ok:
                report=report_from_parts(r.get('error') or '(nav ziņas)',error_type=r.get('error_type') or 'WordComError',
                                         stage=MergeStage.SYSTEM_CHECK,error_code=ErrorCode.E_WORD_UNAVAILABLE,
                                         context={'pywin32':r.get('pywin32'),'windows':r.get('windows')},
                                         repr_text=r.get('error_repr'),traceback_text=r.get('error_traceback'))
                self.post(self._show_background_error,report,'Word COM kļūda')
        threading.Thread(target=w,daemon=True).start()

    # -------------------------------------------------------------------- refresh
    def refresh(self):
        for x in self.tree.get_children():self.tree.delete(x)
        for i,x in enumerate(self.project.items):
            notes='; '.join(x.errors+x.warnings)
            if not x.enabled: notes=('IZSLĒGTS (ārpus order saraksta); '+notes).strip('; ')
            self.tree.insert('', 'end', iid=str(i), values=(x.manual_order,x.filename,x.extension[1:].upper(),x.detected_format,self.size(x.size_bytes),x.health_status.value,notes),
                             tags=('' if x.enabled else 'disabled',))
        order_line=''
        if self.project.options.order_list_enabled:
            summary=self.order_summary or {}
            order_line=(f"\nOrder saraksts: {self.project.options.order_list_path or '(saglabāts projektā)'}"
                        f"\n  Atrasti: {summary.get('found',0)}  Trūkst: {summary.get('missing',0)}  "
                        f"Dublikāti: {summary.get('duplicate',0)}  Neskaidri: {summary.get('ambiguous',0)}"
                        f"\n  STRICT MODE: {'ON' if self.project.options.strict_order_mode else 'OFF'}")
        self.summary.set(f"Kopā: {len(self.project.items)}\nOK: {sum(x.health_status.value=='OK' for x in self.project.items)}\nBrīdinājumi: {sum(bool(x.warnings) for x in self.project.items)}\nKļūdas: {sum(bool(x.errors) for x in self.project.items)}{order_line}\nLog: {self.log_path}\nKļūdu log: {self.error_log}")

    @staticmethod
    def size(n):
        v=float(n)
        for u in ('B','KB','MB','GB'):
            if v<1024 or u=='GB':return f'{v:.0f} {u}' if u=='B' else f'{v:.1f} {u}'
            v/=1024


def run_app():
    root=tk.Tk(); MainWindow(root); root.mainloop()
