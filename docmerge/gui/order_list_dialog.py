"""CSV/JSON secības saraksta priekšskatījuma dialogs.

Rāda Order / Requested / Resolved file / Type / Status, kopsavilkumu
(Atrasti/Trūkst/Dublikāti/Neskaidri) un STRICT MODE slēdzi.
"""
import tkinter as tk
from tkinter import ttk

ROW_COLORS={'FOUND':'#e8f6e8','MISSING':'#fde8e8','AMBIGUOUS':'#fff6e0','DUPLICATE':'#fff0d0','IGNORED':'#efefef'}


class OrderListPreview(tk.Toplevel):
    """Modāls saraksta priekšskatījums; apstiprināšana izsauc `on_apply(matches, strict)`."""

    def __init__(self,parent,matches,summary,meta=None,on_apply=None,strict=False):
        super().__init__(parent)
        self._matches=list(matches); self._summary=dict(summary or {}); self._meta=dict(meta or {})
        self._on_apply=on_apply
        self.title('Kārtot pēc saraksta — priekšskatījums')
        self.geometry('1200x660'); self.minsize(920,480)
        self.transient(parent)
        try: self.grab_set()
        except Exception: pass
        self.strict=tk.BooleanVar(value=bool(strict))
        self._ui(); self._fill()
        self.protocol('WM_DELETE_WINDOW',self.destroy)

    def _ui(self):
        box=ttk.Frame(self,padding=10); box.pack(fill='both',expand=True)
        info=f"Saraksts: {self._meta.get('path') or '(nav)'}  |  Formāts: {self._meta.get('format') or '-'}  |  " \
             f"Kodējums: {self._meta.get('encoding') or '-'}  |  Ieraksti: {self._meta.get('count',len(self._matches))}"
        ttk.Label(box,text=info,font=('Segoe UI',10,'bold')).pack(anchor='w')
        body=ttk.Panedwindow(box,orient='vertical'); body.pack(fill='both',expand=True,pady=(8,0))
        top=ttk.Frame(body); bottom=ttk.Frame(body); body.add(top,weight=4); body.add(bottom,weight=1)
        cols=('n','requested','resolved','type','status'); heads=['Order','Requested','Resolved file','Type','Status']
        widths=[60,300,520,90,110]
        self.tree=ttk.Treeview(top,columns=cols,show='headings',selectmode='browse')
        for c,t,w in zip(cols,heads,widths):
            self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor='w')
        for status,color in ROW_COLORS.items(): self.tree.tag_configure(status.lower(),background=color)
        ys=ttk.Scrollbar(top,orient='vertical',command=self.tree.yview); self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side='left',fill='both',expand=True); ys.pack(side='right',fill='y')
        self.summary=tk.StringVar(); ttk.Label(bottom,textvariable=self.summary,justify='left',
                                               font=('Consolas',10)).pack(anchor='w')
        self.problems=tk.Text(bottom,height=6,wrap='word'); self.problems.pack(fill='both',expand=True,pady=(4,0))
        self.problems.configure(state='disabled')
        row=ttk.Frame(box); row.pack(fill='x',pady=(8,0))
        ttk.Checkbutton(row,text='STRICT MODE (MISSING / AMBIGUOUS / DUPLICATE bloķē merge)',
                        variable=self.strict).pack(side='left')
        ttk.Button(row,text='Aizvērt',command=self.destroy).pack(side='right',padx=(6,0))
        ttk.Button(row,text='Piemērot secību',command=self._apply).pack(side='right')

    def _fill(self):
        for match in self._matches:
            self.tree.insert('','end',values=(match.order,match.requested,
                                              match.resolved_path or (match.message or '-'),
                                              match.detected_format or '-',match.status.value),
                             tags=(match.status.value.lower(),))
        s=self._summary
        self.summary.set(f"Atrasti: {s.get('found',0)}   Trūkst: {s.get('missing',0)}   "
                         f"Dublikāti: {s.get('duplicate',0)}   Neskaidri: {s.get('ambiguous',0)}   "
                         f"Izlaisti (nav Word): {s.get('ignored',0)}   Kopā: {s.get('total',0)}")
        lines=list(self._meta.get('problems') or [])+list(s.get('problems') or [])
        text='\n'.join(lines) if lines else 'Problēmu nav.'
        self.problems.configure(state='normal'); self.problems.delete('1.0','end')
        self.problems.insert('end',text); self.problems.configure(state='disabled')

    def _apply(self):
        if self._on_apply is not None:
            self._on_apply(self._matches,self.strict.get())
        self.destroy()
