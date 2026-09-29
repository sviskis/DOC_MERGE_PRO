"""Thread-safe tilts starp background thread un Tkinter.

Background thread NEDRĪKST tieši modificēt Tk widgetus. Viss, kas maina UI
(log, Treeview, status, messagebox, progress), tiek nodots caur `post()` un
izpildīts Tk galvenajā thread ar `root.after()` pump palīdzību.
"""
import queue


class UiBridge:
    def __init__(self,root,interval_ms=50,on_error=None):
        self._root=root
        self._interval=int(interval_ms)
        self._queue=queue.Queue()
        self._after_id=None
        self._closed=False
        self._on_error=on_error

    def start(self):
        if self._closed or self._after_id is not None: return
        try: self._after_id=self._root.after(self._interval,self._pump)
        except Exception: self._closed=True

    def stop(self):
        self._closed=True
        if self._after_id is not None:
            try: self._root.after_cancel(self._after_id)
            except Exception: pass
            self._after_id=None

    def post(self,func,*args,**kwargs):
        """Droši izsaucams no jebkura thread. Atgriež False, ja tilts ir slēgts."""
        if self._closed or func is None: return False
        self._queue.put((func,args,kwargs)); return True

    @property
    def pending(self): return self._queue.qsize()

    @property
    def closed(self): return self._closed

    def drain(self,limit=None):
        """Izpilda uzkrātos callbackus. Drīkst izsaukt tikai Tk galvenais thread."""
        handled=0
        while limit is None or handled<limit:
            try: func,args,kwargs=self._queue.get_nowait()
            except queue.Empty: break
            handled+=1
            try: func(*args,**kwargs)
            except Exception as exc:  # noqa: BLE001 - UI callback kļūda netiek pazaudēta
                if self._on_error is not None:
                    try: self._on_error(exc)
                    except Exception: pass
        return handled

    def _pump(self):
        self._after_id=None
        self.drain()
        if not self._closed:
            try: self._after_id=self._root.after(self._interval,self._pump)
            except Exception: self._closed=True
