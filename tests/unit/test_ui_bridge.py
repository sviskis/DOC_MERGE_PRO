import threading

from docmerge.gui.ui_bridge import UiBridge


class FakeRoot:
    def __init__(self): self._seq=0; self._scheduled={}
    def after(self,ms,func):
        self._seq+=1; self._scheduled[self._seq]=func; return self._seq
    def after_cancel(self,token): self._scheduled.pop(token,None)
    def pending(self): return len(self._scheduled)
    def run_due(self,steps=1):
        for _ in range(steps):
            if not self._scheduled: return
            key=min(self._scheduled); func=self._scheduled.pop(key); func()


def test_post_from_background_thread_runs_only_on_main_thread():
    root=FakeRoot(); bridge=UiBridge(root); executed=[]
    def worker():
        for i in range(50): bridge.post(executed.append,i)
    thread=threading.Thread(target=worker); thread.start(); thread.join()
    assert executed==[]  # background thread neizpilda neko
    assert bridge.pending==50
    bridge.drain()
    assert executed==list(range(50))


def test_callbacks_execute_in_calling_thread_only():
    root=FakeRoot(); bridge=UiBridge(root); seen=set()
    def worker():
        for _ in range(20): bridge.post(lambda:seen.add(threading.get_ident()))
    thread=threading.Thread(target=worker); thread.start(); thread.join()
    bridge.drain()
    assert seen=={threading.get_ident()}


def test_callback_errors_are_forwarded_and_do_not_stop_drain():
    root=FakeRoot(); errors=[]; done=[]
    bridge=UiBridge(root,on_error=errors.append)
    def boom(): raise ValueError('UI callback boom')
    bridge.post(boom); bridge.post(done.append,'ok')
    handled=bridge.drain()
    assert handled==2
    assert done==['ok']
    assert len(errors)==1 and isinstance(errors[0],ValueError)


def test_start_schedules_pump_and_runs_queue():
    root=FakeRoot(); bridge=UiBridge(root,interval_ms=10); calls=[]
    bridge.start(); assert root.pending()==1
    bridge.post(calls.append,'a'); bridge.post(calls.append,'b')
    root.run_due(steps=1)
    assert calls==['a','b']


def test_stop_cancels_pump_and_rejects_new_posts():
    root=FakeRoot(); bridge=UiBridge(root,interval_ms=10); calls=[]
    bridge.start(); bridge.stop(); bridge.post(calls.append,'a')
    assert calls==[] and bridge.closed is True
    assert bridge.pending==0
    assert bridge.drain()==0


def test_post_before_start_is_kept_and_drained():
    root=FakeRoot(); bridge=UiBridge(root); calls=[]
    bridge.post(calls.append,'x'); bridge.start(); root.run_due(steps=1)
    assert calls==['x']
