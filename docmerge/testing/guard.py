"""OFFLINE režīma aizsargs.

Testi NEDRĪKST izmantot Internetu, API, cloud, GitHub, telemetriju vai veikt
lejupielādes. Šis modulis to nodrošina tehniski: tīkla socket savienojumi tiek
bloķēti, un katrs mēģinājums tiek ierakstīts reportā.

Atļauts:
- lokālais Microsoft Word COM (LRPC, nav TCP);
- lokāli procesi (`tasklist`), failu sistēma, subprocess `python -m ...`.
"""
from __future__ import annotations

import os
import socket
from datetime import datetime

LOCAL_HOSTS={'localhost','127.0.0.1','::1','0.0.0.0',''}
_GUARD_STATE={'active':False,'blocked':[],'installed_at':None,'restore':[]}


class OfflineViolation(RuntimeError):
    """Tīkla piekļuves mēģinājums OFFLINE testa laikā."""


def _is_local(host):
    return str(host).strip().lower() in LOCAL_HOSTS


def _record(kind,target):
    entry={'kind':kind,'target':str(target),'at':datetime.now().isoformat(timespec='seconds')}
    _GUARD_STATE['blocked'].append(entry)
    return entry


def _blocked(kind,target):
    _record(kind,target)
    raise OfflineViolation(f'OFFLINE: tīkla piekļuve bloķēta ({kind} -> {target})')


def install_offline_guard(allow_localhost=True):
    """Uzstāda tīkla aizsargu. Atgriež stāvokļa dict (idempotents)."""
    if _GUARD_STATE['active']:
        return offline_status()
    restore=_GUARD_STATE['restore']
    restore.append(('socket.connect',socket.socket.connect,socket.socket))
    restore.append(('socket.connect_ex',socket.socket.connect_ex,socket.socket))
    restore.append(('socket.create_connection',socket.create_connection,socket))
    restore.append(('socket.getaddrinfo',socket.getaddrinfo,socket))
    original_connect=socket.socket.connect; original_connect_ex=socket.socket.connect_ex
    original_create=socket.create_connection; original_getaddrinfo=socket.getaddrinfo

    def guard_connect(self,address):
        host=address[0] if isinstance(address,(tuple,list)) and address else address
        if allow_localhost and _is_local(host): return original_connect(self,address)
        return _blocked('connect',address)

    def guard_connect_ex(self,address):
        host=address[0] if isinstance(address,(tuple,list)) and address else address
        if allow_localhost and _is_local(host): return original_connect_ex(self,address)
        return _blocked('connect_ex',address)

    def guard_create_connection(address,*args,**kwargs):
        host=address[0] if isinstance(address,(tuple,list)) and address else address
        if allow_localhost and _is_local(host): return original_create(address,*args,**kwargs)
        return _blocked('create_connection',address)

    def guard_getaddrinfo(host,*args,**kwargs):
        if allow_localhost and _is_local(host): return original_getaddrinfo(host,*args,**kwargs)
        return _blocked('getaddrinfo',host)

    socket.socket.connect=guard_connect
    socket.socket.connect_ex=guard_connect_ex
    socket.create_connection=guard_create_connection
    socket.getaddrinfo=guard_getaddrinfo
    try:
        import ssl
        restore.append(('ssl.SSLContext.wrap_socket',ssl.SSLContext.wrap_socket,ssl.SSLContext))
        original_wrap=ssl.SSLContext.wrap_socket
        def guard_wrap_socket(self,sock,*args,**kwargs):
            _record('ssl.wrap_socket',getattr(sock,'getpeername',lambda:'?')())
            raise OfflineViolation('OFFLINE: TLS savienojums bloķēts')
        ssl.SSLContext.wrap_socket=guard_wrap_socket
    except Exception:  # noqa: BLE001 - ssl var būt bez wrap_socket
        pass
    # Nekādas automātiskas atslēgu/lejupielāžu uzvednes.
    os.environ['PIP_NO_INDEX']='1'
    os.environ['GIT_TERMINAL_PROMPT']='0'
    os.environ['HTTP_PROXY']='http://127.0.0.1:9'
    os.environ['HTTPS_PROXY']='http://127.0.0.1:9'
    os.environ['DOC_MERGE_OFFLINE']='1'
    _GUARD_STATE['active']=True; _GUARD_STATE['installed_at']=datetime.now().isoformat(timespec='seconds')
    return offline_status()


def uninstall_offline_guard():
    """Atjauno oriģinālās funkcijas (testu beigās)."""
    for name,func,owner in reversed(_GUARD_STATE['restore']):
        try: setattr(owner,name.split('.')[-1],func)
        except Exception: pass
    _GUARD_STATE['restore']=[]
    _GUARD_STATE['active']=False
    return offline_status()


def offline_status():
    """Stāvoklis reportam (aizsargs aktīvs, bloķētie mēģinājumi)."""
    return {'active':bool(_GUARD_STATE['active']),'installed_at':_GUARD_STATE['installed_at'],
            'blocked_count':len(_GUARD_STATE['blocked']),'blocked':list(_GUARD_STATE['blocked']),
            'env':{'DOC_MERGE_OFFLINE':os.environ.get('DOC_MERGE_OFFLINE'),
                   'PIP_NO_INDEX':os.environ.get('PIP_NO_INDEX'),
                   'GIT_TERMINAL_PROMPT':os.environ.get('GIT_TERMINAL_PROMPT')}}


def offline_guard_active():
    return bool(_GUARD_STATE['active'])
