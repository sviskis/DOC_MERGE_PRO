"""OFFLINE testu ieeja.

`run_offline_test(full=False, soak=False)` uzstāda offline aizsargu, izpilda
attiecīgo komplektu (QUICK / FULL / SOAK), ieraksta reportus mapē `reports/`
un atgriež kopsavilkumu.
"""
from __future__ import annotations

import platform
import subprocess
import sys
from pathlib import Path

from docmerge.testing.guard import install_offline_guard, offline_status, uninstall_offline_guard
from docmerge.testing.harness import FAIL, PASS, Suite

PROJECT_ROOT=Path(__file__).resolve().parent.parent.parent


def git_commit(short=True):
    """Īsais (vai pilnais) Git commit hash; None, ja nav repo."""
    try:
        args=['git','rev-parse','--short','HEAD'] if short else ['git','rev-parse','HEAD']
        cp=subprocess.run(args,cwd=str(PROJECT_ROOT),capture_output=True,text=True,timeout=20)
        if cp.returncode==0: return cp.stdout.strip()
    except (OSError,subprocess.SubprocessError): pass
    return None


def git_dirty():
    """True, ja Git darba koks nav tīrs."""
    try:
        cp=subprocess.run(['git','status','--porcelain'],cwd=str(PROJECT_ROOT),capture_output=True,text=True,timeout=30)
        if cp.returncode==0: return bool((cp.stdout or '').strip())
    except (OSError,subprocess.SubprocessError): pass
    return None


def word_status():
    """Word COM statuss (lokāls, atļauts OFFLINE režīmā)."""
    try:
        from docmerge.engines.word_com import system_check
        return system_check()
    except Exception as exc:  # noqa: BLE001 - statuss tiek ziņots, nevis nomests
        return {'word_com':False,'error':repr(exc),'word_version':None}


def environment_meta(word_info=None):
    word_info=word_info or {}
    return {'python':sys.version.split()[0],'platform':platform.platform(),'executable':sys.executable,
            'project_root':str(PROJECT_ROOT),'cwd':str(Path.cwd()),
            'word_com':word_info.get('word_com'),'word_version':word_info.get('word_version'),
            'word_owned':word_info.get('word_owned'),'leftover_winword':word_info.get('leftover_winword'),
            'git_commit':git_commit(),'git_dirty':git_dirty(),'offline':True,
            'started_at':None}


def run_offline_test(full=False, soak=False, report_dir='reports', work_dir=None, doc_count=None):
    """Izpilda OFFLINE testu komplektu un ieraksta reportus."""
    # Word probi tiek veikti tikai FULL/SOAK; QUICK testi ir ātri un neprasa Word.
    word_info=word_status() if (full or soak) else {'word_com':None,'word_version':None}
    guard=install_offline_guard()
    suite=None; md_path=json_path=None
    try:
        if soak:
            from docmerge.testing.soak import build_soak_suite
            suite=build_soak_suite(report_dir=report_dir,work_dir=work_dir,word_info=word_info,doc_count=doc_count)
        elif full:
            from docmerge.testing.full import build_full_suite
            suite=build_full_suite(report_dir=report_dir,work_dir=work_dir,word_info=word_info)
        else:
            from docmerge.testing.quick import build_quick_suite
            suite=build_quick_suite(report_dir=report_dir,work_dir=work_dir)
        suite.meta.update(environment_meta(word_info))
        suite.meta['started_at']=suite.started_at.isoformat(timespec='seconds')
        suite.meta['offline_guard_active']=guard.get('active')
        suite.meta['offline_blocked_attempts']=guard.get('blocked_count')
        kind='SOAK' if soak else ('OFFLINE' if full else 'OFFLINE_QUICK')
        if full and not soak: kind='OFFLINE'
        md_path,json_path=suite.write_reports(kind)
    finally:
        guard_after=offline_status()
        uninstall_offline_guard()
    summary=suite.summary()
    summary['report_md']=str(md_path) if md_path else None
    summary['result_json']=str(json_path) if json_path else None
    summary['offline_guard']={'active':guard.get('active'),'blocked_count':guard.get('blocked_count'),
                              'blocked':guard.get('blocked')}
    if guard_after.get('blocked_count'):
        summary.setdefault('notes',[]).append(f'OFFLINE bloķēja {guard_after["blocked_count"]} tīkla mēģinājumu')
    return summary


__all__=['run_offline_test','Suite','PASS','FAIL','git_commit','word_status','environment_meta']
