"""Compatibility launcher.

Galvenais lietotāja launcher tagad ir `DOC_MERGE_PRO.py` (projekta root).
Abi palaiž vienu un to pašu GUI entry point (`docmerge.gui.main_window.run_app`),
tāpēc `app.py` tiek saglabāts saderībai ar esošajām instrukcijām/īsceļiem.
"""
from docmerge.gui.main_window import run_app

if __name__ == "__main__":
    run_app()
