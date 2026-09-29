"""Temporary: exercise the real WordComEngine with absolute paths. Not production code."""
import sys, shutil, traceback
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DIR = Path(r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\Izejmaterials\PY_materiali")
WORK = Path("runtime/_diag").resolve()
NORM = WORK / "normalized"


def step(name, fn):
    try:
        r = fn()
        print(f"[OK  ] {name} -> {r!r}")
        return True, r
    except Exception as e:  # noqa: BLE001
        print(f"[FAIL] {name}")
        print(f"       repr : {e!r}")
        return False, None


def main():
    import pythoncom
    import win32com.client

    srcs = [p for p in sorted(SRC_DIR.glob("*.docx"), key=lambda p: p.name) if p.stem][:3]
    print("ABSOLUTE sources:")
    for p in srcs:
        print("   ", str(p))
    (WORK / "sources").mkdir(parents=True, exist_ok=True)
    NORM.mkdir(parents=True, exist_ok=True)
    copies = []
    for p in srcs:
        dst = (WORK / "sources" / p.name).resolve()
        shutil.copy2(p, dst)
        copies.append(dst)

    pythoncom.CoInitialize()
    word = master = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        step("word.Visible=False", lambda: setattr(word, "Visible", False))
        step("word.DisplayAlerts=0", lambda: setattr(word, "DisplayAlerts", 0))
        step("word.Options.ConfirmConversions (read)", lambda: word.Options.ConfirmConversions)
        step("word.Options.ConfirmConversions=False", lambda: setattr(word.Options, "ConfirmConversions", False))
        for c in copies:
            ok, doc = step(f"Open(abs) {c.name}", lambda c=c: word.Documents.Open(str(c), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False))
            if ok and doc is not None:
                step("  Close", lambda doc=doc: doc.Close(False))
        ok, master = step("Documents.Add()", lambda: word.Documents.Add())
        for i, c in enumerate(copies, 1):
            if i > 1:
                step(f"InsertBreak(7) #{i}", lambda: master.Range(master.Content.End - 1, master.Content.End - 1).InsertBreak(7))
            step(f"InsertFile(abs) #{i} {c.name}", lambda c=c: master.Range(master.Content.End - 1, master.Content.End - 1).InsertFile(FileName=str(c), ConfirmConversions=False, Link=False, Attachment=False))
        tmp = (WORK / "out.tmp.docx").resolve()
        step(f"SaveAs2(abs) {tmp}", lambda: master.SaveAs2(str(tmp), FileFormat=12))
        print("      temp exists:", tmp.exists(), tmp.stat().st_size if tmp.exists() else None)
        step("Close", lambda: master.Close(False))
        master = None
    finally:
        try:
            if master is not None:
                master.Close(False)
        except Exception:
            traceback.print_exc()
        try:
            if word is not None:
                word.Quit()
        except Exception:
            traceback.print_exc()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    print("\n=== REAL WordComEngine.merge with ABSOLUTE paths ===")
    from docmerge.domain.models import DocumentItem, Project
    from docmerge.core.preflight import preflight_all
    from docmerge.engines.word_com import WordComEngine

    items = [DocumentItem(str(c)) for c in copies]
    items[0].source_path = str(copies[0])
    for it in items:
        it.hydrate_from_path()
    items = preflight_all(items)
    for it in items:
        print("   preflight:", it.filename, it.health_status.value, it.errors, it.warnings)
    out = (WORK / "engine_out.docx").resolve()
    try:
        r = WordComEngine().merge(items, str(out))
        print("RESULT:", r)
    except Exception as e:  # noqa: BLE001
        print("ENGINE RAISED repr:", repr(e))
        print("ENGINE RAISED type:", type(e).__name__)
        traceback.print_exc()
    print("out exists:", out.exists(), out.stat().st_size if out.exists() else None)


if __name__ == "__main__":
    main()
