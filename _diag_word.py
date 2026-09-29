"""Temporary Word COM step-by-step diagnostic. Not part of production code."""
import os, sys, shutil, traceback
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DIR = Path(r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\Izejmaterials\PY_materiali")
WORK = Path("runtime/_diag")
NORM = WORK / "normalized"


def step(name, fn):
    try:
        r = fn()
        print(f"[OK  ] {name} -> {r!r}")
        return True, r
    except Exception as e:  # noqa: BLE001 - diagnostic
        print(f"[FAIL] {name}")
        print(f"       repr      : {e!r}")
        print(f"       type      : {type(e).__name__}")
        print(f"       args      : {getattr(e, 'args', None)!r}")
        print(f"       excepinfo : {getattr(e, 'excepinfo', None)!r}")
        print(f"       hresult   : {getattr(e, 'hresult', None)!r}")
        return False, None


def main():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    print("--- collected sources ---")
    if SRC_DIR.is_dir():
        srcs = sorted(SRC_DIR.glob("*.docx"), key=lambda p: p.name)[:3]
    else:
        srcs = []
    if not srcs:
        srcs = sorted(Path("runtime/_diag/samples").glob("*.docx"))[:3]
    print([str(p) for p in srcs])

    WORK.mkdir(parents=True, exist_ok=True)
    NORM.mkdir(parents=True, exist_ok=True)
    copies = []
    for p in srcs:
        dst = WORK / "sources" / p.name
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, dst)
        copies.append(dst)
    print("copies:", [str(c) for c in copies])

    word = None
    master = None
    try:
        step("CoInitialize", lambda: "done")
        ok, word = step("DispatchEx('Word.Application')", lambda: win32com.client.DispatchEx("Word.Application"))
        if not ok:
            return
        print("      word.Version =", word.Version)
        step("word.Visible=False", lambda: setattr(word, "Visible", False))
        step("word.DisplayAlerts=0", lambda: setattr(word, "DisplayAlerts", 0))
        step("Options.ConfirmConversions=False", lambda: setattr(word.Options, "ConfirmConversions", False))

        # Can Word even Open the source files read-only?
        for c in copies:
            ok, doc = step(f"Documents.Open({c.name}, ReadOnly)", lambda c=c: word.Documents.Open(str(c), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, NoEncodingDialog=True))
            if ok and doc is not None:
                step(f"  Open({c.name}) doc.Close(False)", lambda doc=doc: doc.Close(False))

        # InsertFile directly on the ORIGINAL F: path
        if srcs:
            ok, master = step("Documents.Add() [for direct F: InsertFile test]", lambda: word.Documents.Add())
            if ok:
                step("  Range direct-F InsertFile", lambda: master.Range(master.Content.End - 1, master.Content.End - 1).InsertFile(FileName=str(srcs[0]), ConfirmConversions=False, Link=False, Attachment=False))
                step("  master.Close(False)", lambda: master.Close(False))
                master = None

        ok, master = step("Documents.Add()", lambda: word.Documents.Add())
        if not ok:
            return
        for i, c in enumerate(copies, 1):
            ok, _ = step(f"Range({i}) create", lambda: master.Range(master.Content.End - 1, master.Content.End - 1))
            if i > 1:
                ok, _ = step(f"  Range({i}) InsertBreak(7)", lambda: master.Range(master.Content.End - 1, master.Content.End - 1).InsertBreak(7))
            step(f"  Range({i}) InsertFile({c.name})", lambda c=c: master.Range(master.Content.End - 1, master.Content.End - 1).InsertFile(FileName=str(c), ConfirmConversions=False, Link=False, Attachment=False))

        tmp = WORK / "out.tmp.docx"
        step("SaveAs2(temp, FileFormat=12)", lambda: master.SaveAs2(str(tmp), FileFormat=12))
        print("      temp exists:", tmp.exists(), "size:", tmp.stat().st_size if tmp.exists() else None)
        step("master.Close(False)", lambda: master.Close(False))
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


if __name__ == "__main__":
    main()
