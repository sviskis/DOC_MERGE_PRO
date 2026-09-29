"""Temporary: probe Word Documents.Add() behaviour in a fresh process."""
import sys, traceback
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\Izejmaterials\PY_materiali\Debesskr\u0101pis.docx")


def t(label, fn):
    try:
        r = fn()
        print(f"[OK  ] {label} -> {r!r}")
        return True, r
    except Exception as e:  # noqa: BLE001
        print(f"[FAIL] {label} -> {repr(e)}")
        return False, None


def main():
    import pythoncom, win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        t("Documents.Count (before)", lambda: word.Documents.Count)
        t("Add() #1 no args", lambda: word.Documents.Add())
        t("Add() #2 retry", lambda: word.Documents.Add())
        t("Add(Template='') ", lambda: word.Documents.Add(""))
        t("Add(Template='Normal')", lambda: word.Documents.Add("Normal"))
        t("Options.ConfirmConversions read", lambda: word.Options.ConfirmConversions)
        t("Add() #3 after Options read", lambda: word.Documents.Add())
        ok, doc = t(f"Open({SRC.name})", lambda: word.Documents.Open(str(SRC), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False))
        if ok and doc is not None:
            t("  Close", lambda: doc.Close(False))
        t("Add() #4 after Open", lambda: word.Documents.Add())
        t("Add(NewTemplate=True)", lambda: word.Documents.Add("", True))
        t("Add(DocumentType=0,Visible=False)", lambda: word.Documents.Add("", False, 0, False))
        print("Documents.Count (after):", word.Documents.Count)
        t("word.Options.DefaultFilePath(wdDocumentsPath)", lambda: word.Options.DefaultFilePath(0))
        t("word.StartupPath", lambda: word.StartupPath)
        t("word.Options.ConfirmConversions (read again)", lambda: word.Options.ConfirmConversions)
        t("master SaveAs2 test", lambda: word.Documents.Add().SaveAs2(str(Path("runtime/_diag/add_test.docx").resolve()), FileFormat=12))
    except Exception as e:  # noqa: BLE001
        print("UNEXPECTED:", repr(e))
        traceback.print_exc()
    finally:
        try:
            for d in list(word.Documents):
                print("  closing leftover:", d.Name)
                d.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
