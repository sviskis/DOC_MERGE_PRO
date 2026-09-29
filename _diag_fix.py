"""Temporary: verify Documents.Add retry + legacy .doc handling. Not production code."""
import sys, time, traceback
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DIR = Path(r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\Izejmaterials\PY_materiali")
WORK = Path("runtime/_diag").resolve()
SIX = [
    "Debesskr\u0101pis.docx", "D\u0101rzs.docx", "D\u0101vana bez maksas.docx",
    "D\u0101vana.docx", "D\u0101vanas.docx", "D\u0101vanu maiss.docx",
]


def add_with_retry(word, attempts=3):
    last = None
    for i in range(1, attempts + 1):
        try:
            doc = word.Documents.Add()
            print(f"  Add() attempt {i}: OK")
            return doc, i, None
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  Add() attempt {i}: FAIL {repr(e)[:110]}")
            time.sleep(0.25)
    return None, attempts, last


def merge_fixed(sources, out_path, label, separator_break=7):
    import pythoncom, win32com.client

    print(f"=== {label} ===")
    out = Path(out_path).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    temp = out.with_name(out.stem + ".tmp.docx")
    pythoncom.CoInitialize()
    word = master = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        master, attempts, err = add_with_retry(word)
        if master is None:
            print("  ABORT: Add failed after retries:", repr(err))
            return
        for i, src in enumerate(sources, 1):
            src = Path(src).resolve()
            print(f"  [{i}] {src.name} exists={src.exists()}")
            try:
                if i > 1:
                    master.Range(master.Content.End - 1, master.Content.End - 1).InsertBreak(separator_break)
                master.Range(master.Content.End - 1, master.Content.End - 1).InsertFile(
                    FileName=str(src), ConfirmConversions=False, Link=False, Attachment=False
                )
                print(f"      InsertFile OK (content end={master.Content.End})")
            except Exception as e:  # noqa: BLE001
                print(f"      InsertFile FAIL {repr(e)[:170]}")
                try:
                    norm = (WORK / "normalized" / (src.stem + ".normalized.docx")).resolve()
                    norm.parent.mkdir(parents=True, exist_ok=True)
                    d = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
                    d.SaveAs2(str(norm), FileFormat=12)
                    d.Close(False)
                    master.Range(master.Content.End - 1, master.Content.End - 1).InsertFile(
                        FileName=str(norm), ConfirmConversions=False, Link=False, Attachment=False
                    )
                    print(f"      normalized+InsertFile OK -> {norm.name} ({norm.stat().st_size} B)")
                except Exception as e2:  # noqa: BLE001
                    print(f"      normalisation FAIL {repr(e2)[:170]}")
        master.SaveAs2(str(temp), FileFormat=12)
        master.Close(False)
        master = None
        if out.exists():
            out.unlink()
        temp.rename(out)
        print(f"  OUTPUT {out.name} exists={out.exists()} size={out.stat().st_size}")
    except Exception as e:  # noqa: BLE001
        print("  FATAL:", repr(e))
        traceback.print_exc()
    finally:
        try:
            if master is not None:
                master.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def make_legacy_doc(src_docx, dst_doc):
    import pythoncom, win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Add()
        except Exception:  # noqa: BLE001
            d = word.Documents.Add()
        d.Content.InsertFile(str(Path(src_docx).resolve()), ConfirmConversions=False, Link=False, Attachment=False)
        d.SaveAs2(str(Path(dst_doc).resolve()), FileFormat=0)  # wdFormatDocument97
        d.Close(False)
        p = Path(dst_doc)
        print("  legacy .doc created:", p.exists(), p.stat().st_size if p.exists() else None)
    except Exception as e:  # noqa: BLE001
        print("  legacy creation failed:", repr(e))
        traceback.print_exc()
    finally:
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
    which = sys.argv[1] if len(sys.argv) > 1 else "six"
    if which == "six":
        srcs = [SRC_DIR / n for n in SIX if (SRC_DIR / n).is_file()]
        merge_fixed(srcs, WORK / "fixed_out.docx", f"6 real F: docx ({len(srcs)})")
    elif which == "legacy":
        s = SRC_DIR / SIX[0]
        if not s.is_file():
            s = sorted([p for p in SRC_DIR.glob("*.docx") if p.stem], key=lambda p: p.name)[0]
        doc = WORK / "legacy_sample.doc"
        WORK.mkdir(parents=True, exist_ok=True)
        make_legacy_doc(s, doc)
        merge_fixed([doc, s], WORK / "legacy_out.docx", "legacy .doc + docx")
    elif which == "normalize":
        s = SRC_DIR / SIX[0]
        merge_fixed([s, s], WORK / "norm_out.docx", "normalisation path (same file twice)")
