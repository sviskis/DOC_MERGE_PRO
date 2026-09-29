"""Temporary: isolated real-engine cases. Not production code.

usage: python _diag_case.py <case>
cases: minimal | copies | fdrive-local | fdrive-target
"""
import sys, shutil, traceback
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DIR = Path(r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\Izejmaterials\PY_materiali")
WORK = Path("runtime/_diag").resolve()
MANIFEST_SIX = [
    "Debesskr\u0101pis.docx",
    "D\u0101rzs.docx",
    "D\u0101vana bez maksas.docx",
    "D\u0101vana.docx",
    "D\u0101vanas.docx",
    "D\u0101vanu maiss.docx",
]


def case_minimal():
    import pythoncom, win32com.client

    pythoncom.CoInitialize()
    word = master = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        print("DispatchEx OK")
        word.Visible = False
        print("Visible=0 OK")
        word.DisplayAlerts = 0
        print("DisplayAlerts=0 OK")
        master = word.Documents.Add()
        print("Documents.Add OK", master.Name)
        rng = master.Range(master.Content.End - 1, master.Content.End - 1)
        print("Range OK")
        rng.Text = "Hello"
        print("Text OK")
        master.Close(False)
        master = None
    except Exception as e:  # noqa: BLE001
        print("MINIMAL FAILED:", repr(e))
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
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _run_engine(sources, out_path, label):
    from docmerge.domain.models import DocumentItem
    from docmerge.core.preflight import preflight_all
    from docmerge.engines.word_com import WordComEngine

    print(f"--- {label} ---")
    items = [DocumentItem(str(s)) for s in sources]
    for it in items:
        it.hydrate_from_path()
    items = preflight_all(items)
    for it in items:
        print("  preflight:", it.health_status.value, "|", it.filename, "|", it.errors, it.warnings, "| eligible:", it.eligible_for_merge)
    try:
        r = WordComEngine().merge(items, str(out_path))
        print("  RESULT:", r)
        print("  output exists:", Path(out_path).exists())
    except Exception as e:  # noqa: BLE001
        print("  ENGINE RAISED:", repr(e))
        traceback.print_exc()


def case_copies():
    (WORK / "sources").mkdir(parents=True, exist_ok=True)
    srcs = [p for p in sorted(SRC_DIR.glob("*.docx"), key=lambda p: p.name) if p.stem][:3]
    copies = []
    for p in srcs:
        dst = (WORK / "sources" / p.name).resolve()
        shutil.copy2(p, dst)
        copies.append(dst)
    _run_engine(copies, (WORK / "engine_copies.docx").resolve(), "3 local copies")


def case_fdrive(local_output):
    srcs = [SRC_DIR / n for n in MANIFEST_SIX if (SRC_DIR / n).is_file()]
    print("six sources found:", len(srcs))
    out = (WORK / "engine_fdrive.docx").resolve() if local_output else Path(
        r"F:\DIZAINA FOLDERIS\DIZANA DISKS DZEEST 2023\Munise\2026\test\diag_out.docx"
    )
    _run_engine(srcs, out, f"real F: sources -> {out}")


if __name__ == "__main__":
    case = sys.argv[1]
    if case == "minimal":
        case_minimal()
    elif case == "copies":
        case_copies()
    elif case == "fdrive-local":
        case_fdrive(True)
    elif case == "fdrive-target":
        case_fdrive(False)
    else:
        raise SystemExit(f"unknown case {case}")
