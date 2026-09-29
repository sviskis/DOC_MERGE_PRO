"""OFFLINE soak tests (unattended) — 100–300 DOCX, vairāki merge cikli.

Palaišana:

    python _diag_soak.py                 # 120 dokumenti
    python _diag_soak.py 300             # 300 dokumenti
    DOC_MERGE_SOAK_DOCS=200 python _diag_soak.py

Nedrīkst mākslīgi gaidīt 2 stundas: mērķis ir reāla slodze, nevis gaidīšana.
Reporti: reports/SOAK_TEST_REPORT.md un reports/SOAK_TEST_RESULT.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from docmerge.testing import run_offline_test  # noqa: E402


def main():
    doc_count = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else None
    summary = run_offline_test(soak=True, doc_count=doc_count)
    print(json.dumps({key: summary.get(key) for key in ('suite', 'status', 'passed', 'failed', 'warnings', 'skipped',
                                                        'tests_run', 'duration_s', 'report_md', 'result_json')},
                     ensure_ascii=False, indent=2))
    for result in summary.get('results', []):
        if result['status'] != 'PASS':
            print(f"[{result['status']}] {result['section']} / {result['name']}: {result.get('error') or result['detail']}")
    return 0 if summary.get('status') == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
