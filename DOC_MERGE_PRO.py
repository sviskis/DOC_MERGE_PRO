"""DOC_MERGE_PRO — galvenais lietotāja launcher (GUI).

Palaišana:

    python DOC_MERGE_PRO.py

Tas palaiž to pašu GUI entry point kā `app.py` (compatibility launcher).
CLI režīmi (piem. `--offline-test`, `--merge-all`, `--system-check`) paliek
`python -m docmerge.cli.main ...`.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Launcher strādā arī tad, ja to palaiž no citas mapes vai ar pilnu ceļu.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docmerge.gui.main_window import APP_TITLE, run_app  # noqa: E402


def main():
    print(f'{APP_TITLE} — GUI tiek atvērts...', flush=True)
    run_app()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
