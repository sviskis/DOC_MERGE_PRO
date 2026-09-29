"""Launcher un BAT failu testi (DOC_MERGE_PRO nosaukums un vienots GUI entry point)."""
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]
BAT=ROOT/'BAT'


def test_doc_merge_pro_launcher_exists_and_uses_run_app():
    launcher=ROOT/'DOC_MERGE_PRO.py'
    assert launcher.is_file(),'DOC_MERGE_PRO.py nav projekta root mapē'
    text=launcher.read_text(encoding='utf-8')
    assert 'from docmerge.gui.main_window import APP_TITLE, run_app' in text
    assert 'run_app()' in text


def test_app_py_is_compatibility_launcher():
    text=(ROOT/'app.py').read_text(encoding='utf-8')
    assert 'run_app' in text,'app.py jāpalaiž tas pats GUI entry point'


def test_gui_title_is_doc_merge_pro():
    from docmerge.gui.main_window import APP_TITLE
    assert APP_TITLE=='DOC_MERGE_PRO'


def test_main_window_imports_without_tk_root():
    from docmerge.gui import main_window
    assert hasattr(main_window,'MainWindow')
    assert hasattr(main_window,'sort_label') and hasattr(main_window,'sort_mode_from_label')


def test_sort_mode_order_list_is_available_in_gui_choices():
    from docmerge.domain.enums import SortMode
    from docmerge.gui.main_window import SORT_CHOICES, sort_label, sort_mode_from_label
    labels=[label for label,_ in SORT_CHOICES]
    assert sort_label(SortMode.ORDER_LIST) in labels
    assert 'sarakst' in sort_label(SortMode.ORDER_LIST).lower()
    assert sort_mode_from_label(sort_label(SortMode.ORDER_LIST))==SortMode.ORDER_LIST
    assert sort_mode_from_label('nezinams')==SortMode.NATURAL_NAME


def test_bat_launchers_exist_and_are_portable():
    expected=['START_DOC_MERGE_PRO.bat','TEST_DOC_MERGE_PRO.bat','SYSTEM_CHECK.bat',
              'OFFLINE_TEST.bat','OFFLINE_TEST_FULL.bat','SOAK_TEST.bat']
    for name in expected:
        path=BAT/name
        assert path.is_file(),f'Trūkst BAT/{name}'
        text=path.read_text(encoding='utf-8',errors='replace')
        assert '%~dp0' in text,f'{name} neizmanto %~dp0'
        assert 'cd /d "%ROOT%"' in text,f'{name} nepāriet uz projekta root'
        assert '.venv' in text,f'{name} nepārbauda .venv'
    start=(BAT/'START_DOC_MERGE_PRO.bat').read_text(encoding='utf-8',errors='replace')
    assert 'DOC_MERGE_PRO.py' in start,'START BAT jāpalaiž DOC_MERGE_PRO.py'
    assert 'pause' in start,'START BAT nedrīkst aizvērties kļūdas gadījumā'


def test_offline_test_bat_files_reference_cli_flags():
    assert '--offline-test' in (BAT/'OFFLINE_TEST.bat').read_text(encoding='utf-8',errors='replace')
    assert '--offline-test-full' in (BAT/'OFFLINE_TEST_FULL.bat').read_text(encoding='utf-8',errors='replace')
    assert '--system-check' in (BAT/'SYSTEM_CHECK.bat').read_text(encoding='utf-8',errors='replace')
    assert '_diag_soak.py' in (BAT/'SOAK_TEST.bat').read_text(encoding='utf-8',errors='replace')


def test_gitignore_covers_runtime_and_documents():
    text=(ROOT/'.gitignore').read_text(encoding='utf-8')
    for pattern in ('__pycache__/','.pytest_cache/','.venv/','runtime/*','logs/*','reports/*','~$*','*.tmp','*.docx'):
        assert pattern in text,f'.gitignore trūkst {pattern}'
    for keep in ('runtime/.gitkeep','logs/.gitkeep','reports/.gitkeep'):
        assert (ROOT/keep).is_file(),f'Trūkst {keep}'
