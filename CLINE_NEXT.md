# CLINE NEXT

## Pabeigts (salīdzināts ar pārbaudīto skriptu docx_atbloketajs + dzīvs Word tests)
- [x] Protected View atbloķēšana (`core/unlock.py`): Zone.Identifier + Read-only,
      avota immutable, izvade atbloķēta, kopijas `runtime/unlocked` tiek izdzēstas.
      Atklāts un izlabots: ADS dzēšana no Read-only faila dod PermissionError
      (winerror 5) → vispirms Read-only, tad plūsma; `copy2` pārnes ADS uz kopiju.
- [x] Absolūti ceļi Word COM (normalizētie/atbloķētie faili) — relatīvos Word
      risina pret `C:\WINDOWS\system32`.
- [x] `except ... as e` mainīgā izdzēšana bloka beigās (UnboundLocalError
      STOP_ON_ERROR un normalizācijas kļūdas ceļā) — izlabots.
- [x] Lappuses pārtraukums/virsraksts netiek pazaudēts vai dublēts, ja InsertFile
      nokrīt un tiek izmantota normalizācija.
- [x] Faila nosaukums kā trekns virsraksts (`MergeOptions.insert_titles`).
- [x] COM atsauču atbrīvošana + gc + CoFreeUnusedLibraries.
- [x] "Visi DOC/DOCX → viens DOCX": GUI poga + `--merge-all` CLI.
- [x] Dzīvs Word E2E (`_diag_e2e.py`): 2×.docx + 1×.doc ar Zone.Identifier →
      viens DOCX, 3 dokumenti, 2 lappušu pārtraukumi, virsraksti, .txt ignorēts,
      WINWORD.EXE nepaliek, izvade bez marķējuma.

## Tālāk
1. Palai `pytest -q` pirms katra posma (pašlaik 53 testi zaļi).
2. Pilns WordMergeWorker subprocess + MergeSupervisor; GUI nedrīkst tieši vadīt COM.
3. Timeout/kill/checkpoint/resume un `retry only errors`.
4. `capabilities.py` un normalizācijas adapteri DOC/ODT/RTF/TXT.
5. Persistent state.json, CSV/HTML atskaites.
6. Īsts Treeview drag-and-drop.
7. Offline, golden fidelity, package fidelity, fault injection, crash/resume, soak testi.
8. Source faili ir immutable.
9. Nekad neizmanto `python-docx` kā production merge engine.
10. Pirms katra posma parādi plānu un pēc tam testu rezultātus.

## Palīgrīki
- `_diag_e2e.py` — dzīvs Word tests (2 .docx + 1 .doc + .txt + Zone.Identifier).
- `_diag_gui.py` — GUI smoke tests (logi, opciju saglabāšana projektā).
