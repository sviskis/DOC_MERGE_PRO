# DOC_MERGE_PRO

**DOC_MERGE_PRO** apvieno daudzus `.doc` / `.docx` failus **vienā DOCX** ar
Microsoft Word COM (`Range.InsertFile`) — tāpat kā praksē pārbaudītais variants
(`docx_atbloketajs_old.py`), bet ar modulāru arhitektūru, CSV/JSON secību,
GUI, CLI un nopietniem offline testiem.

- GUI launcher: `DOC_MERGE_PRO.py` (vecais `app.py` paliek kā compatibility launcher).
- GUI nosaukums: **DOC_MERGE_PRO**.
- Source faili ir **immutable** — oriģinālie DOC/DOCX nekad netiek mainīti.
- Production merge **neizmanto python-docx** (tas ir tikai minimālu testu
  fixture ģenerēšanai).

## Instalācija

```powershell
pip install -r requirements.txt        # pywin32 (Windows) + pytest
python -m pytest -q
python -m docmerge.cli.main --system-check
```

## Palaišana

```powershell
python DOC_MERGE_PRO.py      # GUI
python app.py                # tas pats GUI (compatibility launcher)
```

vai dubultklikšķis uz `BAT\START_DOC_MERGE_PRO.bat`
(BAT faili strādā arī tad, ja tos palaiž no citas mapes).

| BAT fails | Ko dara |
| --- | --- |
| `BAT\START_DOC_MERGE_PRO.bat` | palaiž GUI (`DOC_MERGE_PRO.py`), kļūdas gadījumā logs neaizveras |
| `BAT\TEST_DOC_MERGE_PRO.bat` | `python -m pytest -q` |
| `BAT\SYSTEM_CHECK.bat` | `python -m docmerge.cli.main --system-check` |
| `BAT\OFFLINE_TEST.bat` | `--offline-test` (QUICK, bez Word) |
| `BAT\OFFLINE_TEST_FULL.bat` | `--offline-test-full` (Word E2E + fault injection) |
| `BAT\SOAK_TEST.bat` | `python _diag_soak.py` (100–300 DOCX, vairāki cikli) |

## DOC/DOCX apvienošana

GUI: `Pievienot mapi` / `Pievienot failus` → `Izvades fails` → **SĀKT APVIENOŠANU**.
Viena klikšķa režīms: **Visus DOC/DOCX → viens DOCX**.

CLI:

```powershell
python -m docmerge.cli.main --merge-all "C:\Mape" --output "C:\out\APVIENOTS.docx"
python -m docmerge.cli.main --merge-all "C:\Mape" --output out.docx --no-titles --no-unlock
python -m docmerge.cli.main --merge-all "C:\Mape" --output out.docx --order-list order.json --strict-order-list
python -m docmerge.cli.main --merge-all "C:\Mape" --export-order "C:\out\order.csv"
python -m docmerge.cli.main --merge-all "C:\Mape" --output out.docx --show-order-list
python -m docmerge.cli.main --merge-all "C:\Mape" --output out.docx --open-output
```

Katrs dokuments sākas jaunā lapā, un pirms tā tiek ievietots faila nosaukums
kā trekns virsraksts (izslēdzams ar `--no-titles`).
Protected View marķējums (`Zone.Identifier`) no avotiem netiek noņemts — Word
saņem atbloķētu **kopiju**, un marķējums tiek noņemts tikai no izvades.

## CSV secība

```csv
filename
01_ievads.docx
05_stasts.doc
02_nodala.docx
```

```csv
order,filename
1,01_ievads.docx
2,05_stasts.doc
3,02_nodala.docx
```

```csv
path
apaksmape\01_ievads.docx
apaksmape\05_stasts.doc
```

Atbalstīts: UTF-8, UTF-8 BOM, cp1257 (ja droši dekodējams), `,` `;` un tab
atdalītāji, `#` komentāri, tukšas rindas. Bojātas rindas tiek izlaistas un
ierakstītas kā problēmas (nekas netiek pazaudēts).

## JSON secība

```json
["01_ievads.docx", "05_stasts.doc", "02_nodala.docx"]
```

```json
{"files": ["01_ievads.docx", "05_stasts.doc"]}
```

```json
{"files": [{"order": 1, "filename": "01_ievads.docx"}, {"order": 2, "filename": "05_stasts.doc"}]}
```

**Saraksta secība ir absolūta** — natural sort pēc tā netiek pielietots
(piem. `C.docx, A.docx, B.docx` paliek tieši šādā secībā).

## Matching noteikumi

Prioritāte: **1) relative path → 2) precīzs filename → 3) case-insensitive filename**.
Ja vienam nosaukumam atbilst vairāki faili dažādās mapēs, ieraksts kļūst
`AMBIGUOUS` — nejaušs fails **netiek** izvēlēts automātiski. Statusi:
`FOUND`, `MISSING`, `AMBIGUOUS`, `DUPLICATE`, `IGNORED` (nav .doc/.docx).

GUI pēc saraksta izvēles parāda priekšskatījuma tabulu
(`Order | Requested | Resolved file | Type | Status`) ar kopsavilkumu
**Atrasti / Trūkst / Dublikāti / Neskaidri** un `STRICT MODE` slēdzi.

## STRICT MODE

- **ON** — `MISSING`, `AMBIGUOUS`, `DUPLICATE` ieraksti **bloķē merge**
  (pirms jebkādas Word darbības; kļūda tiek reportēta).
- **OFF** (noklusēti) — problemātiskie ieraksti tiek izlaisti un ierakstīti
  `reports/result.json` (`order_list.summary`) un manifestā.

## Manuālā secība

`↑ Augšup`, `↓ Lejup`, `⤒ Sākumā`, `⤓ Beigās`, `Noņemt` (atbalsta arī vairāku
failu bloku) + klaviatūras tausiņi `Alt+↑` / `Alt+↓` / `Alt+Home` / `Alt+End`.
Īsts drag&drop nav ieviests apzināti — tas prasītu papildu atkarību
(`tkinterdnd2`), un to nepieļauj specifikācija.

## Eksportēt secību

`EKSPORTĒT SECĪBU` (GUI) vai `--export-order` (CLI) saglabā pašreizējo secību
CSV vai JSON ar `order,filename,path` — to var vēlāk ielādēt un atkārtot
identisku merge (pārbaudīts ar roundtrip testu).

## Manifest / checksum

`runtime/merge_manifest.json` satur katram avotam **relative path, size, mtime,
SHA-256**. Nākamajā palaišanā `runtime/manifest_diff.json` (un
`reports/result.json` → `manifest_diff`) parāda `UNCHANGED` / `CHANGED` /
`MISSING` / `ADDED`. Tas ir informatīvs un merge **nebloķē** (arī STRICT režīmā
bloķē tikai order saraksta kļūdas).

## Word sesijas drošība

DOC_MERGE_PRO nekad neaizver lietotāja jau atvērto Word sesiju:

- `DispatchEx` atgriež **jaunu** instanci → to drīkst `Quit()`;
- ja `DispatchEx` nokrīt un tiek izmantots `Dispatch` fallback, programma
  pārbauda, vai `WINWORD.EXE` jau darbojās; ja jā — instance tiek uzskatīta par
  **lietotāja** sesiju (`owned=False`) un `Quit()` **netiek** izsaukts;
- lietotāja sesijai netiek mainīts `Visible`, un `DisplayAlerts` tiek atjaunots.

## Offline testi

Visi testi strādā **bez interneta, API, cloud, GitHub un lejupielādēm**
(`docmerge/testing/guard.py` tehniski bloķē socket savienojumus un ieraksta
katru mēģinājumu reportā). Lokālais Microsoft Word COM ir atļauts.

```powershell
python -m docmerge.cli.main --offline-test        # QUICK (bez Word)
python -m docmerge.cli.main --offline-test-full   # FULL (Word E2E, apjomi, fault injection)
python _diag_soak.py                              # SOAK (120 DOCX, vairāki cikli)
python _diag_soak.py 300                          # lielāka slodze
python _diag_soak.py 40                           # ātrs smoke
```

Reporti:

| Fails | Satur |
| --- | --- |
| `reports/OFFLINE_QUICK_TEST_REPORT.md` / `..._RESULT.json` | QUICK rezultāti |
| `reports/OFFLINE_TEST_REPORT.md` / `OFFLINE_TEST_RESULT.json` | FULL rezultāti |
| `reports/SOAK_TEST_REPORT.md` / `SOAK_TEST_RESULT.json` | SOAK rezultāti |
| `reports/UNATTENDED_FINAL_REPORT.md` | autonomās sesijas gala atskaite |

QUICK aptver: imports, `compileall`, config, projekta roundtrip, skeneri,
natural/manual/ORDER_LIST kārtošanu, CSV un JSON parsēšanu (BOM, cp1257,
dublikāti, bojāti dati), matching, STRICT režīmu, dublikātus, trūkstošos failus,
neskaidros nosaukumus, absolūtos ceļus, unlock palīgus, manifestu un atskaites.

FULL papildus: 1 / 2 / 10 / 50 / 100 ģenerēti DOCX, mixed DOC + DOCX, latviešu
unicode nosaukumi (Ā Č Ē Ģ Ķ Ļ Ņ Š Ū Ž), atstarpes/iekavas/garie nosaukumi,
ligzdotas mapes, dublikāti dažādās mapēs, read-only, `Zone.Identifier`, saraksta
kļūdas, fault injection (bojāts ZIP, 0 B, `~$`, aizņemts izvades fails, pazudis
avots, bojāts JSON, dublikāti order numuri, normalizācijas kļūda, Word nav
pieejams, `InsertFile`/`SaveAs2` kļūdas) un reāli Word E2E testi, ieskaitot
**esošās Word sesijas** regresiju (`DispatchEx → Dispatch` fallback).

## Diagnostikas skripti

- `_diag_e2e.py` — dzīvs Word tests (2×.docx + 1×.doc + `.txt` + Zone.Identifier).
- `_diag_minimal.py` — pārbauda, ka Word pieņem minimālos OOXML DOCX.
- `_diag_soak.py` — soak tests.

## Arhitektūra

```
DOC_MERGE_PRO.py          # galvenais GUI launcher
app.py                     # compatibility launcher (tas pats GUI)
docmerge/
  cli/main.py              # CLI: --system-check, --merge-all, --order-list, --offline-test...
  gui/main_window.py       # Tkinter GUI (UiBridge nodrošina thread drošību)
  gui/order_list_dialog.py # CSV/JSON priekšskatījums (Order/Requested/Resolved/Status)
  core/scanner.py          # DOC/DOCX/DOCM/DOT/RTF/ODT/TXT skeneris
  core/sorter.py           # NATURAL / MANUAL / ORDER_LIST (absolūtā) secība
  core/order_list.py       # CSV/JSON parsēšana, matching, strict, eksports
  core/merge_service.py    # preflight + STRICT + manifest + Word dzinēja izsaukums
  core/unlock.py           # Zone.Identifier / Read-only (source immutable)
  core/paths.py            # absolūtie Word COM ceļi, garuma/cloud riska pārbaudes
  core/preflight.py        # ZIP/XML/OLE/RTF/TXT validācija
  engines/word_com.py      # Word COM: sesijas īpašumtiesības, InsertFile, normalizācija
  workers/                 # merge worker (thread), Word probe (subprocess)
  persistence/             # project JSON, merge manifest + diff
  reporting/               # logging, errors.jsonl, result.json
  testing/                 # offline testu ietvars (QUICK/FULL/SOAK), fakes, fixtures
BAT/                       # Windows launcheri
tests/                     # pytest (unit / offline / preflight)
```

Word COM atslēgas detaļas (pārbaudītas dzīvā Word testā):

- `Documents.Add()` tiek atkārtots (vienreizēja `WRD6ER32.CNV` kļūda), citādi
  master tiek atvērts no ģenerēta tukša DOCX;
- ja tiešais `InsertFile` nokrīt, avots tiek normalizēts uz `runtime/normalized`
  un mēģināts vēlreiz (virsraksts/lappuses pārtraukums netiek dublēts);
- visi Word'am padotie ceļi ir **absolūti**;
- normalizētie DOCX un atbloķētās kopijas pēc merge tiek izdzēstas
  (`cleanup_normalized=True`, var izslēgt konstruktorā);
- COM atsauces tiek atbrīvotas + `gc.collect()` + `CoFreeUnusedLibraries()`.

## Zināmie ierobežojumi

- Word (Windows) ir obligāts production merge; bez tā merge nav iespējams
  (QUICK offline testi tomēr darbojas pilnībā).
- Ceļu garums < 255 rakstzīmes (Word COM ierobežojums).
- Manifesta salīdzinājums ir globāls (`runtime/merge_manifest.json`) — tas
  attiecas uz pēdējo palaišanu, nevis uz katru projektu atsevišķi.
- Drag&drop secības maiņa nav ieviesta (tikai pogas/tausiņi).
- OneDrive "tikai tiešsaistē" vietturi tiek tikai brīdināti, nevis lejupielādēti
  automātiski (offline princips).
- `IGNORED` statusa ieraksti (nav .doc/.docx) nebloķē merge arī STRICT režīmā.
