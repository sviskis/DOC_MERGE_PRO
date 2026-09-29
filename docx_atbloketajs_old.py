from __future__ import annotations

import json
import os
import queue
import shutil
import stat
import subprocess
import tempfile
import threading
import tkinter as tk
import unicodedata
from collections import defaultdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "DOCX atbloķētājs"
DOCX_SUFFIX = ".docx"
MERGED_FOLDER_NAME = "_APVIENOTI_PEC_BURTA"
LATVIAN_ALPHABET = "AĀBCČDEĒFGĢHIĪJKĶLĻMNŅOPRSŠTUŪVZŽ"
LATVIAN_ORDER = {letter: index for index, letter in enumerate(LATVIAN_ALPHABET)}

WORD_MERGE_POWERSHELL = r"""
param(
    [Parameter(Mandatory = $true)]
    [string]$TaskFile
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$word = $null

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        } catch {
        }
    }
}

try {
    $data = Get-Content -LiteralPath $TaskFile -Raw | ConvertFrom-Json
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($group in $data.groups) {
        $document = $null
        $range = $null
        try {
            $outputPath = [string]$group.output
            if (Test-Path -LiteralPath $outputPath) {
                Remove-Item -LiteralPath $outputPath -Force
            }

            $document = $word.Documents.Add()
            $isFirst = $true

            foreach ($inputItem in $group.files) {
                if ($isFirst) {
                    $range = $document.Range(0, 0)
                    $isFirst = $false
                } else {
                    $position = $document.Content.End - 1
                    $range = $document.Range($position, $position)
                    if ([bool]$group.page_break) {
                        $range.InsertBreak(7)
                        $range.Collapse(0)
                    }
                }

                if ([bool]$group.add_title) {
                    # Faila nosaukums bez .docx kļūst par katra stāsta
                    # atsevišķu treknu virsraksta rindkopu.
                    $range.Text = ([string]$inputItem.title + "`r")
                    $range.Style = -1
                    $range.Font.Bold = -1
                    $range.Font.Italic = 0
                    $range.ParagraphFormat.SpaceBefore = 0
                    $range.ParagraphFormat.SpaceAfter = 6
                    $range.ParagraphFormat.KeepWithNext = -1
                    $range.Collapse(0)
                    $range.Font.Bold = 0
                    $range.Font.Italic = 0
                    $range.ParagraphFormat.KeepWithNext = 0
                }

                $range.InsertFile([string]$inputItem.path)
                Release-ComObject $range
                $range = $null
            }

            # 16 = wdFormatDocumentDefault (.docx)
            $document.SaveAs2($outputPath, 16)
            $document.Close(0)
            Release-ComObject $document
            $document = $null
            Write-Output ("OK`t" + [string]$group.label + "`t" + $outputPath)
        } catch {
            $errorText = $_.Exception.Message -replace "[`r`n`t]", " "
            if ($null -ne $document) {
                try {
                    $document.Close(0)
                } catch {
                }
            }
            Release-ComObject $range
            Release-ComObject $document
            $range = $null
            $document = $null
            Write-Output (
                "ERROR`t" + [string]$group.label + "`t" + $errorText
            )
        }
    }
} catch {
    $errorText = $_.Exception.Message -replace "[`r`n`t]", " "
    Write-Output ("FATAL`t`t" + $errorText)
    exit 1
} finally {
    if ($null -ne $word) {
        try {
            $word.Quit()
        } catch {
        }
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def remove_zone_identifier(file_path: Path) -> bool:
    """Noņem Windows Mark-of-the-Web alternatīvo datu plūsmu.

    Atgriež True, ja Zone.Identifier tika atrasts un noņemts.
    """
    zone_path = f"{file_path}:Zone.Identifier"
    try:
        os.remove(zone_path)
        return True
    except FileNotFoundError:
        return False
    except OSError as exc:
        # Windows dažkārt neesošai ADS plūsmai atgriež winerror 2 vai 3.
        if getattr(exc, "winerror", None) in (2, 3):
            return False
        raise


def make_writable(file_path: Path) -> bool:
    """Noņem faila Read-only atribūtu. Atgriež True, ja tas bija uzlikts."""
    current_mode = file_path.stat().st_mode
    was_read_only = not bool(current_mode & stat.S_IWRITE)
    if was_read_only:
        os.chmod(file_path, current_mode | stat.S_IWRITE)
    return was_read_only


def latvian_text_sort_key(text: str) -> tuple:
    normalized = unicodedata.normalize("NFC", text).upper()
    key = []
    for character in normalized:
        if character in LATVIAN_ORDER:
            key.append((0, LATVIAN_ORDER[character]))
        elif character.isdigit():
            key.append((1, int(character)))
        else:
            key.append((2, ord(character)))
    return tuple(key)


def initial_letter(file_path: Path) -> str:
    normalized = unicodedata.normalize("NFC", file_path.stem).upper()
    for character in normalized:
        if character.isalpha():
            return character
    return "CITI"


def group_files_by_initial(files: list[Path]) -> dict[str, list[Path]]:
    grouped: dict[str, list[Path]] = defaultdict(list)
    for file_path in files:
        grouped[initial_letter(file_path)].append(file_path)

    ordered_groups: dict[str, list[Path]] = {}
    labels = sorted(
        grouped,
        key=lambda label: (
            0,
            LATVIAN_ORDER[label],
        )
        if label in LATVIAN_ORDER
        else (1, latvian_text_sort_key(label)),
    )
    for label in labels:
        ordered_groups[label] = sorted(
            grouped[label],
            key=lambda path: latvian_text_sort_key(path.name),
        )
    return ordered_groups


def collect_docx_files(source: Path, recursive: bool) -> list[Path]:
    iterator = source.rglob("*") if recursive else source.iterdir()
    return sorted(
        (
            path
            for path in iterator
            if path.is_file()
            and path.suffix.lower() == DOCX_SUFFIX
            and not path.name.startswith("~$")
        ),
        key=lambda path: latvian_text_sort_key(
            str(path.relative_to(source))
        ),
    )


def destination_is_inside_source(source: Path, destination: Path) -> bool:
    source_text = os.path.normcase(os.path.abspath(source))
    destination_text = os.path.normcase(os.path.abspath(destination))
    try:
        return os.path.commonpath([source_text, destination_text]) == source_text
    except ValueError:
        # Dažādi Windows diski, piemēram, C: un D:.
        return False


def copy_and_unlock(
    source_file: Path,
    source_root: Path,
    destination_root: Path,
    overwrite: bool,
) -> tuple[str, Path, bool, bool]:
    relative_path = source_file.relative_to(source_root)
    destination_file = destination_root / relative_path

    if destination_file.exists() and not overwrite:
        zone_removed = remove_zone_identifier(destination_file)
        read_only_removed = make_writable(destination_file)
        return "skipped", destination_file, zone_removed, read_only_removed

    destination_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_file, destination_file)

    zone_removed = remove_zone_identifier(destination_file)
    read_only_removed = make_writable(destination_file)
    return "copied", destination_file, zone_removed, read_only_removed


def merge_groups_with_word(
    groups: dict[str, list[Path]],
    destination_root: Path,
    page_break: bool,
    add_titles: bool,
    events: queue.Queue[tuple],
    cancel_event: threading.Event,
) -> tuple[int, int]:
    merged_folder = destination_root / MERGED_FOLDER_NAME
    merged_folder.mkdir(parents=True, exist_ok=True)

    tasks = {
        "groups": [
            {
                "label": label,
                "output": str(merged_folder / f"{label}_apvienots.docx"),
                "files": [
                    {
                        "path": str(path),
                        "title": path.stem,
                    }
                    for path in files
                ],
                "page_break": page_break,
                "add_title": add_titles,
            }
            for label, files in groups.items()
        ]
    }

    if not tasks["groups"]:
        return 0, 0

    with tempfile.TemporaryDirectory(prefix="docx_apvienosana_") as temp_dir_text:
        temp_dir = Path(temp_dir_text)
        task_file = temp_dir / "tasks.json"
        script_file = temp_dir / "merge_docx.ps1"
        task_file.write_text(
            json.dumps(tasks, ensure_ascii=False, indent=2),
            encoding="utf-8-sig",
        )
        script_file.write_text(WORD_MERGE_POWERSHELL, encoding="utf-8-sig")

        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        process = subprocess.Popen(
            [
                "powershell.exe",
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_file),
                "-TaskFile",
                str(task_file),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creation_flags,
        )

        merged = 0
        failed = 0
        assert process.stdout is not None

        for output_line in process.stdout:
            line = output_line.strip()
            if not line:
                continue

            parts = line.split("\t", 2)
            status = parts[0]
            label = parts[1] if len(parts) > 1 else "?"
            detail = parts[2] if len(parts) > 2 else line

            if status == "OK":
                merged += 1
                try:
                    merged_file = Path(detail)
                    remove_zone_identifier(merged_file)
                    make_writable(merged_file)
                except Exception:
                    pass
                events.put(
                    (
                        "log",
                        f"APVIENOTS [{label}]: {detail}",
                    )
                )
                events.put(("merge_progress", merged + failed, len(groups)))
            elif status in ("ERROR", "FATAL"):
                failed += 1
                events.put(
                    (
                        "log",
                        f"APVIENOŠANAS KĻŪDA [{label}]: {detail}",
                    )
                )
                events.put(("merge_progress", merged + failed, len(groups)))
            else:
                events.put(("log", f"Word: {line}"))

        return_code = process.wait()
        if merged + failed < len(groups):
            failed += len(groups) - merged - failed
        if return_code != 0:
            events.put(
                (
                    "log",
                    f"Microsoft Word apvienošanas process beidzās ar kodu "
                    f"{return_code}.",
                )
            )

    return merged, failed


class DocxUnlockerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("820x650")
        self.minsize(720, 560)

        self.source_var = tk.StringVar()
        self.destination_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.merge_var = tk.BooleanVar(value=True)
        self.page_break_var = tk.BooleanVar(value=True)
        self.add_titles_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Izvēlieties avota un rezultāta mapi.")

        self.events: queue.Queue[tuple] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None

        self._configure_style()
        self._build_ui()
        self.after(100, self._poll_events)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if os.name == "nt":
            try:
                style.theme_use("vista")
            except tk.TclError:
                pass
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Info.TLabel", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(6, weight=1)

        ttk.Label(
            outer,
            text="DOCX Protected View noņēmējs",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            outer,
            text=(
                "Kopē DOCX failus uz jaunu mapi un noņem Windows interneta "
                "drošības marķējumu. Pēc tam apvieno stāstus pēc burta un "
                "pievieno virsrakstus no failu nosaukumiem. Oriģinālie faili "
                "netiek mainīti."
            ),
            style="Info.TLabel",
            wraplength=760,
        ).grid(row=1, column=0, sticky="ew", pady=(4, 16))

        paths = ttk.LabelFrame(outer, text="Mapes", padding=12)
        paths.grid(row=2, column=0, sticky="ew")
        paths.columnconfigure(1, weight=1)

        ttk.Label(paths, text="No kurienes:").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=5
        )
        ttk.Entry(paths, textvariable=self.source_var).grid(
            row=0, column=1, sticky="ew", pady=5
        )
        ttk.Button(paths, text="Izvēlēties…", command=self._choose_source).grid(
            row=0, column=2, padx=(8, 0), pady=5
        )

        ttk.Label(paths, text="Kur nolikt:").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=5
        )
        ttk.Entry(paths, textvariable=self.destination_var).grid(
            row=1, column=1, sticky="ew", pady=5
        )
        ttk.Button(
            paths,
            text="Izvēlēties…",
            command=self._choose_destination,
        ).grid(row=1, column=2, padx=(8, 0), pady=5)

        options = ttk.Frame(outer)
        options.grid(row=3, column=0, sticky="ew", pady=(12, 8))
        ttk.Checkbutton(
            options,
            text="Iekļaut apakšmapes un saglabāt mapju struktūru",
            variable=self.recursive_var,
        ).pack(anchor="w")
        ttk.Checkbutton(
            options,
            text="Pārrakstīt rezultāta mapē jau esošos DOCX",
            variable=self.overwrite_var,
        ).pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(
            options,
            text=(
                "Apvienot pēc pirmā burta latviešu alfabēta kārtībā "
                f"({MERGED_FOLDER_NAME})"
            ),
            variable=self.merge_var,
        ).pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(
            options,
            text="Apvienotajā DOCX katru dokumentu sākt jaunā lapā",
            variable=self.page_break_var,
        ).pack(anchor="w", pady=(4, 0), padx=(22, 0))
        ttk.Checkbutton(
            options,
            text=(
                "Pirms katra stāsta pievienot faila nosaukumu kā "
                "treknu virsrakstu"
            ),
            variable=self.add_titles_var,
        ).pack(anchor="w", pady=(4, 0), padx=(22, 0))

        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, sticky="ew", pady=(4, 8))
        self.scan_button = ttk.Button(
            actions,
            text="Saskaitīt DOCX",
            command=self._start_scan,
        )
        self.scan_button.pack(side="left")
        self.run_button = ttk.Button(
            actions,
            text="ATBLOĶĒT, KOPĒT UN APVIENOT",
            style="Accent.TButton",
            command=self._start_processing,
        )
        self.run_button.pack(side="left", padx=8)
        self.cancel_button = ttk.Button(
            actions,
            text="Atcelt",
            command=self._cancel,
            state="disabled",
        )
        self.cancel_button.pack(side="left")

        progress_frame = ttk.Frame(outer)
        progress_frame.grid(row=5, column=0, sticky="ew", pady=(2, 10))
        progress_frame.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
        )
        self.progress.grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.status_var).grid(
            row=1, column=0, sticky="w", pady=(5, 0)
        )

        log_frame = ttk.LabelFrame(outer, text="Darba žurnāls", padding=8)
        log_frame.grid(row=6, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(
            log_frame,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
            relief="flat",
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        ttk.Label(
            outer,
            text=(
                "Piezīme: programma noņem Protected View interneta marķējumu; "
                "tā nenoņem DOCX paroles vai dokumenta rediģēšanas aizsardzību."
            ),
            foreground="#555555",
            wraplength=760,
        ).grid(row=7, column=0, sticky="ew", pady=(10, 0))

    def _choose_source(self) -> None:
        selected = filedialog.askdirectory(title="Izvēlieties mapi ar DOCX failiem")
        if selected:
            self.source_var.set(selected)

    def _choose_destination(self) -> None:
        selected = filedialog.askdirectory(
            title="Izvēlieties mapi atbloķētajām DOCX kopijām"
        )
        if selected:
            self.destination_var.set(selected)

    def _validated_paths(self, require_destination: bool) -> tuple[Path, Path | None]:
        source_text = self.source_var.get().strip().strip('"')
        destination_text = self.destination_var.get().strip().strip('"')

        if not source_text:
            raise ValueError("Nav izvēlēta avota mape.")

        source = Path(source_text)
        if not source.is_dir():
            raise ValueError("Avota mape neeksistē vai nav pieejama.")

        if not require_destination:
            return source, None

        if not destination_text:
            raise ValueError("Nav izvēlēta rezultāta mape.")

        destination = Path(destination_text)
        if source.resolve() == destination.resolve():
            raise ValueError("Avota un rezultāta mape nedrīkst būt viena un tā pati.")
        if destination_is_inside_source(source, destination):
            raise ValueError(
                "Rezultāta mape nedrīkst atrasties avota mapē. "
                "Izvēlieties atsevišķu mapi."
            )
        return source, destination

    def _set_busy(self, busy: bool) -> None:
        normal_or_disabled = "disabled" if busy else "normal"
        self.scan_button.configure(state=normal_or_disabled)
        self.run_button.configure(state=normal_or_disabled)
        self.cancel_button.configure(state="normal" if busy else "disabled")

    def _start_scan(self) -> None:
        try:
            source, _ = self._validated_paths(require_destination=False)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return

        self._clear_log()
        self.progress.configure(value=0, maximum=1)
        self.status_var.set("Meklē DOCX failus…")
        self._set_busy(True)
        self.cancel_event.clear()
        self.worker = threading.Thread(
            target=self._scan_worker,
            args=(source, self.recursive_var.get()),
            daemon=True,
        )
        self.worker.start()

    def _scan_worker(self, source: Path, recursive: bool) -> None:
        try:
            files = collect_docx_files(source, recursive)
            self.events.put(("scan_done", len(files)))
        except Exception as exc:
            self.events.put(("fatal", f"Neizdevās nolasīt avota mapi:\n{exc}"))

    def _start_processing(self) -> None:
        if os.name != "nt":
            messagebox.showerror(
                APP_TITLE,
                "Šī programma paredzēta Windows 10/11.",
                parent=self,
            )
            return

        try:
            source, destination = self._validated_paths(require_destination=True)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return

        assert destination is not None
        self._clear_log()
        self.progress.configure(value=0, maximum=1)
        self.status_var.set("Meklē DOCX failus…")
        self._set_busy(True)
        self.cancel_event.clear()

        self.worker = threading.Thread(
            target=self._processing_worker,
            args=(
                source,
                destination,
                self.recursive_var.get(),
                self.overwrite_var.get(),
                self.merge_var.get(),
                self.page_break_var.get(),
                self.add_titles_var.get(),
            ),
            daemon=True,
        )
        self.worker.start()

    def _processing_worker(
        self,
        source: Path,
        destination: Path,
        recursive: bool,
        overwrite: bool,
        merge_after_copy: bool,
        page_break: bool,
        add_titles: bool,
    ) -> None:
        copied = 0
        skipped = 0
        failed = 0
        zone_removed = 0
        read_only_removed = 0
        merged = 0
        merge_failed = 0
        merge_inputs: list[Path] = []

        try:
            files = collect_docx_files(source, recursive)
            self.events.put(("total", len(files)))

            if not files:
                self.events.put(
                    (
                        "done",
                        copied,
                        skipped,
                        failed,
                        0,
                        0,
                        merged,
                        merge_failed,
                        destination,
                    )
                )
                return

            destination.mkdir(parents=True, exist_ok=True)

            for index, source_file in enumerate(files, start=1):
                if self.cancel_event.is_set():
                    self.events.put(
                        (
                            "cancelled",
                            copied,
                            skipped,
                            failed,
                            zone_removed,
                            read_only_removed,
                            merged,
                            merge_failed,
                            destination,
                        )
                    )
                    return

                try:
                    result, destination_file, removed_zone, removed_read_only = (
                        copy_and_unlock(
                            source_file,
                            source,
                            destination,
                            overwrite,
                        )
                    )
                    zone_removed += int(removed_zone)
                    read_only_removed += int(removed_read_only)
                    if result == "skipped":
                        skipped += 1
                        self.events.put(
                            ("log", f"IZLAISTS (jau eksistē): {destination_file}")
                        )
                    else:
                        copied += 1
                        self.events.put(("log", f"GATAVS: {destination_file}"))
                    merge_inputs.append(destination_file)
                except Exception as exc:
                    failed += 1
                    self.events.put(("log", f"KĻŪDA: {source_file}\n  {exc}"))

                self.events.put(("progress", index, len(files)))

            if merge_after_copy and merge_inputs and not self.cancel_event.is_set():
                groups = group_files_by_initial(merge_inputs)
                self.events.put(("merge_start", len(groups)))
                try:
                    merged, merge_failed = merge_groups_with_word(
                        groups,
                        destination,
                        page_break,
                        add_titles,
                        self.events,
                        self.cancel_event,
                    )
                except FileNotFoundError:
                    merge_failed = len(groups)
                    self.events.put(
                        (
                            "log",
                            "APVIENOŠANAS KĻŪDA: Windows PowerShell nav atrasts.",
                        )
                    )
                except Exception as exc:
                    merge_failed = len(groups)
                    self.events.put(
                        (
                            "log",
                            "APVIENOŠANAS KĻŪDA: "
                            f"Microsoft Word nevarēja apvienot dokumentus.\n  {exc}",
                        )
                    )

            self.events.put(
                (
                    "done",
                    copied,
                    skipped,
                    failed,
                    zone_removed,
                    read_only_removed,
                    merged,
                    merge_failed,
                    destination,
                )
            )
        except Exception as exc:
            self.events.put(("fatal", f"Darbu neizdevās pabeigt:\n{exc}"))

    def _cancel(self) -> None:
        self.cancel_event.set()
        self.status_var.set("Atcelšana…")
        self.cancel_button.configure(state="disabled")

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                event_type = event[0]

                if event_type == "log":
                    self._append_log(event[1])
                elif event_type == "total":
                    total = event[1]
                    self.progress.configure(value=0, maximum=max(total, 1))
                    self.status_var.set(f"Atrasti {total} DOCX. Sāk apstrādi…")
                elif event_type == "progress":
                    current, total = event[1], event[2]
                    self.progress.configure(value=current, maximum=max(total, 1))
                    self.status_var.set(f"Apstrādā {current} no {total}…")
                elif event_type == "merge_start":
                    group_count = event[1]
                    # Word COM procesa piespiedu apturēšana var atstāt WINWORD.EXE
                    # fonā, tāpēc drošības dēļ apvienošanas laikā pogu atspējo.
                    self.cancel_button.configure(state="disabled")
                    self.progress.configure(value=0, maximum=max(group_count, 1))
                    self.status_var.set(
                        f"Microsoft Word apvieno {group_count} burtu grupas…"
                    )
                    self._append_log("")
                    self._append_log(
                        f"Sāk apvienošanu: {group_count} burtu grupas."
                    )
                elif event_type == "merge_progress":
                    current, total = event[1], event[2]
                    self.progress.configure(value=current, maximum=max(total, 1))
                    self.status_var.set(
                        f"Apvieno burtu grupas: {current} no {total}…"
                    )
                elif event_type == "scan_done":
                    count = event[1]
                    self.status_var.set(f"Atrasti {count} DOCX faili.")
                    self._append_log(f"Atrasti {count} DOCX faili.")
                    self._set_busy(False)
                elif event_type in ("done", "cancelled"):
                    self._finish(event_type, *event[1:])
                elif event_type == "fatal":
                    self._set_busy(False)
                    self.status_var.set("Radās kļūda.")
                    self._append_log(event[1])
                    messagebox.showerror(APP_TITLE, event[1], parent=self)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_events)

    def _finish(
        self,
        event_type: str,
        copied: int,
        skipped: int,
        failed: int,
        zone_removed: int,
        read_only_removed: int,
        merged: int,
        merge_failed: int,
        destination: Path,
    ) -> None:
        self._set_busy(False)
        cancelled = event_type == "cancelled"
        heading = "Darbs atcelts." if cancelled else "Darbs pabeigts."
        summary = (
            f"{heading}\n\n"
            f"Nokopēti un sagatavoti: {copied}\n"
            f"Izlaisti: {skipped}\n"
            f"Kļūdas: {failed}\n"
            f"Noņemts Protected View marķējums: {zone_removed}\n"
            f"Noņemts Read-only atribūts: {read_only_removed}\n\n"
            f"Izveidoti apvienotie DOCX: {merged}\n"
            f"Apvienošanas kļūdas: {merge_failed}\n\n"
            f"Rezultāta mape:\n{destination}"
        )
        self.status_var.set(
            f"{heading} Gatavi: {copied}; apvienoti: {merged}; "
            f"kļūdas: {failed + merge_failed}."
        )
        self._append_log("")
        self._append_log(summary)

        if failed or merge_failed:
            messagebox.showwarning(APP_TITLE, summary, parent=self)
        else:
            messagebox.showinfo(APP_TITLE, summary, parent=self)

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"{text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def enable_windows_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def main() -> None:
    enable_windows_dpi_awareness()
    app = DocxUnlockerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
