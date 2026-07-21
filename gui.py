"""Single-window GUI for architect-template specification formatting."""

from __future__ import annotations

import os
import queue
import threading
import traceback
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Tuple

import customtkinter as ctk
from tkinter import filedialog, messagebox

from spec_formatter.pipeline import (
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    FormatRunResult,
    collect_target_specs,
    default_template_cache_dir,
    format_specifications,
)


COLORS = {
    "bg": "#0D0D0D",
    "card": "#191919",
    "input": "#252525",
    "border": "#353535",
    "text": "#FFFFFF",
    "secondary": "#B0B0B0",
    "muted": "#737373",
    "accent": "#3B82F6",
    "accent_hover": "#2563EB",
    "success": "#22C55E",
    "warning": "#F59E0B",
    "error": "#EF4444",
}

UI_FONT = "Segoe UI"
MONO_FONT = "Consolas"


def _font(size: int, weight: str = "normal", family: str = UI_FONT) -> ctk.CTkFont:
    return ctk.CTkFont(family=family, size=size, weight=weight)


def _load_prompt_file(path: Path) -> str:
    """Compatibility helper retained for the template-engine contract tests."""

    if not path.exists():
        raise FileNotFoundError(f"Missing required prompt file: {path}")
    try:
        return path.read_text(encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Failed reading prompt file {path}: {exc}") from exc


def discover_target_docx(folder: Path) -> list[Path]:
    """Compatibility helper used by tests and folder-preview code."""

    return list(collect_target_specs([Path(folder)]))


def summarize_batch_results(results: Iterable[object]) -> Tuple[str, str]:
    """Return a compact status and message for per-target result objects."""

    items = list(results)
    if not items:
        return "empty", "No target specs found"
    succeeded = sum(1 for item in items if bool(getattr(item, "success", False)))
    total = len(items)
    if succeeded == total:
        return "success", f"Complete: {succeeded}/{total} formatted"
    if succeeded == 0:
        return "failed", f"Failed: 0/{total} formatted"
    return "partial", f"Partial: {succeeded}/{total} formatted"


def conversion_report_log_lines(item: object) -> tuple[str, ...]:
    """Render conversion diagnostics for successful or failed target results."""

    report = getattr(item, "conversion_report", None)
    if report is None:
        return ()
    source_path = Path(getattr(item, "source_path", "target.docx"))
    lines = [
        f"{source_path.name}: Canadian conversion processed "
        f"{report.paragraphs_converted} numbered paragraphs and removed "
        f"{report.literal_markers_removed} typed markers."
    ]
    lines.extend(
        f"{source_path.name} warning p[{issue.paragraph_index}]: {issue.message}"
        for issue in report.warnings
    )
    return tuple(lines)


@dataclass(frozen=True)
class ActiveRunSummary:
    """Values captured for a run, excluding the API key by design."""

    architect_template: Path
    target_inputs: tuple[Path, ...]
    output_root: Path
    conversion_mode: str
    reuse_template_analysis: bool
    max_workers: int


def output_mode_label(conversion_mode: str) -> str:
    """Return the user-facing label for a pipeline conversion mode."""

    if conversion_mode == CSI_TO_CANADIAN:
        return "Canadian CSC PageFormat conversion"
    return "Format only"


def active_run_summary_text(summary: ActiveRunSummary, *, active: bool = True) -> str:
    """Render a stable, non-secret snapshot of the values used by the worker."""

    heading = "ACTIVE RUN (inputs locked)" if active else "LAST RUN"
    analysis = (
        "reuse matching profile"
        if summary.reuse_template_analysis
        else "reanalyze template"
    )
    return (
        f"{heading}\n"
        f"{output_mode_label(summary.conversion_mode)} | "
        f"{len(summary.target_inputs)} target selection(s) | "
        f"{summary.max_workers} worker(s) | {analysis}\n"
        f"Template: {summary.architect_template}\n"
        f"Output root: {summary.output_root}"
    )


def result_run_directory(result: object) -> Optional[Path]:
    """Return the concrete run directory, with backward-compatible fallbacks."""

    for attribute in ("run_dir", "output_dir", "output_root"):
        value = getattr(result, attribute, None)
        if value:
            return Path(value)
    return None


def target_result_log_lines(item: object) -> tuple[str, ...]:
    """Return every processor log line plus any available audit diagnostics."""

    lines: list[str] = []
    raw_log = getattr(item, "log", ()) or ()
    if isinstance(raw_log, str):
        raw_log = (raw_log,)
    for entry in raw_log:
        text = str(entry)
        split = text.splitlines()
        lines.extend(split if split else ("",))

    source_path = Path(getattr(item, "source_path", "target.docx"))
    audit_summary = getattr(item, "audit_summary", None)
    if isinstance(audit_summary, Mapping) and audit_summary:
        preferred_keys = ("styled", "ignored", "out_of_scope", "unresolved")
        ordered_keys = [key for key in preferred_keys if key in audit_summary]
        ordered_keys.extend(
            sorted(
                (key for key in audit_summary if key not in preferred_keys),
                key=str,
            )
        )
        counts = ", ".join(f"{key}={audit_summary[key]}" for key in ordered_keys)
        lines.append(f"{source_path.name}: audit counts: {counts}")

    audit_path = getattr(item, "audit_path", None)
    if audit_path:
        lines.append(f"{source_path.name}: audit: {audit_path}")
    lines.extend(conversion_report_log_lines(item))
    return tuple(lines)


class FormatWorker(threading.Thread):
    """Run the unified pipeline without blocking Tk's event loop."""

    def __init__(
        self,
        architect_template: Path,
        target_inputs: tuple[Path, ...],
        output_dir: Path,
        api_key: str,
        reuse_template_analysis: bool,
        max_workers: int,
        conversion_mode: str,
        events: queue.Queue,
    ) -> None:
        super().__init__(daemon=False)
        self.architect_template = architect_template
        self.target_inputs = target_inputs
        self.output_dir = output_dir
        self.api_key = api_key
        self.reuse_template_analysis = reuse_template_analysis
        self.max_workers = max_workers
        self.conversion_mode = conversion_mode
        self.events = events

    def _progress(self, message: str) -> None:
        self.events.put(("progress", message))

    def run(self) -> None:
        try:
            result = format_specifications(
                architect_template=self.architect_template,
                target_specs=self.target_inputs,
                output_dir=self.output_dir,
                api_key=self.api_key,
                cache_dir=default_template_cache_dir(),
                force_template_analysis=not self.reuse_template_analysis,
                max_workers=self.max_workers,
                conversion_mode=self.conversion_mode,
                progress=self._progress,
            )
            self.events.put(("complete", result))
        except Exception as exc:
            self.events.put(
                (
                    "error",
                    {
                        "message": str(exc),
                        "traceback": traceback.format_exc(),
                        "run_dir": getattr(exc, "run_dir", None),
                        "manifest_path": getattr(exc, "manifest_path", None),
                    },
                )
            )


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Specification Formatter")
        self.geometry("980x930")
        self.minsize(820, 720)
        self.configure(fg_color=COLORS["bg"])

        self.architect_var = ctk.StringVar()
        self.output_var = ctk.StringVar()
        self.api_key_var = ctk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.show_key_var = ctk.BooleanVar(value=False)
        self.reuse_var = ctk.BooleanVar(value=True)
        self.workers_var = ctk.StringVar(value="3")
        self.conversion_mode_var = ctk.StringVar(value=FORMAT_ONLY)
        self.mode_controls: list[ctk.CTkRadioButton] = []
        self.run_affecting_controls: list[object] = []
        self._locked_run_control_states: list[tuple[object, str]] = []
        self.target_inputs: list[Path] = []
        self.output_is_automatic = False
        self.events: queue.Queue = queue.Queue()
        self.worker: Optional[FormatWorker] = None
        self.last_result: Optional[FormatRunResult] = None
        self.failed_run_dir: Optional[Path] = None
        self.active_output_dir: Optional[Path] = None
        self.active_run_summary: Optional[ActiveRunSummary] = None
        self.advanced_visible = False

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_events)

    def _build_ui(self) -> None:
        shell = ctk.CTkFrame(self, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=28, pady=24)

        header = ctk.CTkFrame(shell, fg_color="transparent")
        header.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(
            header,
            text="SPECIFICATION FORMATTER",
            text_color=COLORS["text"],
            font=_font(30, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Apply an architect's Word template to one or more target specifications.",
            text_color=COLORS["secondary"],
            font=_font(15),
        ).pack(anchor="w", pady=(5, 0))

        card = ctk.CTkFrame(
            shell,
            fg_color=COLORS["card"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=10,
        )
        card.pack(fill="x")

        self._section_label(card, "1   Architect template")
        self.architect_entry, self.architect_button = self._path_row(
            card,
            self.architect_var,
            "Select the architect's .docx template",
            self._choose_architect,
            "Choose File",
        )

        self._section_label(card, "2   Target specifications", top=18)
        target_toolbar = ctk.CTkFrame(card, fg_color="transparent")
        target_toolbar.pack(fill="x", padx=22)
        self.add_files_button = ctk.CTkButton(
            target_toolbar,
            text="Add Files",
            command=self._add_target_files,
            width=112,
            height=36,
            font=_font(14),
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
        )
        self.add_files_button.pack(side="left")
        self.add_folder_button = ctk.CTkButton(
            target_toolbar,
            text="Add Folder",
            command=self._add_target_folder,
            width=112,
            height=36,
            font=_font(14),
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
        )
        self.add_folder_button.pack(side="left", padx=(8, 0))
        self.clear_targets_button = ctk.CTkButton(
            target_toolbar,
            text="Clear",
            command=self._clear_targets,
            width=76,
            height=36,
            font=_font(14),
            fg_color="transparent",
            hover_color=COLORS["border"],
            text_color=COLORS["secondary"],
        )
        self.clear_targets_button.pack(side="right")
        self.run_affecting_controls.extend(
            (
                self.add_files_button,
                self.add_folder_button,
                self.clear_targets_button,
            )
        )

        self.target_box = ctk.CTkTextbox(
            card,
            height=86,
            fg_color=COLORS["input"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["secondary"],
            font=_font(13),
            activate_scrollbars=True,
        )
        self.target_box.pack(fill="x", padx=22, pady=(9, 0))
        self.target_box.configure(state="disabled")
        self._refresh_target_preview()

        self._section_label(card, "3   Output mode", top=18)
        mode_row = ctk.CTkFrame(card, fg_color="transparent")
        mode_row.pack(fill="x", padx=22)
        for label, value in (
            ("Format only", FORMAT_ONLY),
            ("Convert CSI hierarchy to Canadian CSC PageFormat", CSI_TO_CANADIAN),
        ):
            control = ctk.CTkRadioButton(
                mode_row,
                text=label,
                value=value,
                variable=self.conversion_mode_var,
                text_color=COLORS["secondary"],
                font=_font(13),
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
            )
            control.pack(side="left", padx=(0, 28))
            self.mode_controls.append(control)
            self.run_affecting_controls.append(control)
        ctk.CTkLabel(
            card,
            text=(
                "Canadian mode converts recognized CSI numbering and hierarchy before "
                "formatting. The architect template must use automatic Canadian 1.1/.1 "
                "numbering; each target article needs a preceding Part, and the architect "
                "must use one coherent automatic Part/list hierarchy. It does not revise "
                "codes, standards, units, terminology, or technical requirements."
            ),
            wraplength=870,
            justify="left",
            text_color=COLORS["muted"],
            font=_font(12),
        ).pack(anchor="w", padx=22, pady=(7, 0))

        self._section_label(card, "4   Output folder", top=18)
        self.output_entry, self.output_button = self._path_row(
            card,
            self.output_var,
            "Formatted Specs",
            self._choose_output,
            "Choose Folder",
        )

        self._section_label(card, "5   Anthropic API key", top=18)
        key_row = ctk.CTkFrame(card, fg_color="transparent")
        key_row.pack(fill="x", padx=22)
        self.api_entry = ctk.CTkEntry(
            key_row,
            textvariable=self.api_key_var,
            placeholder_text="Required when template or target analysis needs AI",
            show="•",
            height=40,
            fg_color=COLORS["input"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            font=_font(14),
        )
        self.api_entry.pack(side="left", fill="x", expand=True)
        self.run_affecting_controls.append(self.api_entry)
        self.show_key_checkbox = ctk.CTkCheckBox(
            key_row,
            text="Show",
            variable=self.show_key_var,
            command=self._toggle_key,
            width=72,
            text_color=COLORS["secondary"],
            font=_font(13),
            fg_color=COLORS["accent"],
        )
        self.show_key_checkbox.pack(side="left", padx=(12, 0))

        advanced_button = ctk.CTkButton(
            card,
            text="Advanced settings  ▸",
            command=self._toggle_advanced,
            width=170,
            height=30,
            fg_color="transparent",
            hover_color=COLORS["input"],
            text_color=COLORS["muted"],
            font=_font(13),
        )
        advanced_button.pack(anchor="w", padx=16, pady=(14, 0))
        self.advanced_button = advanced_button

        self.advanced_frame = ctk.CTkFrame(card, fg_color=COLORS["input"])
        self.reuse_checkbox = ctk.CTkCheckBox(
            self.advanced_frame,
            text="Reuse analysis when the architect template has not changed",
            variable=self.reuse_var,
            text_color=COLORS["secondary"],
            font=_font(13),
            fg_color=COLORS["accent"],
        )
        self.reuse_checkbox.pack(side="left", padx=14, pady=12)
        self.run_affecting_controls.append(self.reuse_checkbox)
        workers = ctk.CTkFrame(self.advanced_frame, fg_color="transparent")
        workers.pack(side="right", padx=14, pady=8)
        ctk.CTkLabel(
            workers,
            text="Concurrent files",
            text_color=COLORS["secondary"],
            font=_font(13),
        ).pack(side="left", padx=(0, 8))
        self.workers_menu = ctk.CTkOptionMenu(
            workers,
            values=["1", "2", "3", "4", "5", "6"],
            variable=self.workers_var,
            width=66,
            height=30,
            fg_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            font=_font(13),
        )
        self.workers_menu.pack(side="left")
        self.run_affecting_controls.append(self.workers_menu)

        action_row = ctk.CTkFrame(card, fg_color="transparent")
        self.action_row = action_row
        action_row.pack(fill="x", padx=22, pady=20)
        self.run_button = ctk.CTkButton(
            action_row,
            text="FORMAT SPECS",
            command=self._start,
            height=48,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            font=_font(17, "bold"),
        )
        self.run_button.pack(side="left", fill="x", expand=True)
        self.open_button = ctk.CTkButton(
            action_row,
            text="Open Output Folder",
            command=self._open_output,
            width=170,
            height=48,
            state="disabled",
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            font=_font(14),
        )
        self.open_button.pack(side="left", padx=(10, 0))

        self.active_run_label = ctk.CTkLabel(
            card,
            text="No run has started.",
            wraplength=870,
            justify="left",
            anchor="w",
            text_color=COLORS["muted"],
            font=_font(12),
        )
        self.active_run_label.pack(fill="x", padx=22, pady=(0, 18))

        status_row = ctk.CTkFrame(shell, fg_color="transparent")
        status_row.pack(fill="x", pady=(15, 7))
        self.status_label = ctk.CTkLabel(
            status_row,
            text="Ready",
            text_color=COLORS["muted"],
            font=_font(13),
        )
        self.status_label.pack(side="left")
        self.progress = ctk.CTkProgressBar(
            status_row,
            mode="indeterminate",
            width=190,
            height=8,
            progress_color=COLORS["accent"],
            fg_color=COLORS["border"],
        )
        self.progress.pack(side="right")
        self.progress.set(0)

        self.log_box = ctk.CTkTextbox(
            shell,
            height=170,
            fg_color=COLORS["card"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["secondary"],
            font=_font(13, family=MONO_FONT),
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

    def _section_label(self, parent: ctk.CTkFrame, text: str, top: int = 20) -> None:
        ctk.CTkLabel(
            parent,
            text=text,
            text_color=COLORS["text"],
            font=_font(14, "bold"),
        ).pack(anchor="w", padx=22, pady=(top, 8))

    def _path_row(
        self,
        parent: ctk.CTkFrame,
        variable: ctk.StringVar,
        placeholder: str,
        command,
        button_text: str,
    ) -> tuple[ctk.CTkEntry, ctk.CTkButton]:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=22)
        entry = ctk.CTkEntry(
            row,
            textvariable=variable,
            placeholder_text=placeholder,
            height=40,
            fg_color=COLORS["input"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            font=_font(14),
        )
        entry.pack(side="left", fill="x", expand=True)
        if variable is self.output_var:
            self.output_entry = entry
            entry.bind(
                "<KeyPress>",
                lambda _event: setattr(self, "output_is_automatic", False),
            )
        button = ctk.CTkButton(
            row,
            text=button_text,
            command=command,
            width=126,
            height=40,
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            font=_font(14),
        )
        button.pack(side="left", padx=(10, 0))
        self.run_affecting_controls.extend((entry, button))
        return entry, button

    def _choose_architect(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose architect template",
            filetypes=[("Word documents", "*.docx")],
        )
        if selected:
            self.architect_var.set(selected)

    def _add_target_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Choose target specifications",
            filetypes=[("Word documents", "*.docx")],
        )
        if selected:
            self.target_inputs.extend(Path(item) for item in selected)
            self._deduplicate_target_inputs()
            self._set_default_output()
            self._refresh_target_preview()

    def _add_target_folder(self) -> None:
        selected = filedialog.askdirectory(title="Choose folder containing target specs")
        if selected:
            self.target_inputs.append(Path(selected))
            self._deduplicate_target_inputs()
            self._set_default_output()
            self._refresh_target_preview()

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="Choose output folder")
        if selected:
            self.output_var.set(selected)
            self.output_is_automatic = False

    def _deduplicate_target_inputs(self) -> None:
        unique: dict[str, Path] = {}
        for item in self.target_inputs:
            unique.setdefault(os.path.normcase(str(item.resolve())), item)
        self.target_inputs = list(unique.values())

    def _set_default_output(self) -> None:
        if (self.output_var.get().strip() and not self.output_is_automatic) or not self.target_inputs:
            return
        first = self.target_inputs[0]
        base = first if first.is_dir() else first.parent
        self.output_var.set(str(base / "Formatted Specs"))
        self.output_is_automatic = True

    def _clear_targets(self) -> None:
        self.target_inputs.clear()
        if self.output_is_automatic:
            self.output_var.set("")
            self.output_is_automatic = False
        self._refresh_target_preview()

    def _refresh_target_preview(self) -> None:
        self.target_box.configure(state="normal")
        self.target_box.delete("1.0", "end")
        if not self.target_inputs:
            text = "No target specifications selected. Add files or a folder."
        else:
            try:
                targets = collect_target_specs(self.target_inputs)
                lines = [f"{len(targets)} target specification(s)"]
                lines.extend(f"  • {item.name}" for item in targets[:6])
                if len(targets) > 6:
                    lines.append(f"  • … and {len(targets) - 6} more")
                text = "\n".join(lines)
            except Exception as exc:
                text = f"Could not preview targets: {exc}"
        self.target_box.insert("1.0", text)
        self.target_box.configure(state="disabled")

    def _toggle_key(self) -> None:
        self.api_entry.configure(show="" if self.show_key_var.get() else "•")

    def _toggle_advanced(self) -> None:
        self.advanced_visible = not self.advanced_visible
        if self.advanced_visible:
            self.advanced_frame.pack(
                fill="x",
                padx=22,
                pady=(8, 0),
                before=self.action_row,
            )
            self.advanced_button.configure(text="Advanced settings  ▾")
        else:
            self.advanced_frame.pack_forget()
            self.advanced_button.configure(text="Advanced settings  ▸")

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{timestamp}] {message.rstrip()}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _lock_run_controls(self) -> None:
        """Disable run inputs while remembering their exact prior states."""

        if self._locked_run_control_states:
            return
        for control in self.run_affecting_controls:
            try:
                previous_state = str(control.cget("state"))
            except Exception:
                previous_state = "normal"
            self._locked_run_control_states.append((control, previous_state))
            control.configure(state="disabled")

    def _unlock_run_controls(self) -> None:
        """Restore the states captured by :meth:`_lock_run_controls`."""

        states = self._locked_run_control_states
        self._locked_run_control_states = []
        for control, previous_state in states:
            try:
                control.configure(state=previous_state)
            except Exception:
                # A control can disappear only while the window is being torn down.
                continue

    def _show_run_summary(self, *, active: bool) -> None:
        if self.active_run_summary is None:
            return
        self.active_run_label.configure(
            text=active_run_summary_text(self.active_run_summary, active=active),
            text_color=COLORS["secondary"] if active else COLORS["muted"],
        )

    def _start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return
        architect = self.architect_var.get().strip()
        output = self.output_var.get().strip()
        if not architect:
            messagebox.showerror("Missing architect template", "Choose the architect's DOCX template.")
            return
        if not self.target_inputs:
            messagebox.showerror("Missing target specs", "Add at least one target specification.")
            return
        if not output:
            messagebox.showerror("Missing output folder", "Choose an output folder.")
            return

        conversion_mode = self.conversion_mode_var.get()
        active_run = ActiveRunSummary(
            architect_template=Path(architect),
            target_inputs=tuple(self.target_inputs),
            output_root=Path(output),
            conversion_mode=conversion_mode,
            reuse_template_analysis=self.reuse_var.get(),
            max_workers=int(self.workers_var.get()),
        )
        api_key = self.api_key_var.get()
        self.last_result = None
        self.failed_run_dir = None
        self.active_output_dir = active_run.output_root
        self.active_run_summary = active_run
        self._clear_log()
        mode_label = output_mode_label(conversion_mode)
        self._append_log(f"Starting specification processing. Output mode: {mode_label}")
        self._append_log(
            f"Run inputs locked: {len(active_run.target_inputs)} target selection(s), "
            f"{active_run.max_workers} worker(s), output root {active_run.output_root}"
        )
        self._show_run_summary(active=True)
        self.run_button.configure(state="disabled", text="PROCESSING...")
        self._lock_run_controls()
        self.open_button.configure(state="disabled")
        self.status_label.configure(text="Checking files", text_color=COLORS["secondary"])
        self.progress.start()
        self.worker = FormatWorker(
            architect_template=active_run.architect_template,
            target_inputs=active_run.target_inputs,
            output_dir=active_run.output_root,
            api_key=api_key,
            reuse_template_analysis=active_run.reuse_template_analysis,
            max_workers=active_run.max_workers,
            conversion_mode=active_run.conversion_mode,
            events=self.events,
        )
        self.worker.start()

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "progress":
                    self.status_label.configure(text=str(payload))
                    self._append_log(str(payload))
                elif kind == "complete":
                    self._handle_complete(payload)
                elif kind == "error":
                    self._handle_error(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _finish_busy_state(self) -> None:
        self.progress.stop()
        self.progress.set(0)
        self.run_button.configure(state="normal", text="FORMAT SPECS")
        self._unlock_run_controls()
        self._show_run_summary(active=False)

    def _handle_complete(self, result: FormatRunResult) -> None:
        self._finish_busy_state()
        self.last_result = result
        run_dir = result_run_directory(result) or self.active_output_dir
        status, summary = summarize_batch_results(result.targets)
        color = {
            "success": COLORS["success"],
            "partial": COLORS["warning"],
            "failed": COLORS["error"],
        }.get(status, COLORS["muted"])
        self.status_label.configure(text=summary, text_color=color)
        run_id = getattr(result, "run_id", None)
        if run_id:
            self._append_log(f"Run ID: {run_id}")
        if run_dir is not None:
            self._append_log(f"Run folder: {run_dir}")
        manifest_path = getattr(result, "manifest_path", None)
        if manifest_path:
            self._append_log(f"Run manifest: {manifest_path}")
        for item in result.targets:
            for line in target_result_log_lines(item):
                self._append_log(line)
            if item.success and item.output_path is not None:
                self._append_log(f"Output: {item.output_path}")
            else:
                self._append_log(f"Failed: {item.source_path.name} — {item.error}")
        self.open_button.configure(state="normal" if run_dir is not None else "disabled")
        if run_dir is not None:
            self._append_log(f"Persisted run log: {run_dir / 'run.log'}")
        self.active_output_dir = None
        if status == "success":
            messagebox.showinfo("Formatting complete", summary, parent=self)
        elif status == "partial":
            messagebox.showwarning(
                "Formatting partially complete",
                f"{summary}. Successful outputs are available in the output folder.",
                parent=self,
            )
        else:
            messagebox.showerror("Formatting failed", summary, parent=self)

    def _handle_error(self, payload: dict) -> None:
        self._finish_busy_state()
        message = payload.get("message") or "Formatting failed."
        self.status_label.configure(text="Formatting failed", text_color=COLORS["error"])
        self._append_log(message)
        run_dir = payload.get("run_dir")
        self.failed_run_dir = Path(run_dir) if run_dir else None
        if self.failed_run_dir is not None:
            self._append_log(f"Run folder: {self.failed_run_dir}")
            manifest_path = payload.get("manifest_path")
            if manifest_path:
                self._append_log(f"Run manifest: {manifest_path}")
            self.open_button.configure(state="normal")
        self.active_output_dir = None
        messagebox.showerror("Formatting failed", message, parent=self)

    def _open_output(self) -> None:
        output = result_run_directory(self.last_result) if self.last_result is not None else None
        if output is None:
            output = self.failed_run_dir
        if output is None:
            output = Path(self.output_var.get().strip())
        if not output.is_dir():
            messagebox.showerror("Output folder unavailable", f"Folder not found: {output}")
            return
        try:
            os.startfile(str(output))  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror("Could not open folder", str(exc))

    def _on_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showwarning(
                "Formatting is still running",
                "Wait for the active formatting run to finish before closing.",
                parent=self,
            )
            return
        self.destroy()


def main() -> None:
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
