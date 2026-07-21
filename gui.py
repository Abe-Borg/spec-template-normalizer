"""Single-window GUI for architect-template specification formatting."""

from __future__ import annotations

import os
import queue
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Tuple

import customtkinter as ctk
from tkinter import filedialog, messagebox

from spec_formatter import __version__, secrets, updates
from spec_formatter.app_paths import default_config_dir
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
        stored_key = secrets.load_api_key()
        self.api_key_var = ctk.StringVar(
            value=stored_key or os.environ.get("ANTHROPIC_API_KEY", "")
        )
        self.show_key_var = ctk.BooleanVar(value=False)
        self.remember_key_var = ctk.BooleanVar(value=bool(stored_key))
        self.reuse_var = ctk.BooleanVar(value=True)
        self.workers_var = ctk.StringVar(value="3")
        self.conversion_mode_var = ctk.StringVar(value=FORMAT_ONLY)
        self.mode_controls: list[ctk.CTkRadioButton] = []
        self.target_inputs: list[Path] = []
        self.output_is_automatic = False
        self.events: queue.Queue = queue.Queue()
        self.worker: Optional[FormatWorker] = None
        self.last_result: Optional[FormatRunResult] = None
        self.active_output_dir: Optional[Path] = None
        self.advanced_visible = False

        self._update_state_path = updates.default_state_path()
        self._update_checking = False
        self._update_downloading = False
        self._update_download_cancelled = False
        self._update_dialog: Optional[ctk.CTkToplevel] = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_events)
        # Silent once-a-day update check, shortly after the window paints.
        self.after(1500, self._maybe_auto_check_for_updates)

    def _build_ui(self) -> None:
        shell = ctk.CTkFrame(self, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=28, pady=24)

        # Footer first so it reserves the bottom edge below the log box.
        self._build_footer(shell)

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
        self._path_row(
            card,
            self.architect_var,
            "Select the architect's .docx template",
            self._choose_architect,
            "Choose File",
        )

        self._section_label(card, "2   Target specifications", top=18)
        target_toolbar = ctk.CTkFrame(card, fg_color="transparent")
        target_toolbar.pack(fill="x", padx=22)
        ctk.CTkButton(
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
        ).pack(side="left")
        ctk.CTkButton(
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
        ).pack(side="left", padx=(8, 0))
        ctk.CTkButton(
            target_toolbar,
            text="Clear",
            command=self._clear_targets,
            width=76,
            height=36,
            font=_font(14),
            fg_color="transparent",
            hover_color=COLORS["border"],
            text_color=COLORS["secondary"],
        ).pack(side="right")

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
        self._path_row(
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
        ctk.CTkCheckBox(
            key_row,
            text="Show",
            variable=self.show_key_var,
            command=self._toggle_key,
            width=72,
            text_color=COLORS["secondary"],
            font=_font(13),
            fg_color=COLORS["accent"],
        ).pack(side="left", padx=(12, 0))
        ctk.CTkCheckBox(
            key_row,
            text="Remember",
            variable=self.remember_key_var,
            command=self._on_remember_key_toggled,
            width=104,
            text_color=COLORS["secondary"],
            font=_font(13),
            fg_color=COLORS["accent"],
        ).pack(side="left", padx=(12, 0))

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
        ctk.CTkCheckBox(
            self.advanced_frame,
            text="Reuse analysis when the architect template has not changed",
            variable=self.reuse_var,
            text_color=COLORS["secondary"],
            font=_font(13),
            fg_color=COLORS["accent"],
        ).pack(side="left", padx=14, pady=12)
        workers = ctk.CTkFrame(self.advanced_frame, fg_color="transparent")
        workers.pack(side="right", padx=14, pady=8)
        ctk.CTkLabel(
            workers,
            text="Concurrent files",
            text_color=COLORS["secondary"],
            font=_font(13),
        ).pack(side="left", padx=(0, 8))
        ctk.CTkOptionMenu(
            workers,
            values=["1", "2", "3", "4", "5", "6"],
            variable=self.workers_var,
            width=66,
            height=30,
            fg_color=COLORS["border"],
            button_color=COLORS["accent"],
            button_hover_color=COLORS["accent_hover"],
            font=_font(13),
        ).pack(side="left")

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
    ) -> None:
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
        ctk.CTkButton(
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
        ).pack(side="left", padx=(10, 0))

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

        self.last_result = None
        self.active_output_dir = Path(output)
        self._clear_log()
        conversion_mode = self.conversion_mode_var.get()
        mode_label = (
            "Canadian CSC PageFormat conversion"
            if conversion_mode == CSI_TO_CANADIAN
            else "Format only"
        )
        self._append_log(f"Starting specification processing. Output mode: {mode_label}")
        self.run_button.configure(state="disabled", text="PROCESSING...")
        for control in self.mode_controls:
            control.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.status_label.configure(text="Checking files", text_color=COLORS["secondary"])
        self.progress.start()
        if self.remember_key_var.get():
            secrets.save_api_key(self.api_key_var.get())
        self.worker = FormatWorker(
            architect_template=Path(architect),
            target_inputs=tuple(self.target_inputs),
            output_dir=Path(output),
            api_key=self.api_key_var.get(),
            reuse_template_analysis=self.reuse_var.get(),
            max_workers=int(self.workers_var.get()),
            conversion_mode=conversion_mode,
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
        for control in self.mode_controls:
            control.configure(state="normal")

    def _handle_complete(self, result: FormatRunResult) -> None:
        self._finish_busy_state()
        self.last_result = result
        status, summary = summarize_batch_results(result.targets)
        color = {
            "success": COLORS["success"],
            "partial": COLORS["warning"],
            "failed": COLORS["error"],
        }.get(status, COLORS["muted"])
        self.status_label.configure(text=summary, text_color=color)
        for item in result.targets:
            if item.success and item.output_path is not None:
                self._append_log(f"Output: {item.output_path}")
            else:
                self._append_log(f"Failed: {item.source_path.name} — {item.error}")
            for line in conversion_report_log_lines(item):
                self._append_log(line)
        self.open_button.configure(state="normal")
        log_path = self._save_log(result.output_dir)
        if log_path is not None:
            self._append_log(f"Log saved: {log_path}")
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
        output_dir = self.active_output_dir
        if output_dir is not None:
            diagnostic = self._save_diagnostic(
                output_dir,
                payload.get("traceback", ""),
            )
            if diagnostic is not None:
                self._append_log(f"Diagnostic log saved: {diagnostic}")
            self._save_log(output_dir)
        self.active_output_dir = None
        messagebox.showerror("Formatting failed", message, parent=self)

    def _save_diagnostic(self, output_dir: Path, traceback_text: str) -> Optional[Path]:
        if not traceback_text:
            return None
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            path = output_dir / f"spec_formatter_diagnostic_{stamp}.log"
            path.write_text(traceback_text.rstrip() + "\n", encoding="utf-8")
            return path
        except OSError:
            return None

    def _save_log(self, output_dir: Path) -> Optional[Path]:
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            log_path = output_dir / f"spec_formatter_{stamp}.log"
            content = self.log_box.get("1.0", "end").rstrip() + "\n"
            log_path.write_text(content, encoding="utf-8")
            return log_path
        except OSError:
            return None

    def _open_output(self) -> None:
        output = (
            self.last_result.output_dir
            if self.last_result is not None
            else Path(self.output_var.get().strip())
        )
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

    # ------------------------------------------------------------------
    # API key persistence (OS keyring)
    # ------------------------------------------------------------------

    def _on_remember_key_toggled(self) -> None:
        if self.remember_key_var.get():
            secrets.save_api_key(self.api_key_var.get())
        else:
            secrets.clear_api_key()

    # ------------------------------------------------------------------
    # Footer + self-update flow
    # ------------------------------------------------------------------

    def _build_footer(self, parent: ctk.CTkFrame) -> None:
        footer = ctk.CTkFrame(parent, fg_color="transparent")
        footer.pack(side="bottom", fill="x", pady=(10, 0))
        ctk.CTkLabel(
            footer,
            text=f"v{__version__}",
            text_color=COLORS["muted"],
            font=_font(12),
        ).pack(side="left")
        self.update_status_label = ctk.CTkLabel(
            footer,
            text="",
            text_color=COLORS["muted"],
            font=_font(12),
        )
        self.update_status_label.pack(side="left", padx=(10, 0))
        self.check_update_button = ctk.CTkButton(
            footer,
            text="Check for Updates",
            command=self._on_check_for_updates_clicked,
            width=150,
            height=28,
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            font=_font(13),
        )
        self.check_update_button.pack(side="right")

    def _set_update_status(self, text: str, *, color: Optional[str] = None) -> None:
        self.update_status_label.configure(text=text, text_color=color or COLORS["muted"])

    def _maybe_auto_check_for_updates(self) -> None:
        # A background convenience check must never break app startup.
        try:
            if updates.update_check_disabled():
                return
            state = updates.load_state(self._update_state_path)
            if not updates.should_auto_check(state, now=datetime.now()):
                return
            self._start_update_check(manual=False)
        except Exception:
            pass

    def _on_check_for_updates_clicked(self) -> None:
        self._start_update_check(manual=True)

    def _start_update_check(self, *, manual: bool) -> None:
        if self._update_checking:
            return
        self._update_checking = True
        if manual:
            self.check_update_button.configure(state="disabled", text="Checking…")
            self._set_update_status("Checking for updates…")
        thread = threading.Thread(
            target=self._update_check_worker, args=(manual,), daemon=True
        )
        thread.start()

    def _update_check_worker(self, manual: bool) -> None:
        result = updates.check_for_update(__version__)
        # Record the check regardless of outcome so the daily throttle holds.
        try:
            state = updates.load_state(self._update_state_path)
            updates.record_check(state, now=datetime.now())
            updates.save_state(self._update_state_path, state)
        except Exception:
            pass
        self.after(0, lambda: self._on_update_check_done(result, manual))

    def _on_update_check_done(self, result, manual: bool) -> None:
        self._update_checking = False
        self.check_update_button.configure(state="normal", text="Check for Updates")
        status = result.status
        if status == updates.STATUS_UPDATE_AVAILABLE and result.info is not None:
            info = result.info
            self._set_update_status(
                f"Update available: v{info.version}", color=COLORS["accent"]
            )
            if not manual:
                state = updates.load_state(self._update_state_path)
                if updates.version_is_skipped(state, info.version):
                    return
            self._show_update_dialog(info)
        elif status == updates.STATUS_UP_TO_DATE:
            self._set_update_status("You're up to date.")
            if manual:
                messagebox.showinfo(
                    "No updates",
                    f"You're running the latest version (v{__version__}).",
                    parent=self,
                )
        elif status == updates.STATUS_DISABLED:
            self._set_update_status("Update checks are disabled.")
            if manual:
                messagebox.showinfo(
                    "Updates disabled",
                    f"Update checks are turned off via {updates.ENV_DISABLE}.",
                    parent=self,
                )
        else:  # STATUS_ERROR
            self._set_update_status("Update check failed.")
            if manual:
                messagebox.showwarning(
                    "Update check failed",
                    "Could not check for updates:\n\n"
                    f"{result.error}\n\n"
                    "You can download the latest version manually from:\n"
                    f"{updates.releases_page_url()}",
                    parent=self,
                )

    def _show_update_dialog(self, info) -> None:
        if self._update_dialog is not None and self._update_dialog.winfo_exists():
            self._update_dialog.lift()
            self._update_dialog.focus_force()
            return

        win = ctk.CTkToplevel(self)
        self._update_dialog = win
        win.title("Update available")
        win.geometry("560x480")
        win.minsize(460, 380)
        win.configure(fg_color=COLORS["bg"])
        win.transient(self)
        win.protocol("WM_DELETE_WINDOW", self._close_update_dialog)
        win.after(150, lambda: self._grab_dialog(win))

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(
            body,
            text=f"Version {info.version} is available",
            text_color=COLORS["text"],
            font=_font(20, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            body,
            text=(
                f"You have v{__version__}. Download and install the update? "
                "The app will close so the installer can replace it."
            ),
            text_color=COLORS["secondary"],
            font=_font(13),
            wraplength=500,
            justify="left",
        ).pack(anchor="w", pady=(6, 14))

        # Button bar sits at the very bottom.
        button_bar = ctk.CTkFrame(body, fg_color="transparent")
        button_bar.pack(side="bottom", fill="x", pady=(14, 0))
        self._update_download_button = ctk.CTkButton(
            button_bar,
            text="Download & Install",
            command=lambda: self._start_update_download(info),
            width=170,
            height=34,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            text_color="#FFFFFF",
            font=_font(14, "bold"),
        )
        self._update_download_button.pack(side="right")
        self._update_later_button = ctk.CTkButton(
            button_bar,
            text="Later",
            command=self._close_update_dialog,
            width=84,
            height=34,
            fg_color=COLORS["input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            font=_font(14),
        )
        self._update_later_button.pack(side="right", padx=(0, 8))
        self._update_skip_button = ctk.CTkButton(
            button_bar,
            text="Skip this Version",
            command=lambda: self._skip_update_version(info),
            width=150,
            height=34,
            fg_color="transparent",
            hover_color=COLORS["border"],
            text_color=COLORS["secondary"],
            font=_font(13),
        )
        self._update_skip_button.pack(side="left")

        # Clickable releases link, above the buttons.
        link = ctk.CTkLabel(
            body,
            text="View this release on GitHub",
            text_color=COLORS["accent"],
            font=_font(12),
            cursor="hand2",
        )
        link.pack(side="bottom", anchor="w", pady=(10, 0))
        link.bind("<Button-1>", lambda _event: self._open_releases_page())

        # Progress row: created now, packed only once a download starts.
        self._update_progress = ctk.CTkProgressBar(
            body, height=10, progress_color=COLORS["accent"], fg_color=COLORS["border"]
        )
        self._update_progress.set(0)
        self._update_progress_label = ctk.CTkLabel(
            body, text="", text_color=COLORS["muted"], font=_font(12)
        )

        ctk.CTkLabel(
            body,
            text="What's new",
            text_color=COLORS["text"],
            font=_font(14, "bold"),
        ).pack(anchor="w", pady=(4, 6))
        notes_box = ctk.CTkTextbox(
            body,
            height=150,
            fg_color=COLORS["card"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["secondary"],
            font=_font(12),
        )
        notes_box.pack(fill="both", expand=True)
        notes_box.insert("1.0", info.notes or "No release notes were provided.")
        notes_box.configure(state="disabled")

    def _grab_dialog(self, win: ctk.CTkToplevel) -> None:
        try:
            if win.winfo_exists():
                win.grab_set()
        except Exception:
            pass

    def _start_update_download(self, info) -> None:
        if self._update_downloading:
            self._set_update_status("A download is already in progress…")
            return
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo(
                "Formatting in progress",
                "Wait for the active formatting run to finish before updating.",
                parent=self,
            )
            return
        self._update_downloading = True
        self._update_download_cancelled = False
        for button in (
            self._update_download_button,
            self._update_later_button,
            self._update_skip_button,
        ):
            button.configure(state="disabled")
        self._update_progress.set(0)
        self._update_progress_label.pack(side="bottom", fill="x", pady=(6, 0))
        self._update_progress.pack(side="bottom", fill="x", pady=(6, 0))
        self._update_progress_label.configure(text="Starting download…")
        thread = threading.Thread(
            target=self._update_download_worker, args=(info,), daemon=True
        )
        thread.start()

    def _update_download_worker(self, info) -> None:
        dest_dir = default_config_dir() / "updates"

        def _progress(done: int, total: int) -> None:
            self.after(0, lambda: self._on_update_download_progress(done, total))

        try:
            path = updates.download_installer(info, dest_dir, progress=_progress)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda: self._on_update_download_error(message))
            return
        self.after(0, lambda: self._on_update_download_done(path))

    def _on_update_download_progress(self, done: int, total: int) -> None:
        if self._update_download_cancelled or self._update_dialog is None:
            return
        mb = 1024 * 1024
        try:
            if total > 0:
                self._update_progress.set(min(1.0, done / total))
                self._update_progress_label.configure(
                    text=f"Downloading… {done // mb} / {total // mb} MB"
                )
            else:
                self._update_progress_label.configure(text=f"Downloading… {done // mb} MB")
        except Exception:
            pass

    def _on_update_download_done(self, path: Path) -> None:
        self._update_downloading = False
        # Respect a dialog the user dismissed mid-download: keep the verified
        # file cached but do not pop a surprise "install & quit" prompt.
        if self._update_download_cancelled or self._update_dialog is None:
            self._set_update_status("Update downloaded — install it later.")
            return
        self._update_progress.set(1.0)
        self._update_progress_label.configure(text="Download verified.")
        if messagebox.askyesno(
            "Install update",
            "The update was downloaded and verified. Install it now? "
            "The app will close so the installer can replace it.",
            parent=self,
        ):
            try:
                updates.spawn_installer(path)
            except Exception as exc:
                messagebox.showerror(
                    "Could not start installer",
                    f"The installer could not be launched:\n\n{exc}\n\n"
                    f"You can run it manually from:\n{path}",
                    parent=self,
                )
                self._reset_update_dialog_buttons()
                return
            self._close_update_dialog()
            self.quit()
        else:
            self._reset_update_dialog_buttons()
            self._set_update_status("Update downloaded (not installed).")

    def _on_update_download_error(self, message: str) -> None:
        self._update_downloading = False
        if self._update_download_cancelled or self._update_dialog is None:
            self._set_update_status("Update download failed.")
            return
        self._reset_update_dialog_buttons()
        messagebox.showerror(
            "Download failed",
            f"The update could not be downloaded:\n\n{message}\n\n"
            "You can download the latest version manually from:\n"
            f"{updates.releases_page_url()}",
            parent=self,
        )

    def _reset_update_dialog_buttons(self) -> None:
        for button in (
            self._update_download_button,
            self._update_later_button,
            self._update_skip_button,
        ):
            try:
                button.configure(state="normal")
            except Exception:
                pass
        try:
            self._update_progress.pack_forget()
            self._update_progress_label.pack_forget()
        except Exception:
            pass

    def _skip_update_version(self, info) -> None:
        state = updates.load_state(self._update_state_path)
        updates.mark_skipped(state, info.version)
        updates.save_state(self._update_state_path, state)
        self._set_update_status(f"Skipped v{info.version}.")
        self._close_update_dialog()

    def _open_releases_page(self) -> None:
        webbrowser.open(updates.releases_page_url())

    def _close_update_dialog(self) -> None:
        if self._update_downloading:
            # The daemon worker can't be killed, but flag it so a completing
            # download doesn't pop a stray install prompt after the dialog closes.
            self._update_download_cancelled = True
        win = self._update_dialog
        self._update_dialog = None
        if win is not None:
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                win.destroy()
            except Exception:
                pass


def main() -> None:
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
