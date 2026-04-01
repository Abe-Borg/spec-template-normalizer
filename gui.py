"""
Tkinter GUI wrapper for the Phase 1 DOCX CSI Normalizer pipeline.

This is a thin wrapper — no business logic. It imports and calls the same
library functions as the smoke test.
"""
from __future__ import annotations

import json
import re
import os
import queue
import shutil
import sys
import threading
import traceback
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, scrolledtext
from pathlib import Path
from typing import Optional
from datetime import datetime


COLORS = {
    "bg_dark": "#0D0D0D",
    "bg_card": "#1A1A1A",
    "bg_input": "#252525",
    "border": "#333333",
    "text_primary": "#FFFFFF",
    "text_secondary": "#B0B0B0",
    "text_muted": "#707070",
    "accent": "#3B82F6",
    "accent_hover": "#2563EB",
    "accent_glow": "#60A5FA",
    "success": "#22C55E",
    "success_glow": "#4ADE80",
    "warning": "#F59E0B",
    "error": "#EF4444",
}

CLEANUP_ON_FAILURE = True




def _load_prompt_file(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Missing required prompt file: {path}")
    try:
        return path.read_text(encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Failed reading prompt file {path}: {exc}") from exc


class LogRedirector:
    """Redirects writes to a thread-safe queue for GUI consumption."""

    def __init__(self, log_queue: queue.Queue) -> None:
        self._queue = log_queue

    def write(self, text: str) -> None:
        if text.strip():
            self._queue.put(text)

    def flush(self) -> None:
        pass


class PipelineThread(threading.Thread):
    """Runs the full Phase 1 pipeline in a background thread."""

    def __init__(
        self,
        docx_path: str,
        api_key: str,
        output_dir: Optional[str],
        log_queue: queue.Queue,
        result_queue: queue.Queue,
    ) -> None:
        super().__init__(daemon=True)
        self.docx_path = docx_path
        self.api_key = api_key
        self.output_dir = output_dir
        self.log_queue = log_queue
        self.result_queue = result_queue

    def _log(self, msg: str) -> None:
        self.log_queue.put(msg)

    def _choose_extract_dir(self, docx_path: Path) -> Path:
        base_dir = Path(self.output_dir) if self.output_dir else docx_path.parent
        candidate = base_dir / f"{docx_path.stem}_extracted"
        if not candidate.exists():
            return candidate
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return base_dir / f"{docx_path.stem}_extracted__{stamp}"

    def run(self) -> None:
        extract_dir: Optional[Path] = None
        try:
            from docx_decomposer import (
                extract_docx,
                build_slim_bundle,
                apply_instructions,
                build_style_registry_dict,
            )
            from llm_classifier import classify_document, compute_coverage

            docx_path = Path(self.docx_path)
            extract_dir = self._choose_extract_dir(docx_path)

            # 1) Extract
            self._log(f"Extracting {docx_path.name}...")
            extract_docx(docx_path, extract_dir)

            # 2) Build slim bundle
            self._log("Building slim bundle...")
            bundle = build_slim_bundle(extract_dir)

            n_paras = len(bundle.get("paragraphs", []))
            self._log(f"Slim bundle: {n_paras} paragraphs")

            # 3) Read prompts
            script_dir = Path(__file__).resolve().parent
            master_prompt = _load_prompt_file(script_dir / "master_prompt.txt")
            run_instruction = _load_prompt_file(script_dir / "run_instruction_prompt.txt")

            # 4) Classify
            self._log("Classifying via LLM...")
            instructions = classify_document(
                slim_bundle=bundle,
                master_prompt=master_prompt,
                run_instruction=run_instruction,
                api_key=self.api_key,
            )

            # 5) Coverage
            coverage, styled, classifiable = compute_coverage(bundle, instructions)
            coverage_msg = f"Coverage: {coverage:.1%} ({styled}/{classifiable})"
            self._log(coverage_msg)
            if coverage < 1.0:
                raise ValueError(f"Coverage must be 100% for classifiable paragraphs; got {coverage_msg}")

            # 6) Apply
            self._log("Applying instructions...")
            # Redirect stdout so apply_instructions prints go to our log
            old_stdout = sys.stdout
            sys.stdout = LogRedirector(self.log_queue)
            try:
                apply_instructions(extract_dir, instructions)
            finally:
                sys.stdout = old_stdout

            # 7) Build registries in memory
            self._log("Building style registry...")
            style_registry = build_style_registry_dict(
                extract_dir,
                docx_path.name,
                instructions,
                pre_apply_bundle=bundle,
            )

            self._log("Extracting environment...")
            from arch_env_extractor import extract_arch_template_registry
            template_registry = extract_arch_template_registry(extract_dir, docx_path)

            # 8) Validate both registries before writing
            from phase1_validator import validate_phase1_contracts
            self._log("Validating Phase 1 contracts...")
            validate_phase1_contracts(style_registry, template_registry)

            # 9) Write registries (only reached if validation passes)
            reg_path = extract_dir / "arch_style_registry.json"
            reg_path.write_text(json.dumps(style_registry, indent=2), encoding="utf-8")
            self._log(f"Style registry: {reg_path.name}")

            env_path = extract_dir / "arch_template_registry.json"
            env_path.write_text(json.dumps(template_registry, indent=2), encoding="utf-8")
            self._log(f"Environment registry: {env_path.name}")

            # 10) Copy deliverables to output_dir if specified
            output_dir_path = Path(self.output_dir) if self.output_dir else extract_dir
            if output_dir_path != extract_dir:
                output_dir_path.mkdir(parents=True, exist_ok=True)
                shutil.copy2(reg_path, output_dir_path / reg_path.name)
                self._log(f"Copied {reg_path.name} to {output_dir_path}")
                shutil.copy2(env_path, output_dir_path / env_path.name)
                self._log(f"Copied {env_path.name} to {output_dir_path}")

            raw_styles_src = extract_dir / "word" / "styles.xml"
            raw_settings_src = extract_dir / "word" / "settings.xml"
            raw_styles_dst = output_dir_path / "arch_styles_raw.xml"
            raw_settings_dst = output_dir_path / "arch_settings_raw.xml"

            if raw_styles_src.exists():
                shutil.copy2(raw_styles_src, raw_styles_dst)
                self._log(f"Preserved raw styles.xml as {raw_styles_dst.name}")

            if raw_settings_src.exists():
                shutil.copy2(raw_settings_src, raw_settings_dst)
                self._log(f"Preserved raw settings.xml as {raw_settings_dst.name}")

            if extract_dir.resolve() != output_dir_path.resolve():
                try:
                    shutil.rmtree(extract_dir)
                    self._log(f"Cleaned up working directory: {extract_dir.name}")
                except Exception as cleanup_err:
                    self._log(f"Warning: could not clean up {extract_dir.name}: {cleanup_err}")

            self.result_queue.put({
                "success": True,
                "extract_dir": str(extract_dir),
                "output_dir": str(output_dir_path),
                "registry_path": str(output_dir_path / reg_path.name),
                "env_path": str(output_dir_path / env_path.name),
                "coverage": coverage_msg,
            })

        except Exception:
            self._log(f"ERROR:\n{traceback.format_exc()}")
            if CLEANUP_ON_FAILURE and extract_dir and extract_dir.exists():
                try:
                    output_dir_path = Path(self.output_dir) if self.output_dir else extract_dir
                    if extract_dir.resolve() != output_dir_path.resolve():
                        shutil.rmtree(extract_dir)
                        self._log(f"Cleaned up working directory after failure: {extract_dir.name}")
                except Exception:
                    pass
            self.result_queue.put({"success": False})


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("DOCX CSI Normalizer — Phase 1")
        self.geometry("900x750")
        self.minsize(750, 600)
        self.configure(fg_color=COLORS["bg_dark"])

        self.log_queue: queue.Queue = queue.Queue()
        self.result_queue: queue.Queue = queue.Queue()
        self._result: Optional[dict] = None
        self._help_windows: list[ctk.CTkToplevel] = []

        self._inputs_expanded = True
        self._log_expanded = True

        self._build_ui()
        self._poll_queues()

    def _build_ui(self) -> None:
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=24, pady=24)

        # --- Header ---
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")

        ctk.CTkLabel(
            title_row,
            text="DOCX CSI NORMALIZER",
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
        ).pack(side="left")

        help_btn_frame = ctk.CTkFrame(title_row, fg_color="transparent")
        help_btn_frame.pack(side="right")

        ctk.CTkButton(
            help_btn_frame,
            text="How It Works",
            command=lambda: self._show_info_popup("How It Works", HOW_IT_WORKS_TEXT),
            width=100,
            height=32,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            help_btn_frame,
            text="How to Use",
            command=lambda: self._show_info_popup("How to Use", HOW_TO_USE_TEXT),
            width=100,
            height=32,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        ).pack(side="right")

        ctk.CTkLabel(
            header,
            text="Phase 1 Pipeline — Architect Template Analysis",
            text_color=COLORS["text_secondary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).pack(fill="x", pady=(6, 0), anchor="w")

        # --- Inputs card ---
        inputs_card = ctk.CTkFrame(container, fg_color=COLORS["bg_card"], corner_radius=8)
        inputs_card.pack(fill="x", pady=(0, 12))

        inputs_header = ctk.CTkFrame(inputs_card, fg_color="transparent", cursor="hand2")
        inputs_header.pack(fill="x", padx=16, pady=12)
        inputs_header.bind("<Button-1>", self._toggle_inputs)

        self._inputs_arrow = ctk.CTkLabel(
            inputs_header,
            text="▼",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=COLORS["text_muted"],
            width=20,
        )
        self._inputs_arrow.pack(side="left")
        self._inputs_arrow.bind("<Button-1>", self._toggle_inputs)

        inputs_lbl = ctk.CTkLabel(
            inputs_header,
            text="INPUTS",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=COLORS["text_muted"],
        )
        inputs_lbl.pack(side="left", padx=(4, 0))
        inputs_lbl.bind("<Button-1>", self._toggle_inputs)

        self._inputs_content = ctk.CTkFrame(inputs_card, fg_color="transparent")
        self._inputs_content.pack(fill="x", padx=16, pady=(0, 16))

        # Row 0: Template
        ctk.CTkLabel(
            self._inputs_content,
            text="Template",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            width=100,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=8)

        template_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        template_frame.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=8)
        template_frame.columnconfigure(0, weight=1)

        self.path_var = tk.StringVar()
        self.path_entry = ctk.CTkEntry(
            template_frame,
            textvariable=self.path_var,
            placeholder_text="Select architect template .docx",
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
            height=36,
        )
        self.path_entry.grid(row=0, column=0, sticky="ew")

        ctk.CTkButton(
            template_frame,
            text="Browse",
            width=70,
            command=self._browse,
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        ).grid(row=0, column=1, padx=(8, 0))

        # Row 1: API Key
        ctk.CTkLabel(
            self._inputs_content,
            text="API Key",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            width=100,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=8)

        key_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        key_frame.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=8)
        key_frame.columnconfigure(0, weight=1)

        self.key_var = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.key_entry = ctk.CTkEntry(
            key_frame,
            textvariable=self.key_var,
            show="•",
            placeholder_text="sk-ant-...",
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
            height=36,
        )
        self.key_entry.grid(row=0, column=0, sticky="ew")
        self._key_visible = False

        self.key_toggle = ctk.CTkButton(
            key_frame,
            text="Show",
            command=self._toggle_key,
            width=70,
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        )
        self.key_toggle.grid(row=0, column=1, padx=(8, 0))

        # Row 2: Output Folder
        ctk.CTkLabel(
            self._inputs_content,
            text="Output Folder",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            width=100,
            anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=8)

        output_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        output_frame.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=8)
        output_frame.columnconfigure(0, weight=1)

        self.output_dir_var = tk.StringVar()
        self.output_entry = ctk.CTkEntry(
            output_frame,
            textvariable=self.output_dir_var,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=COLORS["bg_input"],
            border_color=COLORS["border"],
            text_color=COLORS["text_primary"],
            height=36,
        )
        self.output_entry.grid(row=0, column=0, sticky="ew")

        ctk.CTkButton(
            output_frame,
            text="Browse",
            command=self._browse_output,
            width=70,
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        ).grid(row=0, column=1, padx=(8, 0))

        self._inputs_content.columnconfigure(1, weight=1)

        # --- Run button ---
        self.run_btn = ctk.CTkButton(
            container,
            text="Run Phase 1",
            command=self._run,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=44,
            corner_radius=8,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        )
        self.run_btn.pack(fill="x", pady=(0, 0))

        self.progress_bar = ctk.CTkProgressBar(
            container,
            height=4,
            corner_radius=2,
            fg_color=COLORS["bg_input"],
            progress_color=COLORS["accent"],
            indeterminate_speed=0.5,
        )

        # --- Log card ---
        log_card = ctk.CTkFrame(container, fg_color=COLORS["bg_card"], corner_radius=8)
        log_card.pack(fill="both", expand=True, pady=(12, 0))

        log_header = ctk.CTkFrame(log_card, fg_color="transparent", cursor="hand2")
        log_header.pack(fill="x", padx=16, pady=12)
        log_header.bind("<Button-1>", self._toggle_log)

        self._log_arrow = ctk.CTkLabel(
            log_header,
            text="▼",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=COLORS["text_muted"],
            width=20,
        )
        self._log_arrow.pack(side="left")
        self._log_arrow.bind("<Button-1>", self._toggle_log)

        log_lbl = ctk.CTkLabel(
            log_header,
            text="ACTIVITY LOG",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=COLORS["text_muted"],
        )
        log_lbl.pack(side="left", padx=(4, 0))
        log_lbl.bind("<Button-1>", self._toggle_log)

        ctk.CTkButton(
            log_header,
            text="Clear",
            width=50,
            height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            hover_color=COLORS["bg_input"],
            text_color=COLORS["text_muted"],
            command=self._clear_log,
        ).pack(side="right")

        self._log_content = ctk.CTkFrame(log_card, fg_color="transparent")
        self._log_content.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.log_text = ctk.CTkTextbox(
            self._log_content,
            fg_color=COLORS["bg_input"],
            corner_radius=4,
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=COLORS["text_secondary"],
            wrap="word",
            state="disabled",
            activate_scrollbars=True,
        )
        self.log_text.pack(fill="both", expand=True)

        # --- Status bar ---
        status_frame = ctk.CTkFrame(container, fg_color="transparent")
        status_frame.pack(fill="x", pady=(8, 0))

        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.status_var,
            anchor="w",
            text_color=COLORS["text_secondary"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.status_label.pack(side="left")

    def _toggle_inputs(self, event=None) -> None:
        if self._inputs_expanded:
            self._inputs_content.pack_forget()
            self._inputs_arrow.configure(text="▶")
            self._inputs_expanded = False
        else:
            self._inputs_content.pack(fill="x", padx=16, pady=(0, 16))
            self._inputs_arrow.configure(text="▼")
            self._inputs_expanded = True

    def _toggle_log(self, event=None) -> None:
        if self._log_expanded:
            self._log_content.pack_forget()
            self._log_arrow.configure(text="▶")
            self._log_expanded = False
        else:
            self._log_content.pack(fill="both", expand=True, padx=16, pady=(0, 16))
            self._log_arrow.configure(text="▼")
            self._log_expanded = True

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _show_info_popup(self, title: str, body: str) -> None:
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry("760x650")
        popup.minsize(640, 480)
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self)
        popup.grab_set()
        popup.lift()
        popup.focus_force()

        frame = ctk.CTkFrame(popup, fg_color=COLORS["bg_card"], corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text=title,
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        ).pack(fill="x", padx=14, pady=(14, 8), anchor="w")

        text_frame = ctk.CTkFrame(frame, fg_color=COLORS["bg_input"], corner_radius=4)
        text_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        text_widget = scrolledtext.ScrolledText(
            text_frame,
            wrap="word",
            font=("Segoe UI", 10),
            bg=COLORS["bg_input"],
            fg=COLORS["text_secondary"],
            insertbackground=COLORS["text_primary"],
            relief="flat",
            highlightthickness=0,
            padx=12,
            pady=10,
        )
        text_widget.pack(fill="both", expand=True)
        self._insert_markdown(text_widget, body)
        text_widget.configure(state="disabled")

        ctk.CTkButton(
            frame,
            text="Close",
            width=100,
            height=32,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            command=popup.destroy,
        ).pack(pady=(0, 16))

        self._help_windows.append(popup)
        popup.protocol("WM_DELETE_WINDOW", lambda p=popup: self._close_help_popup(p))

    def _close_help_popup(self, popup: ctk.CTkToplevel) -> None:
        if popup in self._help_windows:
            self._help_windows.remove(popup)
        popup.destroy()

    def _insert_markdown(self, text_widget: scrolledtext.ScrolledText, markdown_text: str) -> None:
        """Render a small markdown subset into a Tk text widget."""
        text_widget.tag_configure("h1", font=("Segoe UI", 15, "bold"), spacing1=8, spacing3=6)
        text_widget.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=6, spacing3=4)
        text_widget.tag_configure("h3", font=("Segoe UI", 11, "bold"), spacing1=4, spacing3=2)
        text_widget.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        text_widget.tag_configure("code", font=("Consolas", 10), background="#2D2D2D")

        for raw_line in markdown_text.strip().splitlines():
            line = raw_line.rstrip()
            stripped = line.strip()

            if not stripped:
                text_widget.insert("end", "\n")
                continue

            heading_level = 0
            if stripped.startswith("### "):
                heading_level = 3
                stripped = stripped[4:]
            elif stripped.startswith("## "):
                heading_level = 2
                stripped = stripped[3:]
            elif stripped.startswith("# "):
                heading_level = 1
                stripped = stripped[2:]

            if heading_level:
                text_widget.insert("end", stripped + "\n", (f"h{heading_level}",))
                continue

            bullet_match = re.match(r"^[-*]\s+(.*)$", stripped)
            numbered_match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
            if bullet_match:
                text_widget.insert("end", "• ")
                self._insert_inline_markdown(text_widget, bullet_match.group(1))
                text_widget.insert("end", "\n")
                continue
            if numbered_match:
                text_widget.insert("end", f"{numbered_match.group(1)}. ")
                self._insert_inline_markdown(text_widget, numbered_match.group(2))
                text_widget.insert("end", "\n")
                continue

            self._insert_inline_markdown(text_widget, stripped)
            text_widget.insert("end", "\n")

    def _insert_inline_markdown(self, text_widget: scrolledtext.ScrolledText, text: str) -> None:
        token_pattern = re.compile(r"(`[^`]+`|\*\*[^*]+\*\*)")
        pos = 0
        for match in token_pattern.finditer(text):
            if match.start() > pos:
                text_widget.insert("end", text[pos:match.start()])

            token = match.group(0)
            if token.startswith("**") and token.endswith("**"):
                text_widget.insert("end", token[2:-2], ("bold",))
            elif token.startswith("`") and token.endswith("`"):
                text_widget.insert("end", token[1:-1], ("code",))
            else:
                text_widget.insert("end", token)

            pos = match.end()

        if pos < len(text):
            text_widget.insert("end", text[pos:])

    def _browse(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.path_var.set(path)
            self.output_dir_var.set(str(Path(path).parent))

    def _browse_output(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir_var.set(folder)

    def _toggle_key(self) -> None:
        self._key_visible = not self._key_visible
        self.key_entry.configure(show="" if self._key_visible else "•")
        self.key_toggle.configure(text="Hide" if self._key_visible else "Show")

    def _set_run_processing(self) -> None:
        self.run_btn.configure(
            text="Processing...",
            state="disabled",
            text_color_disabled="#FFFFFF",
        )

    def _set_run_complete(self) -> None:
        self.run_btn.configure(
            text="✓ Complete",
            fg_color=COLORS["success"],
            state="disabled",
        )
        self.after(2500, self._reset_run_button)

    def _set_run_failed(self) -> None:
        self.run_btn.configure(
            text="✗ Failed",
            fg_color=COLORS["error"],
            state="disabled",
        )
        self.after(2500, self._reset_run_button)

    def _reset_run_button(self) -> None:
        self.run_btn.configure(
            text="Run Phase 1",
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            state="normal",
        )

    def _run(self) -> None:
        docx_path = self.path_var.get().strip()
        api_key = self.key_var.get().strip()

        if not docx_path:
            self.status_var.set("Error: No template selected")
            self.status_label.configure(text_color=COLORS["error"])
            return
        if not Path(docx_path).exists():
            self.status_var.set("Error: File not found")
            self.status_label.configure(text_color=COLORS["error"])
            return
        if not api_key:
            self.status_var.set("Error: No API key")
            self.status_label.configure(text_color=COLORS["error"])
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self._set_run_processing()
        self.status_var.set("Running...")
        self.status_label.configure(text_color=COLORS["text_secondary"])
        self._result = None

        self.progress_bar.pack(fill="x", pady=(8, 0), after=self.run_btn)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        thread = PipelineThread(
            docx_path=docx_path,
            api_key=api_key,
            output_dir=self.output_dir_var.get().strip() or None,
            log_queue=self.log_queue,
            result_queue=self.result_queue,
        )
        thread.start()

    def _poll_queues(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            except queue.Empty:
                break

        try:
            result = self.result_queue.get_nowait()
            self._result = result
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            if result["success"]:
                self._set_run_complete()
                self.status_var.set("Success — " + result.get("coverage", ""))
                self.status_label.configure(text_color=COLORS["success"])

                self.log_text.configure(state="normal")
                self.log_text.insert("end", "\nPhase 1 complete. Both registries ready for Phase 2.\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            else:
                self._set_run_failed()
                self.status_var.set("Failed — see log for details")
                self.status_label.configure(text_color=COLORS["error"])
        except queue.Empty:
            pass

        self.after(100, self._poll_queues)


HOW_TO_USE_TEXT = """
# What You Need Before Starting

- An architect's Word specification template (`.docx`)
- An Anthropic API key (get one at console.anthropic.com)

# Steps

1. **Select the template**
   Click **Browse** next to Template and select the architect's `.docx` spec file.

2. **Enter your API key**
   Paste your Anthropic API key into the API Key field. Click **Show** to verify it if needed.
   The key is pre-filled automatically if the `ANTHROPIC_API_KEY` environment variable is set on your machine.

3. **Set an output folder (optional)**
   By default, output files are saved to the same folder as the `.docx` you selected.
   Click **Browse** next to Output Folder to choose a different location.

4. **Run**
   Click **Run Phase 1**. The activity log will show live progress.
   Processing typically takes 1–3 minutes depending on document length.

5. **Review results**
   When complete, the status bar shows the coverage result.

# Output Files

- `arch_style_registry.json`: Maps CSI structural roles (PART, Article, Paragraph, etc.) to Word paragraph styles.
- `arch_template_registry.json`: A complete snapshot of the architect's formatting environment.

Both files are required inputs for the Phase 2 formatting tool.

# If Something Goes Wrong

- **"Coverage must be 100%"** — The AI didn't classify every paragraph. Click **Run Phase 1** again; this usually resolves on retry.
- **"File not found"** — Make sure the `.docx` file isn't open in Word when you run.
- **"No API key"** — Check that your key is pasted correctly and hasn't expired.
- Any other error — The full error message is in the activity log. The document is never modified in place; re-running is always safe.
"""


HOW_IT_WORKS_TEXT = """
# The Problem This Solves

Every architecture firm formats their specification templates differently.
When a mechanical, electrical, or plumbing consultant needs to reformat their specs to match the architect's style — fonts, indentation, numbering appearance, heading weights — it's tedious manual work that has to be repeated for every project and every architect.

This tool automates the first step: reading and understanding an architect's Word template so that the formatting can be applied to other documents automatically.

# The Two-Phase Pipeline

This tool is **Phase 1** of a two-step process.

- **Phase 1 (this tool):** Analyzes the architect's template and produces two output files that describe its structure and formatting.
- **Phase 2 (separate tool):** Uses those output files to apply the architect's formatting to MEP consultant specs.

# What Phase 1 Actually Does

1. **Unpack the Document**
   A `.docx` file is actually a ZIP archive containing XML files.
   The tool unpacks it into a working folder so it can be read and modified safely.
   Your original file is never touched.

2. **Read the Structure**
   The tool reads every paragraph in the document and records its text, indentation, numbering, and any existing paragraph style.
   It strips out everything that isn't needed for classification and produces a compact summary (the slim bundle), which is what gets sent to the AI.

3. **AI Classification**
   The slim bundle is sent to Claude (Anthropic's AI) with detailed instructions.
   Claude identifies the CSI structural role of each paragraph (Section Title, PART, Article, Paragraph, Subparagraph).

4. **Derive Formatting Locally**
   For each CSI role, the tool identifies a representative exemplar paragraph from the template and extracts exact formatting from it to create a new style.

5. **Apply Styles (Surgically)**
   The tool writes new paragraph styles and tags each paragraph with its assigned style.
   Only the style tag is added. Text, spacing, numbering, and layout remain unchanged.

6. **Capture the Formatting Environment**
   The tool snapshots additional document settings including font defaults, theme colors, compatibility flags, page layout, headers/footers, and numbering definitions.

7. **Validate and Write Output**
   Before writing anything, the tool verifies required roles, full classification coverage, and byte-for-byte integrity for protected document components.
   If any check fails, the run aborts.

# The Two Output Files

- `arch_style_registry.json`: Maps each CSI role to the Word style that represents it.
- `arch_template_registry.json`: Captures the template's broader formatting environment for downstream rendering consistency.

# What This Tool Does Not Do

- It does not change how the architect's template looks.
- It does not generate new content.
- It does not modify headers, footers, or page layout.
- It does not produce a new `.docx` file.
- It does not touch numbering definitions.
"""


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    App().mainloop()


if __name__ == "__main__":
    main()
