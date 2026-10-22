"""The Why Trust It? topic (Level 1) and the dossier it opens (Level 2).

Purpose: draw the app's two trust surfaces. The words live in
``spec_formatter/trust.py``; this module only lays them out, with one small
drawing primitive per block type (section, paragraph, bullets, key points,
table, two-column contrast, note, exact claim, links, diagram, runtime card)
so every section looks alike.

Reader: a working engineer who puts their name on the output.

If the implementation changes, this changes with it. A trust document that
has drifted from the code is worse than none.

Behaviour the tests hold:

- Level 2 opens on top of Level 1 and never replaces it; closing Level 2
  returns you to Level 1, with keyboard focus back on the button that opened
  it. Closing Level 1 returns focus to the main window's Why Trust It? button.
- One Escape closes exactly one window, the topmost.
- The open dialog holds the input grab, so keyboard and mouse stay inside it.
- Nothing here loads an image, a font file or anything from the network. The
  diagram is drawn on a Tk canvas. The only addresses that can be opened are
  the further-reading links, and only when you click one.
"""

from __future__ import annotations

import tkinter as tk
import webbrowser
from typing import Callable, Mapping, Optional

import customtkinter as ctk

from spec_formatter import trust

FontFactory = Callable[..., ctk.CTkFont]

#: Below this content width the contents rail is hidden and the dossier is
#: one column.
RAIL_BREAKPOINT = 900
RAIL_WIDTH = 230
#: Level 2 takes at most this share of the screen's height.
DOSSIER_HEIGHT_SHARE = 0.88
DOSSIER_WIDTH_SHARE = 0.9
DOSSIER_MAX_WIDTH = 1180


def _focus_target(widget: tk.Misc) -> tk.Misc:
    """The inner Tk widget that really takes focus for a CTk button or label.

    CustomTkinter draws a button or label on a canvas with a plain Tk label
    on top, and forwards ``focus_set`` to that label, so keyboard bindings
    and Tab traversal have to live there.
    """

    for attribute in ("_text_label", "_label"):
        inner = getattr(widget, attribute, None)
        if isinstance(inner, tk.Misc):
            return inner
    return widget


def _make_keyboard_button(
    button: ctk.CTkButton, command: Callable[[], None], colors: Mapping[str, str]
) -> ctk.CTkButton:
    """Let Tab reach *button*, Return/space press it, and show its focus."""

    idle_border = button.cget("border_width")
    idle_color = button.cget("border_color")
    target = _focus_target(button)
    target.configure(takefocus=1)

    def press(_event: object = None) -> str:
        command()
        return "break"

    def focus_in(_event: object = None) -> None:
        try:
            button.configure(border_width=2, border_color=colors["text"])
        except tk.TclError:
            pass

    def focus_out(_event: object = None) -> None:
        try:
            button.configure(border_width=idle_border, border_color=idle_color)
        except tk.TclError:
            pass

    for sequence in ("<Return>", "<KP_Enter>", "<space>"):
        target.bind(sequence, press)
    target.bind("<FocusIn>", focus_in, add="+")
    target.bind("<FocusOut>", focus_out, add="+")
    return button


def _scaling(widget: tk.Misc, kind: str) -> float:
    """CustomTkinter's display-scaling factor for *widget* (1.0 if unknown).

    CTk multiplies window sizes (``kind="window"``) and widget sizes, wrap
    lengths included (``kind="widget"``), by the Windows display-scaling
    factor, while Tk reports screen and resize sizes in real pixels. Every
    real-pixel size handed back to CTk is divided by this first.
    """

    getter = getattr(ctk.ScalingTracker, f"get_{kind}_scaling", None)
    try:
        value = float(getter(widget)) if callable(getter) else 1.0
    except Exception:  # an unregistered window, or a CTk without the tracker
        value = 1.0
    return value if value > 0 else 1.0


def _give_focus(widget: Optional[tk.Misc]) -> None:
    """Move keyboard focus to *widget*, across windows if need be."""

    if widget is None:
        return
    try:
        if widget.winfo_exists():
            _focus_target(widget).focus_force()
    except tk.TclError:
        pass


class _TrustDialog(ctk.CTkToplevel):
    """A modal dialog: holds the grab, closes on one Escape, restores focus."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        title: str,
        colors: Mapping[str, str],
        font: FontFactory,
        on_closed: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(master)
        self.colors = colors
        self.font = font
        self._on_closed = on_closed
        self._closed = False
        self.title(title)
        self.configure(fg_color=colors["bg"])
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", self._on_escape)
        # CTkToplevel finishes mapping itself a moment later on Windows; a
        # grab on an unmapped window fails, so it is taken once it shows.
        self.after(150, self.take_grab)

    def _on_escape(self, _event: object = None) -> str:
        self.close()
        return "break"  # one Escape closes one window

    def take_grab(self) -> None:
        if self._closed:
            return
        try:
            if self.winfo_exists():
                self.grab_set()
        except tk.TclError:
            self.after(100, self.take_grab)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        if self._on_closed is not None:
            self._on_closed()

    @property
    def is_open(self) -> bool:
        try:
            return not self._closed and bool(self.winfo_exists())
        except tk.TclError:
            return False


# ---------------------------------------------------------------------------
# Level 1
# ---------------------------------------------------------------------------


class WhyTrustItDialog(_TrustDialog):
    """Level 1: the lead, the claims with their mechanisms, and the way in."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        colors: Mapping[str, str],
        font: FontFactory,
        opener: Optional[tk.Misc] = None,
    ) -> None:
        self._opener = opener
        super().__init__(
            master,
            title=trust.LEVEL1_TITLE,
            colors=colors,
            font=font,
            on_closed=self._return_focus_to_opener,
        )
        self.geometry("680x640")
        self.minsize(520, 460)
        self.dossier: Optional[DossierDialog] = None

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=20)

        close = ctk.CTkButton(
            body,
            text="Close",
            command=self.close,
            width=100,
            height=34,
            fg_color=colors["input"],
            hover_color=colors["border"],
            border_width=1,
            border_color=colors["border"],
            font=font(14),
        )
        close.pack(side="bottom", anchor="e", pady=(12, 0))
        _make_keyboard_button(close, self.close, colors)

        ctk.CTkLabel(
            body,
            text=trust.LEVEL1_TITLE,
            text_color=colors["text"],
            font=font(22, "bold"),
            anchor="w",
        ).pack(anchor="w")

        scroll = ctk.CTkScrollableFrame(body, fg_color="transparent")
        scroll.pack(fill="both", expand=True, pady=(10, 0))
        self._wrapped: list[ctk.CTkLabel] = []

        self._wrap(
            ctk.CTkLabel(
                scroll,
                text=trust.LEVEL1_LEAD,
                text_color=colors["text"],
                font=font(15, "bold"),
                anchor="w",
                justify="left",
            )
        ).pack(anchor="w", fill="x", pady=(0, 8))

        self.point_titles: list[str] = []
        for point in trust.LEVEL1_POINTS:
            ctk.CTkLabel(
                scroll,
                text=point.title,
                text_color=colors["accent"],
                font=font(14, "bold"),
                anchor="w",
                justify="left",
            ).pack(anchor="w", pady=(10, 2))
            self._wrap(
                ctk.CTkLabel(
                    scroll,
                    text=point.body,
                    text_color=colors["secondary"],
                    font=font(13),
                    anchor="w",
                    justify="left",
                )
            ).pack(anchor="w", fill="x")
            self.point_titles.append(point.title)

        box = ctk.CTkFrame(
            scroll,
            fg_color=colors["card"],
            border_width=1,
            border_color=colors["border"],
            corner_radius=10,
        )
        box.pack(fill="x", pady=(18, 6))
        ctk.CTkLabel(
            box,
            text=trust.LEVEL1_CLOSING_HEADING,
            text_color=colors["text"],
            font=font(15, "bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(14, 2))
        self._wrap(
            ctk.CTkLabel(
                box,
                text=trust.LEVEL1_CLOSING_TEXT,
                text_color=colors["secondary"],
                font=font(13),
                anchor="w",
                justify="left",
            ),
            inset=40,
        ).pack(anchor="w", fill="x", padx=16)
        self.dossier_button = ctk.CTkButton(
            box,
            text=trust.LEVEL1_BUTTON,
            command=self.open_dossier,
            height=38,
            fg_color=colors["accent"],
            hover_color=colors["accent_hover"],
            border_width=0,
            border_color=colors["accent"],
            text_color="#FFFFFF",
            font=font(14, "bold"),
        )
        self.dossier_button.pack(anchor="w", padx=16, pady=(12, 16))
        _make_keyboard_button(self.dossier_button, self.open_dossier, colors)

        tk.Frame.bind(body, "<Configure>", self._rewrap, add="+")
        self.after(
            50, lambda: self._rewrap_to(int(body.winfo_width() / _scaling(self, "widget")))
        )
        self.after(200, lambda: _give_focus(self.dossier_button))

    def _wrap(self, label: ctk.CTkLabel, inset: int = 24) -> ctk.CTkLabel:
        label._trust_inset = inset  # type: ignore[attr-defined]
        label.configure(wraplength=600 - inset)
        self._wrapped.append(label)
        return label

    def _rewrap(self, event: tk.Event) -> None:
        self._rewrap_to(int(event.width / _scaling(self, "widget")))

    def _rewrap_to(self, body_width: int) -> None:
        if body_width <= 1:
            return
        width = max(body_width - 30, 240)
        for label in self._wrapped:
            try:
                label.configure(wraplength=width - label._trust_inset)  # type: ignore[attr-defined]
            except tk.TclError:
                pass

    def open_dossier(self) -> "DossierDialog":
        """Open Level 2 on top of this window, or bring it forward."""

        if self.dossier is not None and self.dossier.is_open:
            self.dossier.lift()
            _give_focus(self.dossier.close_button)
            return self.dossier
        self.dossier = DossierDialog(
            self,
            colors=self.colors,
            font=self.font,
            on_closed=self._dossier_closed,
        )
        return self.dossier

    def _dossier_closed(self) -> None:
        self.dossier = None
        if not self.is_open:
            return
        # Level 1 takes the grab back, and focus returns to the button that
        # opened Level 2.
        self.take_grab()
        _give_focus(self.dossier_button)

    def _return_focus_to_opener(self) -> None:
        _give_focus(self._opener)


# ---------------------------------------------------------------------------
# Level 2
# ---------------------------------------------------------------------------


class DossierDialog(_TrustDialog):
    """Level 2: every section in one scroll, with a contents rail beside it."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        colors: Mapping[str, str],
        font: FontFactory,
        on_closed: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(
            master, title=trust.DOSSIER_TITLE, colors=colors, font=font, on_closed=on_closed
        )
        # Screen sizes are real pixels; CTk scales the sizes given to it.
        window_scale = _scaling(self, "window")
        screen_w = int(self.winfo_screenwidth() / window_scale)
        screen_h = int(self.winfo_screenheight() / window_scale)
        width = min(DOSSIER_MAX_WIDTH, int(screen_w * DOSSIER_WIDTH_SHARE))
        height = int(screen_h * DOSSIER_HEIGHT_SHARE)
        self.geometry(f"{width}x{height}")
        self.minsize(min(560, width), min(420, height))
        self.maxsize(screen_w, height)

        #: Widgets the tests and the rail inspect.
        self.section_frames: dict[str, ctk.CTkFrame] = {}
        self.rail_buttons: dict[str, ctk.CTkButton] = {}
        self.card_rows: list[tuple[str, tuple[str, ...]]] = []
        self.link_labels: list[ctk.CTkLabel] = []
        self.diagram: Optional[tk.Canvas] = None
        self._current_anchor: Optional[str] = None
        # (label, share of content width, inset) re-wrapped on every resize.
        self._wrapped: list[tuple[ctk.CTkLabel, float, int]] = []
        self._content_width = width - RAIL_WIDTH - 80

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(18, 8))
        close = ctk.CTkButton(
            header,
            text="Close",
            command=self.close,
            width=100,
            height=34,
            fg_color=colors["input"],
            hover_color=colors["border"],
            border_width=1,
            border_color=colors["border"],
            font=font(14),
        )
        close.pack(side="right", anchor="n")
        _make_keyboard_button(close, self.close, colors)
        self.close_button = close
        titles = ctk.CTkFrame(header, fg_color="transparent")
        titles.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            titles,
            text=trust.DOSSIER_TITLE,
            text_color=colors["text"],
            font=font(24, "bold"),
            anchor="w",
        ).pack(anchor="w")
        self._label(
            titles, trust.DOSSIER_SUBTITLE, size=13, color="secondary", share=1.0, inset=150
        ).pack(anchor="w", fill="x", pady=(2, 0))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=(0, 18))
        self._body = body
        tk.Frame.bind(body, "<Configure>", self._on_resize, add="+")

        self.rail = ctk.CTkFrame(
            body,
            width=RAIL_WIDTH,
            fg_color=colors["card"],
            border_width=1,
            border_color=colors["border"],
            corner_radius=10,
        )
        self.rail.pack_propagate(False)
        ctk.CTkLabel(
            self.rail,
            text="CONTENTS",
            text_color=colors["muted"],
            font=font(11, "bold"),
            anchor="w",
        ).pack(anchor="w", padx=14, pady=(12, 6))
        for section in trust.DOSSIER_SECTIONS:
            button = ctk.CTkButton(
                self.rail,
                text=section.rail_label,
                command=lambda anchor=section.anchor: self.scroll_to(anchor),
                anchor="w",
                height=30,
                fg_color="transparent",
                hover_color=colors["input"],
                border_width=0,
                border_color=colors["card"],
                text_color=colors["secondary"],
                font=font(12),
            )
            button.pack(fill="x", padx=6, pady=1)
            _make_keyboard_button(
                button, lambda anchor=section.anchor: self.scroll_to(anchor), colors
            )
            self.rail_buttons[section.anchor] = button
        self._rail_visible = False

        self.scroll = ctk.CTkScrollableFrame(body, fg_color="transparent")
        self.scroll.pack(side="right", fill="both", expand=True)
        canvas = self.scroll._parent_canvas
        scrollbar = self.scroll._scrollbar
        # Follow the scroll to highlight the current section, without polling.
        canvas.configure(
            yscrollcommand=lambda first, last: (
                scrollbar.set(first, last),
                self._update_current_section(),
            )
        )

        for section in trust.DOSSIER_SECTIONS:
            self._draw_section(section)

        for sequence, amount, unit in (
            ("<Down>", 3, "units"),
            ("<Up>", -3, "units"),
            ("<Next>", 1, "pages"),
            ("<Prior>", -1, "pages"),
        ):
            self.bind(sequence, lambda _e, a=amount, u=unit: self._scroll(a, u))
        self.bind("<Home>", lambda _e: self._moveto(0.0))
        self.bind("<End>", lambda _e: self._moveto(1.0))
        # Building the widgets can lay the window out before the binding
        # above sees a change of size, so apply the current width once too.
        self.after(
            50,
            lambda: self._apply_width(int(body.winfo_width() / _scaling(self, "widget"))),
        )
        self.after(200, self._initial_focus)

    # -- behaviour --------------------------------------------------------

    def _initial_focus(self) -> None:
        if not self.is_open:
            return
        first = self.rail_buttons[trust.DOSSIER_SECTIONS[0].anchor]
        _give_focus(first if self._rail_visible else self.close_button)

    def _scroll(self, amount: int, unit: str) -> str:
        self.scroll._parent_canvas.yview_scroll(amount, unit)
        return "break"

    def _moveto(self, fraction: float) -> str:
        self.scroll._parent_canvas.yview_moveto(fraction)
        return "break"

    def scroll_to(self, anchor: str) -> None:
        """Scroll so the section with *anchor* sits at the top."""

        frame = self.section_frames[anchor]
        self.update_idletasks()
        total = max(self.scroll.winfo_height(), 1)
        self.scroll._parent_canvas.yview_moveto(frame.winfo_y() / total)
        self._update_current_section()

    def current_section(self) -> str:
        canvas = self.scroll._parent_canvas
        total = max(self.scroll.winfo_height(), 1)
        top = canvas.yview()[0] * total + 24
        current = trust.DOSSIER_SECTIONS[0].anchor
        for section in trust.DOSSIER_SECTIONS:
            if self.section_frames[section.anchor].winfo_y() <= top:
                current = section.anchor
        if canvas.yview()[1] >= 0.999:
            current = trust.DOSSIER_SECTIONS[-1].anchor
        return current

    def _update_current_section(self) -> None:
        try:
            anchor = self.current_section()
        except tk.TclError:
            return
        if anchor == self._current_anchor:
            return
        self._current_anchor = anchor
        for key, button in self.rail_buttons.items():
            active = key == anchor
            button.configure(
                fg_color=self.colors["input"] if active else "transparent",
                text_color=self.colors["text"] if active else self.colors["secondary"],
            )

    def _on_resize(self, event: tk.Event) -> None:
        self._apply_width(int(event.width / _scaling(self, "widget")))

    def _apply_width(self, width: int) -> None:
        # Layout can start while the widgets are still being built.
        if width <= 1 or not self.is_open or getattr(self, "scroll", None) is None:
            return
        wants_rail = width >= RAIL_BREAKPOINT
        if wants_rail and not self._rail_visible:
            self.rail.pack(side="left", fill="y", padx=(0, 16), before=self.scroll._parent_frame)
            self._rail_visible = True
        elif not wants_rail and self._rail_visible:
            self.rail.pack_forget()
            self._rail_visible = False
        content = width - (RAIL_WIDTH + 16 if self._rail_visible else 0) - 40
        content = max(content, 300)
        if abs(content - self._content_width) < 4:
            return
        self._content_width = content
        for label, share, inset in self._wrapped:
            try:
                label.configure(wraplength=max(int(content * share) - inset, 120))
            except tk.TclError:
                pass
        if self.diagram is not None:
            self._draw_diagram_shapes(content)

    # -- primitives -------------------------------------------------------

    def _label(
        self,
        parent: tk.Misc,
        text: str,
        *,
        size: int = 13,
        weight: str = "normal",
        color: str = "secondary",
        share: float = 1.0,
        inset: int = 0,
    ) -> ctk.CTkLabel:
        label = ctk.CTkLabel(
            parent,
            text=text,
            text_color=self.colors[color],
            font=self.font(size, weight),
            anchor="w",
            justify="left",
            wraplength=max(int(self._content_width * share) - inset, 120),
        )
        self._wrapped.append((label, share, inset))
        return label

    def _draw_section(self, section: trust.Section) -> None:
        frame = ctk.CTkFrame(self.scroll, fg_color="transparent")
        frame.pack(fill="x", pady=(6, 22))
        self.section_frames[section.anchor] = frame
        ctk.CTkLabel(
            frame,
            text=section.kicker.upper(),
            text_color=self.colors["accent"],
            font=self.font(11, "bold"),
            anchor="w",
        ).pack(anchor="w")
        self._label(frame, section.title, size=20, weight="bold", color="text").pack(
            anchor="w", fill="x", pady=(0, 8)
        )
        for block in section.blocks:
            self._draw_block(frame, block)

    def _draw_block(self, parent: tk.Misc, block: trust.Block) -> None:
        if isinstance(block, trust.Paragraph):
            self._paragraph(parent, block)
        elif isinstance(block, trust.Bullets):
            self._bullets(parent, block)
        elif isinstance(block, trust.KeyPoints):
            self._key_points(parent, block)
        elif isinstance(block, trust.Table):
            self._table(parent, block)
        elif isinstance(block, trust.Contrast):
            self._contrast(parent, block)
        elif isinstance(block, trust.Note):
            self._note(parent, block)
        elif isinstance(block, trust.ExactClaim):
            self._exact_claim(parent, block)
        elif isinstance(block, trust.Links):
            self._links(parent, block)
        elif isinstance(block, trust.DataFlowDiagram):
            self._diagram(parent, block)
        elif isinstance(block, trust.RuntimeCards):
            for number, card in enumerate(block.cards, 1):
                self._runtime_card(parent, number, card)
        else:  # pragma: no cover - every block type is listed above
            raise TypeError(type(block).__name__)

    def _paragraph(self, parent: tk.Misc, block: trust.Paragraph) -> None:
        self._label(parent, block.text).pack(anchor="w", fill="x", pady=(0, 8))

    def _bullets(self, parent: tk.Misc, block: trust.Bullets) -> None:
        for number, item in enumerate(block.items, 1):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=2)
            marker = f"{number}." if block.numbered else "•"
            ctk.CTkLabel(
                row,
                text=marker,
                width=26,
                anchor="ne",
                text_color=self.colors["accent"],
                font=self.font(13, "bold"),
            ).pack(side="left", anchor="n", padx=(0, 8))
            self._label(row, item, inset=40).pack(side="left", fill="x", expand=True)

    def _key_points(self, parent: tk.Misc, block: trust.KeyPoints) -> None:
        for lead, text in block.items:
            self._label(parent, lead, size=14, weight="bold", color="text").pack(
                anchor="w", fill="x", pady=(8, 0)
            )
            self._label(parent, text).pack(anchor="w", fill="x")

    def _boxed(self, parent: tk.Misc, *, accent: bool = False) -> ctk.CTkFrame:
        box = ctk.CTkFrame(
            parent,
            fg_color=self.colors["card"],
            border_width=1,
            border_color=self.colors["accent"] if accent else self.colors["border"],
            corner_radius=10,
        )
        box.pack(fill="x", pady=(6, 10))
        return box

    def _table(self, parent: tk.Misc, block: trust.Table) -> None:
        box = self._boxed(parent)
        grid = ctk.CTkFrame(box, fg_color="transparent")
        grid.pack(fill="x", padx=12, pady=10)
        count = len(block.columns)
        share = 1.0 / count
        for column in range(count):
            grid.grid_columnconfigure(column, weight=1, uniform="col")
        for column, heading in enumerate(block.columns):
            self._label(
                grid, heading, size=12, weight="bold", color="text", share=share, inset=34
            ).grid(row=0, column=column, sticky="nw", padx=6, pady=(0, 6))
        for row_number, row in enumerate(block.rows, 1):
            ctk.CTkFrame(grid, height=1, fg_color=self.colors["border"]).grid(
                row=2 * row_number - 1, column=0, columnspan=count, sticky="ew", pady=4
            )
            for column, cell in enumerate(row):
                self._label(
                    grid,
                    cell,
                    size=12,
                    weight="bold" if column == 0 else "normal",
                    color="text" if column == 0 else "secondary",
                    share=share,
                    inset=34,
                ).grid(row=2 * row_number, column=column, sticky="nw", padx=6)

    def _contrast(self, parent: tk.Misc, block: trust.Contrast) -> None:
        pair = ctk.CTkFrame(parent, fg_color="transparent")
        pair.pack(fill="x", pady=(6, 10))
        pair.grid_columnconfigure((0, 1), weight=1, uniform="half")
        for column, (title, items) in enumerate(
            ((block.left_title, block.left_items), (block.right_title, block.right_items))
        ):
            box = ctk.CTkFrame(
                pair,
                fg_color=self.colors["card"],
                border_width=1,
                border_color=self.colors["warning" if column == 0 else "success"],
                corner_radius=10,
            )
            box.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 0))
            ctk.CTkLabel(
                box,
                text=title,
                text_color=self.colors["warning" if column == 0 else "success"],
                font=self.font(14, "bold"),
                anchor="w",
            ).pack(anchor="w", padx=14, pady=(10, 4))
            for item in items:
                self._label(box, f"•  {item}", size=12, share=0.5, inset=44).pack(
                    anchor="w", fill="x", padx=14, pady=(0, 6)
                )

    def _note(self, parent: tk.Misc, block: trust.Note) -> None:
        box = self._boxed(parent, accent=True)
        self._label(box, block.title, size=14, weight="bold", color="text", inset=32).pack(
            anchor="w", fill="x", padx=16, pady=(12, 2)
        )
        self._label(box, block.text, inset=32).pack(anchor="w", fill="x", padx=16, pady=(0, 12))

    def _exact_claim(self, parent: tk.Misc, block: trust.ExactClaim) -> None:
        self._label(parent, block.term, size=15, weight="bold", color="text").pack(
            anchor="w", fill="x", pady=(10, 4)
        )
        box = self._boxed(parent, accent=True)
        self._label(box, "The exact claim", size=11, weight="bold", color="accent", inset=32).pack(
            anchor="w", fill="x", padx=16, pady=(10, 0)
        )
        self._label(box, block.claim, color="text", inset=32).pack(
            anchor="w", fill="x", padx=16, pady=(0, 10)
        )
        self._label(parent, f"What it does not claim: {block.not_claimed}").pack(
            anchor="w", fill="x", pady=(0, 6)
        )

    def _links(self, parent: tk.Misc, block: trust.Links) -> None:
        box = self._boxed(parent)
        ctk.CTkLabel(
            box,
            text=block.title,
            text_color=self.colors["text"],
            font=self.font(14, "bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 4))
        for label_text, url in block.items:
            link = ctk.CTkLabel(
                box,
                text=f"{label_text} ↗",
                text_color=self.colors["accent"],
                font=self.font(13),
                anchor="w",
                cursor="hand2",
            )
            link.pack(anchor="w", padx=16, pady=1)
            # The further-reading links are the only addresses this window
            # can open, and only on a click (or Return once focused).
            link.bind("<Button-1>", lambda _event, target=url: webbrowser.open(target))
            inner = _focus_target(link)
            inner.configure(takefocus=1)
            inner.bind("<Return>", lambda _event, target=url: webbrowser.open(target))
            self.link_labels.append(link)
        ctk.CTkFrame(box, height=8, fg_color="transparent").pack()

    def _diagram(self, parent: tk.Misc, block: trust.DataFlowDiagram) -> None:
        box = self._boxed(parent)
        canvas = tk.Canvas(
            box,
            height=int(300 * _scaling(self, "widget")),
            bg=self.colors["card"],
            highlightthickness=0,
            takefocus=0,
        )
        canvas.pack(fill="x", padx=12, pady=(12, 4))
        self.diagram = canvas
        self._diagram_block = block
        # The description restates the drawing in words, for every reader.
        self._label(box, block.description, size=12, color="muted", inset=32).pack(
            anchor="w", fill="x", padx=16, pady=(4, 12)
        )
        self._draw_diagram_shapes(self._content_width)

    def _draw_diagram_shapes(self, content_width: int) -> None:
        canvas = self.diagram
        block = self._diagram_block
        if canvas is None:
            return
        colors = self.colors
        # A plain Tk canvas works in real pixels; layout below is in CTk's
        # logical units, so every coordinate goes through ``px``.
        scale = _scaling(self, "widget")

        def px(value: float) -> int:
            return int(round(value * scale))

        width = max(content_width - 30, 480)
        canvas.configure(width=px(width), height=px(300))
        canvas.delete("all")
        family = self.font(12).cget("family")
        font = (family, 10)
        bold = (family, 11, "bold")
        # Your computer, on the left.
        left_w = int(width * 0.34)
        canvas.create_rectangle(px(4), px(4), px(left_w), px(296), outline=colors["success"], width=2)
        canvas.create_text(px(16), px(20), text=block.machine_label, anchor="w",
                           fill=colors["success"], font=bold)
        for index, item in enumerate(block.machine_items):
            y = 58 + index * 48
            canvas.create_rectangle(px(16), px(y - 16), px(left_w - 12), px(y + 16),
                                    outline=colors["border"], fill=colors["input"])
            canvas.create_text(px(26), px(y), text=item, anchor="w", fill=colors["text"],
                               font=font, width=px(left_w - 48))
        # Destinations, on the right.
        right_x = int(width * 0.70)
        for index, (label, arrow_text, dashed) in enumerate(block.destinations):
            y = 50 + index * 100
            color = colors["muted"] if dashed else colors["warning"]
            canvas.create_rectangle(px(right_x), px(y - 26), px(width - 4), px(y + 26),
                                    outline=color, width=2)
            canvas.create_text(px(right_x + 10), px(y), text=label, anchor="w",
                               fill=colors["text"], font=bold, width=px(width - right_x - 20))
            canvas.create_line(
                px(left_w + 4), px(y), px(right_x - 4), px(y),
                fill=color, width=2, arrow="last",
                dash=(6, 4) if dashed else None,
            )
            canvas.create_text(
                px((left_w + right_x) / 2), px(y - 10), text=arrow_text, anchor="s",
                fill=colors["secondary"], font=font, width=px(right_x - left_w - 16),
                justify="center",
            )
        canvas.create_text(
            px(left_w + 8), px(292), anchor="sw", fill=colors["muted"], font=font,
            text=block.legend,
        )

    def _runtime_card(self, parent: tk.Misc, number: int, card: trust.RuntimeCard) -> None:
        box = self._boxed(parent)
        heading = ctk.CTkFrame(box, fg_color="transparent")
        heading.pack(fill="x", padx=14, pady=(10, 6))
        ctk.CTkLabel(
            heading,
            text=f"{number:02d}",
            text_color=self.colors["accent"],
            font=self.font(13, "bold"),
            width=30,
            anchor="w",
        ).pack(side="left")
        self._label(heading, card.title, size=14, weight="bold", color="text", inset=160).pack(
            side="left", fill="x", expand=True
        )
        if card.automatic:
            ctk.CTkLabel(
                heading,
                text="AUTOMATIC",
                text_color=self.colors["warning"],
                font=self.font(10, "bold"),
            ).pack(side="right")
        grid = ctk.CTkFrame(box, fg_color="transparent")
        grid.pack(fill="x", padx=14, pady=(0, 10))
        grid.grid_columnconfigure(1, weight=1)
        labels: list[str] = []
        for row, (label, value) in enumerate(card.rows()):
            ctk.CTkLabel(
                grid,
                text=label,
                text_color=self.colors["muted"],
                font=self.font(12, "bold"),
                width=110,
                anchor="nw",
                justify="left",
            ).grid(row=row, column=0, sticky="nw", pady=2)
            emphasis = label == "AI involved"
            self._label(
                grid,
                value,
                size=12,
                weight="bold" if emphasis else "normal",
                color="text" if emphasis else "secondary",
                inset=170,
            ).grid(row=row, column=1, sticky="nw", pady=2)
            labels.append(label)
        self.card_rows.append((card.title, tuple(labels)))


def open_why_trust_it(
    master: tk.Misc,
    *,
    colors: Mapping[str, str],
    font: FontFactory,
    opener: Optional[tk.Misc] = None,
    existing: Optional[WhyTrustItDialog] = None,
) -> WhyTrustItDialog:
    """Open Level 1, or bring an open one forward."""

    if existing is not None and existing.is_open:
        existing.lift()
        _give_focus(existing.dossier_button)
        return existing
    return WhyTrustItDialog(master, colors=colors, font=font, opener=opener)
