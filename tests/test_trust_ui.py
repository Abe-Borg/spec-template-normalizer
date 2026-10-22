"""Behaviour of the Why Trust It? dialog (Level 1) and the dossier (Level 2).

These open real windows. Like ``tests/test_gui_modes.py`` they skip where
``tkinter`` is missing, and they also skip where no display can be opened;
the Windows CI job runs them.
"""

from __future__ import annotations

import time

import pytest

ctk = pytest.importorskip("customtkinter")

import tkinter as tk  # noqa: E402

import gui  # noqa: E402
import trust_ui  # noqa: E402
from spec_formatter import trust  # noqa: E402


@pytest.fixture()
def root():
    try:
        window = ctk.CTk()
    except tk.TclError as exc:  # no display
        pytest.skip(f"no display: {exc}")
    window.geometry("900x700+0+0")
    yield window
    try:
        window.destroy()
    except tk.TclError:
        pass


def _pump(window, seconds: float = 0.5) -> None:
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        window.update()
        time.sleep(0.01)


def _require_keyboard_focus(window) -> None:
    """Skip when the desktop will not give this process keyboard focus.

    Key events reach a Tk window only while the application has focus; a
    headless CI desktop may refuse it to a background process.
    """

    window.focus_force()
    _pump(window, 0.3)
    if window.focus_get() is None:
        pytest.skip("the desktop did not give this test process keyboard focus")


def _focused_in(window, widget) -> bool:
    focus = window.focus_get()
    target = trust_ui._focus_target(widget)
    return focus is not None and str(focus) == str(target)


def _open_level1(root):
    opener = ctk.CTkButton(root, text="Why Trust It?")
    opener.pack()
    dialog = trust_ui.open_why_trust_it(
        root, colors=gui.COLORS, font=gui._font, opener=opener
    )
    _pump(root)
    return opener, dialog


def test_level1_shows_every_point_and_the_dossier_button(root):
    _opener, level1 = _open_level1(root)
    assert level1.title() == trust.LEVEL1_TITLE
    assert level1.point_titles == [point.title for point in trust.LEVEL1_POINTS]
    assert level1.dossier_button.cget("text") == trust.LEVEL1_BUTTON


def test_level1_button_opens_level2_on_top_without_closing_level1(root):
    _opener, level1 = _open_level1(root)
    level1.dossier_button.invoke()
    _pump(root, 0.8)
    level2 = level1.dossier
    assert isinstance(level2, trust_ui.DossierDialog)
    assert level2.is_open and level1.is_open
    assert str(level2.master) == str(level1)
    assert str(root.grab_current()) == str(level2)
    assert set(level2.section_frames) == {s.anchor for s in trust.DOSSIER_SECTIONS}


def test_one_escape_closes_one_dialog_and_focus_returns(root):
    _require_keyboard_focus(root)
    opener, level1 = _open_level1(root)
    level1.dossier_button.invoke()
    _pump(root, 0.8)
    level2 = level1.dossier

    root.focus_get().event_generate("<Escape>")
    _pump(root)
    assert not level2.is_open
    assert level1.is_open, "Escape on Level 2 must not close Level 1"
    assert str(root.grab_current()) == str(level1)
    assert _focused_in(root, level1.dossier_button)

    root.focus_get().event_generate("<Escape>")
    _pump(root)
    assert not level1.is_open
    assert _focused_in(root, opener)


def test_return_on_the_focused_button_opens_level2(root):
    _require_keyboard_focus(root)
    _opener, level1 = _open_level1(root)
    assert _focused_in(root, level1.dossier_button)
    root.focus_get().event_generate("<Return>")
    _pump(root, 0.8)
    assert level1.dossier is not None and level1.dossier.is_open


def test_every_runtime_card_is_drawn_with_the_five_rows(root):
    _opener, level1 = _open_level1(root)
    level2 = level1.open_dossier()
    _pump(root, 0.8)
    assert len(level2.card_rows) == len(trust.runtime_cards())
    for title, rows in level2.card_rows:
        assert rows == trust.RUNTIME_CARD_ROWS, title
    assert level2.diagram is not None
    assert level2.diagram.find_all(), "the data-flow diagram is drawn on the canvas"


def test_rail_navigates_and_follows_the_scroll(root):
    root.geometry("1400x900+0+0")
    _opener, level1 = _open_level1(root)
    level2 = level1.open_dossier()
    level2.geometry("1200x800")
    _pump(root, 0.8)
    assert level2._rail_visible
    for section in trust.DOSSIER_SECTIONS[1:]:
        level2.scroll_to(section.anchor)
        _pump(root, 0.1)
        assert level2.current_section() == section.anchor


def test_rail_hides_at_narrow_widths(root):
    _opener, level1 = _open_level1(root)
    level2 = level1.open_dossier()
    level2.geometry("620x700")
    _pump(root, 0.8)
    assert not level2._rail_visible


def test_level2_height_is_bounded_by_the_screen(root):
    _opener, level1 = _open_level1(root)
    level2 = level1.open_dossier()
    _pump(root, 0.8)
    # Real pixels on both sides, whatever the display scaling.
    limit = level2.winfo_screenheight() * trust_ui.DOSSIER_HEIGHT_SHARE
    assert level2.winfo_height() <= limit + 2


def test_sizes_handed_to_ctk_are_divided_by_display_scaling(root, monkeypatch):
    monkeypatch.setattr(trust_ui, "_scaling", lambda widget, kind: 1.25)
    _opener, level1 = _open_level1(root)
    level2 = level1.open_dossier()
    _pump(root, 0.5)
    assert trust_ui._scaling(level2, "window") == 1.25
    logical = level2.winfo_screenheight() / 1.25 * trust_ui.DOSSIER_HEIGHT_SHARE
    assert level2._current_height <= int(logical) + 1


def test_the_main_window_opens_level1_from_its_help_button(monkeypatch, root):
    calls = {}

    def fake_open(master, **kwargs):
        calls.update(kwargs, master=master)
        return "dialog"

    monkeypatch.setattr(trust_ui, "open_why_trust_it", fake_open)
    button = ctk.CTkButton(root, text="Why Trust It?")
    app = type(
        "AppStub",
        (),
        {"help_buttons": {"Why Trust It?": button}, "_trust_dialog": None},
    )()
    gui.App._show_why_trust_it_dialog(app)
    assert calls["master"] is app
    assert calls["opener"] is button
    assert app._trust_dialog == "dialog"
