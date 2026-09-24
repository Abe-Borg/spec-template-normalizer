"""Enforce the bookkeeping rules of the DOCX Method Hardening program.

The program in ``docs/docx_method_hardening/`` runs over many independent
chat sessions. Nothing but these files carries state from one session to the
next, so this test makes the bookkeeping mandatory: a session that forgets to
record its status, tick its definition of done, or write the next session's
handoff prompt turns CI red.

The rules are stated in ``DOCX_METHOD_HARDENING_PLAN.md`` section 3.4 and at
the top of ``PROGRESS_TRACKER.md``. This test is the executable version of
them; when the two disagree, fix the documents.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PROGRAM_DIR = ROOT / "docs" / "docx_method_hardening"
PLAN = PROGRAM_DIR / "DOCX_METHOD_HARDENING_PLAN.md"
TRACKER = PROGRAM_DIR / "PROGRESS_TRACKER.md"
HANDOFFS = PROGRAM_DIR / "handoffs"

REQUIRED_STATUSES = frozenset(
    {"not_started", "in_progress", "in_review", "merged", "blocked", "dropped"}
)
OPTIONAL_STATUSES = REQUIRED_STATUSES | {"not_scheduled"}
COMPLETE_STATUSES = frozenset({"merged", "dropped"})

ITEM_COLUMNS = ("ID", "Title", "Status", "Session", "PR", "Merge commit", "Notes")
LOG_COLUMNS = ("Session", "Date (UTC)", "Work item", "Outcome", "PR", "Handoff written")

WORK_ITEM_ID_RE = re.compile(r"WI-\d{2}\Z")
SESSION_RE = re.compile(r"\d{2}\Z")
SHA_RE = re.compile(r"[0-9a-f]{7,40}\Z")
PR_URL_RE = re.compile(r"https://github\.com/[^/\s]+/[^/\s]+/pull/\d+\Z")
CHECKBOX_RE = re.compile(r"^- \[([ xX])\] ", re.M)
PLAN_ITEM_HEADING_RE = re.compile(r"^### (WI-\d{2}): (.+)$", re.M)
HANDOFF_NAME_RE = re.compile(r"handoff-for-session-(\d{2})\.md\Z")


def _read(path: Path) -> str:
    assert path.is_file(), f"missing program file: {path.relative_to(ROOT)}"
    return path.read_text(encoding="utf-8")


def _table_rows(markdown: str, heading: str, columns: tuple[str, ...]) -> list[dict[str, str]]:
    """Return the rows of the first table under ``heading`` as dicts."""

    lines = markdown.splitlines()
    try:
        start = next(i for i, line in enumerate(lines) if line.startswith(heading))
    except StopIteration:
        pytest.fail(f"PROGRESS_TRACKER.md has no section starting with {heading!r}")
    table: list[list[str]] = []
    for line in lines[start + 1 :]:
        if line.startswith("## "):
            break
        if not line.startswith("|"):
            if table:
                break
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        table.append(cells)
    assert len(table) >= 2, f"table under {heading!r} needs a header row and a separator row"
    header = tuple(table[0])
    assert header == columns, (
        f"table under {heading!r} has columns {header}, expected {columns}"
    )
    rows = []
    for cells in table[2:]:
        assert len(cells) == len(columns), (
            f"row under {heading!r} has {len(cells)} cells, expected {len(columns)}: {cells}"
        )
        rows.append(dict(zip(columns, cells)))
    return rows


def _plan_items() -> dict[str, list[bool]]:
    """Map each ``### WI-NN:`` section of the plan to its checkbox states."""

    plan = _read(PLAN)
    headings = list(PLAN_ITEM_HEADING_RE.finditer(plan))
    assert headings, "the plan has no '### WI-NN:' sections"
    items: dict[str, list[bool]] = {}
    for index, match in enumerate(headings):
        item_id = match.group(1)
        assert item_id not in items, f"the plan defines {item_id} twice"
        section_start = match.end()
        section_end = headings[index + 1].start() if index + 1 < len(headings) else len(plan)
        # A '## ' heading ends the item's section before the next '### '.
        next_h2 = re.search(r"^## ", plan[section_start:section_end], re.M)
        if next_h2:
            section_end = section_start + next_h2.start()
        section = plan[section_start:section_end]
        assert "**Definition of done**" in section, (
            f"{item_id} in the plan has no '**Definition of done**' list"
        )
        boxes = [mark.lower() == "x" for mark in CHECKBOX_RE.findall(section)]
        assert boxes, f"{item_id} in the plan has no checkboxes in its Definition of done"
        items[item_id] = boxes
    return items


def _program_status() -> str:
    match = re.search(r"^Program status: (.+)$", _read(TRACKER), re.M)
    assert match, "PROGRESS_TRACKER.md has no 'Program status:' line"
    return match.group(1).strip()


@pytest.fixture(scope="module")
def tracker_tables() -> dict[str, list[dict[str, str]]]:
    tracker = _read(TRACKER)
    return {
        "required": _table_rows(tracker, "## Required work items", ITEM_COLUMNS),
        "optional": _table_rows(tracker, "## Optional work items", ITEM_COLUMNS),
        "log": _table_rows(tracker, "## Session log", LOG_COLUMNS),
    }


def test_every_tracked_item_is_specified_in_the_plan_and_vice_versa(tracker_tables) -> None:
    tracked = [row["ID"] for row in tracker_tables["required"] + tracker_tables["optional"]]
    assert len(tracked) == len(set(tracked)), f"duplicate work item IDs in the tracker: {tracked}"
    for item_id in tracked:
        assert WORK_ITEM_ID_RE.fullmatch(item_id), f"malformed work item ID in the tracker: {item_id!r}"
    planned = set(_plan_items())
    assert set(tracked) == planned, (
        "tracker and plan disagree about the work items: "
        f"only in tracker={sorted(set(tracked) - planned)}, only in plan={sorted(planned - set(tracked))}"
    )


@pytest.mark.parametrize("table_name", ["required", "optional"])
def test_item_rows_are_internally_consistent(tracker_tables, table_name: str) -> None:
    allowed = REQUIRED_STATUSES if table_name == "required" else OPTIONAL_STATUSES
    plan_items = _plan_items()
    for row in tracker_tables[table_name]:
        item_id, status = row["ID"], row["Status"]
        assert status in allowed, (
            f"{item_id}: status {status!r} is not one of {sorted(allowed)}"
        )
        assert row["Title"], f"{item_id}: empty title"
        if status in {"in_progress", "in_review", "merged"}:
            assert SESSION_RE.fullmatch(row["Session"]), (
                f"{item_id}: status {status} needs a two-digit Session, got {row['Session']!r}"
            )
        if status in {"in_review", "merged"}:
            assert PR_URL_RE.fullmatch(row["PR"]), (
                f"{item_id}: status {status} needs a GitHub pull request URL, got {row['PR']!r}"
            )
            boxes = plan_items[item_id]
            assert all(boxes), (
                f"{item_id}: status {status} requires every Definition of done box ticked "
                f"in the plan; {boxes.count(False)} of {len(boxes)} are not"
            )
        if status == "merged":
            assert SHA_RE.fullmatch(row["Merge commit"]), (
                f"{item_id}: status merged needs the merge commit SHA, got {row['Merge commit']!r}"
            )
        if status in {"blocked", "dropped"}:
            assert row["Notes"], f"{item_id}: status {status} needs a reason in Notes"


def test_program_status_line_matches_the_required_table(tracker_tables) -> None:
    status = _program_status()
    all_done = all(row["Status"] in COMPLETE_STATUSES for row in tracker_tables["required"])
    if all_done:
        assert status == "PROGRAM COMPLETE", (
            "every required item is merged or dropped, so the 'Program status:' line must read "
            f"'PROGRAM COMPLETE' (it reads {status!r}); this is the moment to print the banner "
            "from plan section 3.5"
        )
    else:
        assert status == "IN PROGRESS", (
            f"required items are still open, so the 'Program status:' line must read "
            f"'IN PROGRESS' (it reads {status!r})"
        )


def test_every_session_wrote_the_next_handoff_prompt(tracker_tables) -> None:
    rows = tracker_tables["log"]
    assert rows, "the session log is empty; session 00 must be recorded"
    seen: set[str] = set()
    for row in rows:
        session = row["Session"]
        assert SESSION_RE.fullmatch(session), f"session log: malformed session number {session!r}"
        assert session not in seen, f"session log: session {session} appears twice"
        seen.add(session)
        assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", row["Date (UTC)"]), (
            f"session {session}: date must be YYYY-MM-DD, got {row['Date (UTC)']!r}"
        )
        assert WORK_ITEM_ID_RE.fullmatch(row["Work item"]) or row["Work item"] == "closeout", (
            f"session {session}: work item must be WI-NN or 'closeout', got {row['Work item']!r}"
        )
        assert row["Outcome"], f"session {session}: empty outcome"
        expected_name = f"handoffs/handoff-for-session-{int(session) + 1:02d}.md"
        assert row["Handoff written"] == expected_name, (
            f"session {session}: 'Handoff written' must be {expected_name!r}, "
            f"got {row['Handoff written']!r}"
        )
        handoff = PROGRAM_DIR / row["Handoff written"]
        assert handoff.is_file(), (
            f"session {session} did not write its handoff prompt: {expected_name} is missing"
        )
        first_line = handoff.read_text(encoding="utf-8").splitlines()[0].strip()
        expected_title = f"# Handoff prompt for session {int(session) + 1:02d}"
        assert first_line == expected_title, (
            f"{expected_name} must start with {expected_title!r}, got {first_line!r}"
        )


def test_handoff_directory_contains_only_well_named_prompts() -> None:
    assert HANDOFFS.is_dir(), "docs/docx_method_hardening/handoffs/ is missing"
    for path in HANDOFFS.iterdir():
        match = HANDOFF_NAME_RE.fullmatch(path.name)
        assert match, f"unexpected file in handoffs/: {path.name}"
        first_line = path.read_text(encoding="utf-8").splitlines()[0].strip()
        assert first_line == f"# Handoff prompt for session {match.group(1)}", (
            f"{path.name}: title line does not match its filename ({first_line!r})"
        )
