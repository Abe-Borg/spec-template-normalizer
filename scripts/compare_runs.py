#!/usr/bin/env python3
"""Compare target dispositions and observed tokens across formatter runs.

    python scripts/compare_runs.py RUN_HIGH RUN_MEDIUM [RUN_LOW ...]

Only run.json and its sibling target audit JSON files are read. Source paths,
DOCX files, paragraph text, ignored reasons, and architect usage are not used.
"""

from __future__ import annotations

import argparse
import itertools
import json
import re
from pathlib import Path, PureWindowsPath


TOKEN_FIELDS = (
    "input_tokens",
    "output_tokens",
    "cache_read_input_tokens",
    "cache_creation_input_tokens",
)
_HASH = re.compile(r"[0-9a-f]{64}")
_ROLE = re.compile(r"[A-Za-z0-9_.:/#@+\-]{1,160}")


def _read_object(path: Path) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object in {path}")
    return value


def _dispositions(audit: dict) -> dict[int, tuple[str, str | None]] | None:
    application = audit.get("application_audit")
    counts = audit.get("disposition_counts", {})
    if not isinstance(application, dict) or not application:
        return None
    if counts.get("unresolved", 0) != 0:
        return None
    result: dict[int, tuple[str, str | None]] = {}
    for key in ("classifications", "ignored_paragraphs"):
        items = application.get(key)
        if not isinstance(items, list):
            return None
        for item in items:
            if not isinstance(item, dict):
                return None
            index = item.get("paragraph_index")
            if type(index) is not int or index < 0 or index in result:
                return None
            role = item.get("csi_role") if key == "classifications" else None
            if key == "classifications" and (
                not isinstance(role, str) or not _ROLE.fullmatch(role)
            ):
                return None
            result[index] = ("styled", role) if role is not None else ("ignored", None)
    return result


def _usage(audit: dict) -> dict:
    # Audits retain all engine events even when diagnostics.jsonl is filtered
    # at warning/error level. Each classify event already aggregates retries
    # and chunks for this target, including requests that failed.
    events = [
        event for event in audit.get("diagnostics", [])
        if isinstance(event, dict)
        and event.get("component") == "target"
        and event.get("event") == "classify"
        and isinstance(event.get("fields"), dict)
    ]
    fields = events[0]["fields"] if len(events) == 1 else {}
    complete = fields.get("usage_complete")
    no_requests = (
        type(fields.get("requests_attempted")) is int
        and fields["requests_attempted"] == 0
        and complete is True
    )
    usage = {"usage_complete": complete if isinstance(complete, bool) else None}
    for key in TOKEN_FIELDS:
        value = fields.get(key)
        if type(value) is int and value >= 0:
            usage[key] = value
        elif no_requests or (
            complete is True and key in TOKEN_FIELDS[2:]
        ):
            usage[key] = 0
        else:
            usage[key] = None
    return usage


def _load_run(run_dir: Path) -> tuple[dict, dict]:
    manifest = _read_object(run_dir / "run.json")
    models = manifest.get("models", {})
    metadata = {
        "directory": str(run_dir),
        "model": models.get("target"),
        "effort": models.get("target_effort"),
    }
    targets = {}
    for target in manifest.get("targets", []):
        digest = target.get("source_sha256")
        if digest is None:
            continue  # Initialization may fail before a source is identified.
        if not isinstance(digest, str) or not _HASH.fullmatch(digest):
            raise ValueError(f"Invalid target source SHA-256 in {run_dir / 'run.json'}")
        if digest in targets:
            raise ValueError(f"Duplicate target source SHA-256 {digest} in {run_dir}")
        # Manifests store absolute paths. Use the co-located basename so copied
        # run directories work, including runs originally created on Windows.
        raw_path = target.get("audit_path")
        audit = {}
        if isinstance(raw_path, str) and raw_path:
            audit_path = run_dir / PureWindowsPath(raw_path).name
            if audit_path.is_file():
                audit = _read_object(audit_path)
                if audit.get("source", {}).get("sha256") != digest:
                    raise ValueError(f"Audit source SHA-256 mismatch in {audit_path}")
        targets[digest] = {
            "dispositions": _dispositions(audit),
            "usage": _usage(audit),
        }
    return metadata, targets


def compare_runs(run_dirs: list[Path]) -> dict:
    """Return pairwise agreement and per-target usage, matched only by hash."""
    if len(run_dirs) < 2:
        raise ValueError("Provide at least two run directories")
    runs = [_load_run(Path(directory)) for directory in run_dirs]
    targets = []
    for digest in sorted(set().union(*(set(records) for _, records in runs))):
        records = [records.get(digest) for _, records in runs]
        agreements = []
        for left, right in itertools.combinations(range(len(runs)), 2):
            a = records[left]["dispositions"] if records[left] else None
            b = records[right]["dispositions"] if records[right] else None
            agreement = {"runs": [left + 1, right + 1], "available": False}
            if a is not None and b is not None:
                indices = set(a) | set(b)
                different = sorted(index for index in indices if a.get(index) != b.get(index))
                agreement.update(
                    available=True,
                    agreed=len(indices) - len(different),
                    total=len(indices),
                    exact=not different,
                    differing_paragraph_indices=different,
                )
            agreements.append(agreement)
        targets.append({
            "source_sha256": digest,
            "agreement": agreements,
            "usage": [record["usage"] if record else None for record in records],
        })
    return {"runs": [metadata for metadata, _ in runs], "targets": targets}


def _print_report(report: dict) -> None:
    for index, run in enumerate(report["runs"], 1):
        print(f"Run {index}: {run['directory']} (model={run['model']}, effort={run['effort']})")
    for target in report["targets"]:
        print(f"\nTarget SHA-256: {target['source_sha256']}")
        for agreement in target["agreement"]:
            left, right = agreement["runs"]
            if not agreement["available"]:
                print(f"  Agreement {left}/{right}: unavailable (missing target or dispositions)")
                continue
            print(
                f"  Agreement {left}/{right}: {agreement['agreed']}/{agreement['total']} "
                f"paragraphs; exact={agreement['exact']}"
            )
            if agreement["differing_paragraph_indices"]:
                print(f"    Differing paragraph indices: {agreement['differing_paragraph_indices']}")
        for index, usage in enumerate(target["usage"], 1):
            if usage is None:
                print(f"  Run {index}: target missing")
                continue
            counters = ", ".join(
                f"{key}={usage[key] if usage[key] is not None else 'unknown'}"
                for key in TOKEN_FIELDS
            )
            print(f"  Run {index}: {counters}; usage_complete={usage['usage_complete']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_dirs", nargs="+", type=Path)
    parser.add_argument("--json", action="store_true", help="Print a machine-readable report")
    args = parser.parse_args()
    try:
        report = compare_runs(args.run_dirs)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    if args.json:
        print(json.dumps(report, indent=2))
    else:
        _print_report(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
