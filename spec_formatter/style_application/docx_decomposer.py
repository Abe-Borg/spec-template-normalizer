"""
Word Document Decomposer

Extracts DOCX files into their constituent XML parts for processing
by the Phase 2 MEP Specification Styling Engine.
"""

from __future__ import annotations

import re
import shutil
import stat
import zipfile
from pathlib import Path
from typing import Set


MAX_PACKAGE_ENTRIES = 10_000
MAX_PACKAGE_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_PACKAGE_PART_BYTES = 128 * 1024 * 1024
MAX_COMPRESSION_RATIO = 1_000

REQUIRED_PACKAGE_PARTS = (
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/styles.xml",
)


def _safe_package_member_path(extract_dir: Path, member_name: str) -> Path:
    """Resolve a ZIP member below *extract_dir* or reject its name."""
    if not member_name or "\x00" in member_name or "\\" in member_name:
        raise ValueError(f"Unsafe DOCX package member name: {member_name!r}")
    if member_name.startswith(("/", "//")) or re.match(r"^[A-Za-z]:", member_name):
        raise ValueError(f"Unsafe absolute DOCX package member: {member_name!r}")

    parts = member_name.rstrip("/").split("/")
    if any(part in {"", ".", ".."} for part in parts):
        raise ValueError(f"Unsafe DOCX package member traversal: {member_name!r}")

    # Keep extraction deterministic and safe on Windows even if a package was
    # created on a platform that permits these names.
    reserved_windows_names = {"CON", "PRN", "AUX", "NUL"} | {
        f"{prefix}{number}"
        for prefix in ("COM", "LPT")
        for number in range(1, 10)
    }
    for part in parts:
        stem = part.split(".", 1)[0].upper()
        if ":" in part or part.endswith((" ", ".")) or stem in reserved_windows_names:
            raise ValueError(f"Unsafe Windows DOCX package member: {member_name!r}")

    destination = (extract_dir / Path(*parts)).resolve()
    root = extract_dir.resolve()
    try:
        destination.relative_to(root)
    except ValueError as exc:
        raise ValueError(
            f"DOCX package member escapes extraction root: {member_name!r}"
        ) from exc
    return destination


class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.

        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None

    def extract(self, output_dir=None):
        """Extract the DOCX into a new, bounded workspace directory.

        The extraction target must not exist. Any validation or extraction
        failure removes the partial tree before the exception is propagated.
        """
        if output_dir is None:
            output_dir = Path(f"{self.docx_path.stem}_extracted")
        else:
            output_dir = Path(output_dir)

        if output_dir.exists():
            raise FileExistsError(f"Extraction target already exists: {output_dir}")

        output_dir.mkdir(parents=True, exist_ok=False)
        try:
            print(f"Extracting {self.docx_path} to {output_dir}...")
            with zipfile.ZipFile(self.docx_path, "r") as archive:
                entries = archive.infolist()
                if len(entries) > MAX_PACKAGE_ENTRIES:
                    raise ValueError(
                        f"DOCX package has {len(entries)} entries; "
                        f"limit is {MAX_PACKAGE_ENTRIES}"
                    )

                total_size = sum(entry.file_size for entry in entries)
                if total_size > MAX_PACKAGE_UNCOMPRESSED_BYTES:
                    raise ValueError(
                        f"DOCX package expands to {total_size} bytes; "
                        f"limit is {MAX_PACKAGE_UNCOMPRESSED_BYTES}"
                    )

                seen_names: Set[str] = set()
                for entry in entries:
                    # orig_filename preserves crafted backslashes that
                    # ZipInfo.filename may normalize or truncate.
                    _safe_package_member_path(output_dir, entry.orig_filename)
                    destination = _safe_package_member_path(output_dir, entry.filename)

                    normalized_name = entry.filename.casefold()
                    if normalized_name in seen_names:
                        raise ValueError(
                            f"DOCX package contains duplicate member: {entry.filename!r}"
                        )
                    seen_names.add(normalized_name)

                    unix_mode = (entry.external_attr >> 16) & 0xFFFF
                    if stat.S_ISLNK(unix_mode):
                        raise ValueError(
                            f"DOCX package contains a symbolic link: {entry.filename!r}"
                        )
                    if entry.file_size > MAX_PACKAGE_PART_BYTES:
                        raise ValueError(
                            f"DOCX package member {entry.filename!r} is "
                            f"{entry.file_size} bytes; per-part limit is "
                            f"{MAX_PACKAGE_PART_BYTES}"
                        )
                    if entry.file_size and entry.compress_size == 0:
                        raise ValueError(
                            "DOCX package member has invalid compressed size: "
                            f"{entry.filename!r}"
                        )
                    if (
                        entry.compress_size
                        and entry.file_size / entry.compress_size > MAX_COMPRESSION_RATIO
                    ):
                        raise ValueError(
                            "DOCX package member has suspicious compression ratio: "
                            f"{entry.filename!r}"
                        )

                    if entry.is_dir():
                        destination.mkdir(parents=True, exist_ok=True)
                        continue
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(entry, "r") as source, destination.open("xb") as target:
                        shutil.copyfileobj(source, target, length=1024 * 1024)

            missing_parts = [
                part
                for part in REQUIRED_PACKAGE_PARTS
                if not (output_dir / Path(*part.split("/"))).is_file()
            ]
            if missing_parts:
                raise ValueError(f"DOCX package is missing required parts: {missing_parts}")
        except Exception:
            shutil.rmtree(output_dir, ignore_errors=True)
            raise

        self.extract_dir = output_dir
        print(f"Extraction complete: {len(entries)} items extracted")
        return output_dir
