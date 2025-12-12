"""Helpers for resolving image hyperlinks from optional CSV mappings."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, Tuple


def _find_mapping_file(folder: Path) -> Path | None:
    # Prefer an explicit links.csv file but allow any *links*.csv pattern.
    explicit = folder / "links.csv"
    if explicit.exists():
        return explicit

    candidates = sorted(folder.glob("*links*.csv"))
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def load_link_mapping(folder: Path) -> Tuple[Dict[str, str], str]:
    """Load a filename→URL mapping from CSV if present.

    Returns a tuple of (mapping, status_message).
    The CSV must contain at least two columns: filename,url. Extra columns are ignored.
    """

    mapping_file = _find_mapping_file(folder)
    if not mapping_file:
        return {}, "No link CSV found; image file paths will be used as hyperlinks."

    mapping: Dict[str, str] = {}
    with mapping_file.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        lower_fieldnames = [name.lower() for name in reader.fieldnames or []]
        if "filename" not in lower_fieldnames or "url" not in lower_fieldnames:
            return {}, f"{mapping_file.name} missing required 'filename' and 'url' columns; file paths will be used."

        for row in reader:
            filename = row.get("filename") or row.get("Filename") or row.get("FILENAME")
            url = row.get("url") or row.get("URL") or row.get("Url")
            if not filename or not url:
                continue
            mapping[filename] = url

    return mapping, f"Loaded link mapping from {mapping_file.name}."


def resolve_hyperlink(path: Path, mapping: Dict[str, str]) -> str:
    """Return the hyperlink for *path* using the provided mapping when available."""

    if path.name in mapping:
        return mapping[path.name]
    # Default to a file URI pointing at the local synced file.
    return path.resolve().as_uri()
