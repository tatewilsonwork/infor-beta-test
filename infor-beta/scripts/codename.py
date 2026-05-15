"""Codename resolver — Phase 1.

Implements the G3 convention from Obsidian note `12 — Locked Decisions`:

    Deal codename is the literal string the analyst types (e.g. "Project OpenText").

    - Case-preserving for display
    - Case-insensitive for lookup
    - Path-unsafe characters stripped: /, \\, :, *, ?, ", <, >, |
    - Common characters (&, ',', '.', spaces) preserved — macOS handles them fine
    - Collision-aware: caller (the conductor) decides what to do on hit; this module
      provides find_existing() and disambiguate() helpers

This module is intentionally schema-free (no pydantic import) — it is used at deal-init
before the typed contract is constructed.
"""

from __future__ import annotations

import re
from pathlib import Path

# Characters that are unsafe in a filesystem path on macOS / Linux / Windows
# (the conductor targets all three indirectly). Stripped, not replaced.
_PATH_UNSAFE = re.compile(r'[\\/:*?"<>|]')

# Default deals root per E1 (deal directory decision).
DEFAULT_DEALS_ROOT = Path("~/Documents/INFOR Deals").expanduser()


def _strip_unsafe(name: str) -> str:
    """Strip path-unsafe characters and collapse the resulting whitespace runs."""
    stripped = _PATH_UNSAFE.sub("", name)
    # Collapse any runs of whitespace caused by stripped characters.
    collapsed = re.sub(r"\s+", " ", stripped).strip()
    return collapsed


def resolve(codename: str, deals_root: Path | str = DEFAULT_DEALS_ROOT) -> tuple[str, Path]:
    """Return (display_name, dir_path) for a freshly-typed codename.

    The display name preserves the analyst's casing but with path-unsafe chars
    removed and whitespace collapsed. The dir_path is `<deals_root>/<display_name>`.

    Raises ValueError if the codename is empty after sanitisation.
    """
    if codename is None:
        raise ValueError("codename must be a non-empty string")
    display = _strip_unsafe(codename)
    if not display:
        raise ValueError(f"codename {codename!r} contains only path-unsafe characters")
    root = Path(deals_root).expanduser()
    return display, root / display


def find_existing(deals_root: Path | str, codename: str) -> Path | None:
    """Case-insensitive directory lookup under `deals_root`.

    Returns the existing Path if a folder whose name matches `codename` (after
    sanitisation, case-insensitive) is found; otherwise None.

    Non-existent or non-directory `deals_root` returns None — the conductor will
    create the root lazily on first write.
    """
    root = Path(deals_root).expanduser()
    if not root.exists() or not root.is_dir():
        return None
    display, _ = resolve(codename, root)
    needle = display.casefold()
    for child in root.iterdir():
        if child.is_dir() and child.name.casefold() == needle:
            return child
    return None


def disambiguate(deals_root: Path | str, codename: str, max_suggestions: int = 4) -> list[str]:
    """Suggest disambiguating codenames when `codename`'s directory already exists.

    Returns a list of display-form suggestions (e.g. `Project OpenText II`,
    `Project OpenText 2026`, `Project OpenText III`). The conductor presents
    these to the analyst — this helper does NOT pick one.

    Suggestions that already collide on disk are skipped.
    """
    root = Path(deals_root).expanduser()
    display, _ = resolve(codename, root)

    # Order matters: roman numerals first (analyst convention), then current year,
    # then -B / -C suffixes for the rare third/fourth collision.
    candidates = [
        f"{display} II",
        f"{display} 2026",
        f"{display} III",
        f"{display} 2025",
        f"{display}-B",
        f"{display}-C",
    ]
    out: list[str] = []
    for cand in candidates:
        if find_existing(root, cand) is None:
            out.append(cand)
        if len(out) >= max_suggestions:
            break
    return out
