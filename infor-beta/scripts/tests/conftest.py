"""Pytest config — put `scripts/` on sys.path so tests can `import schemas`,
and gate the Excel-COM tests behind an opt-in env var.

The repo root holds `pyproject.toml`, the plugin root holds `scripts/`, and the
tests live in `scripts/tests/`. From the tests directory, the package layout is:

    ../schemas/      ← `from schemas import Company, ...`
    ../codename.py   ← `import codename`

so we prepend the parent directory (`scripts/`) to sys.path here. This makes
`pytest infor-beta/scripts/tests/` work from any cwd without depending on
pyproject's `pythonpath` setting.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))


# ─── One shared render cache per run (Phase C) ────────────────────────────────
# `slide_render` caches converted PDFs by deck content. Pointing every process at
# one directory makes the cache shared, which matters most for the work that is
# otherwise duplicated per worker: the blank library's baseline render and its
# attribution probe deck are identical in all six workers, and paying for them
# once instead of six times is the difference between a distributed run helping
# and just thrashing.
#
# Set at import, before xdist spawns its workers, so they inherit it. A worker
# re-importing this file finds the variable already set and joins the existing
# directory rather than making its own — and only the process that created it
# removes it, at exit, so nothing is deleted from under a live worker.
#
# Deliberately fresh per run rather than persistent: a cache surviving between
# runs would make the reported suite time depend on whether someone had run it
# before, and would keep serving decks whose source template had changed.
if not os.environ.get("INFOR_RENDER_CACHE_DIR"):
    import atexit
    import tempfile

    _RENDER_CACHE_DIR = tempfile.mkdtemp(prefix="infor-suite-render-cache-")
    os.environ["INFOR_RENDER_CACHE_DIR"] = _RENDER_CACHE_DIR
    atexit.register(shutil.rmtree, _RENDER_CACHE_DIR, True)


# ─── Office-COM serialization under xdist — no longer needed ────────────────
# Phase D deleted every COM path (the Excel range renderer and chart builders,
# and slide_render's PowerPoint backend), so nothing in the suite drives a
# per-user COM singleton any more. What went with it:
#
#   - the `excel_com` marker and its `INFOR_EXCEL_COM_TESTS` opt-in gate. Those
#     four tests each spawned a real, VISIBLE Excel (an invisible instance
#     exports a blank picture after a recalc), into which the Cap IQ Office Tools
#     add-in loaded and raised a modal dialog `DisplayAlerts = False` cannot
#     suppress. Gating them was the prevention that worked; deleting the path is
#     the cure.
#   - the `office_com` xdist group that pinned `test_render_parity` and
#     `test_slide_render` to one worker so two workers could not drive
#     PowerPoint at once. Both now render only through LibreOffice, which is safe
#     concurrently because each process gets its own profile (`slide_render`).
#
# `test_powerpoint_com_render_still_works` — the accepted flake in a full run —
# is deleted with the backend it covered.
