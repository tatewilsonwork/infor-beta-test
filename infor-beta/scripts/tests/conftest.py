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

import pytest

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


# ─── Excel-COM opt-in gate (Phase I item 2) ───────────────────────────────────
# The `excel_com` marker covers the tests that spawn a REAL Excel through
# `DispatchEx` and must run it `Visible = True` (`CopyPicture(xlScreen)` captures
# what the instance renders, so an invisible one exports a blank picture). The S&P
# Cap IQ Office Tools add-in loads into that throwaway instance, finds its Cap IQ
# Pro sibling absent, and raises its own `MessageBox` — which
# `excel.DisplayAlerts = False` cannot suppress, because that gates Excel's own
# alerts, not a third-party add-in's. Up to four modal dialogs per run, each
# stalling the suite until someone clicks it.
#
# Default off, matching where Phase A already went (LibreOffice everywhere by
# default; PowerPoint COM opt-in via `INFOR_SLIDE_RENDER_BACKEND`). Accepted
# consequence: the Windows COM path loses routine coverage — Phase D deletes it,
# and nothing else depends on these tests. The marker is declared in
# `pyproject.toml`'s `markers` list.
_EXCEL_COM_ENV_VAR = "INFOR_EXCEL_COM_TESTS"


def _excel_com_opted_in() -> bool:
    return os.environ.get(_EXCEL_COM_ENV_VAR, "").strip().lower() not in (
        "",
        "0",
        "false",
        "no",
    )


# ─── Office-COM serialization under xdist (Phase C) ──────────────────────────
# The suite runs distributed by default (`-n auto` in pyproject). Excel and
# PowerPoint's COM servers are per-user singletons, so two workers driving one
# of them at the same time can interfere — `test_powerpoint_com_render_still_works`
# already flakes in a full serial run and passes in isolation, and running it
# beside `test_render_parity`'s PowerPoint pass would only widen that window.
#
# Every Office-COM test is therefore put in one xdist group, which pins the
# group to a single worker and so keeps them serialized relative to each other.
# Grouping is applied here rather than as per-test markers so a new COM test
# cannot forget it: membership follows from the marker/module that already
# identifies these tests.
_COM_GROUP = "office_com"
_COM_MODULES = ("test_render_parity.py", "test_slide_render.py")


def pytest_collection_modifyitems(config, items) -> None:
    """Gate the Excel-COM tests, and keep every Office-COM test on one worker."""
    skip = pytest.mark.skip(
        reason=(
            f"Excel-COM test: set {_EXCEL_COM_ENV_VAR}=1 to run. Spawns a real "
            "visible Excel, whose Cap IQ Office Tools add-in can raise a modal "
            "dialog that stalls the run."
        )
    )
    opted_in = _excel_com_opted_in()
    for item in items:
        if "excel_com" in item.keywords:
            if not opted_in:
                item.add_marker(skip)
            item.add_marker(pytest.mark.xdist_group(_COM_GROUP))
        elif _module_name(item) in _COM_MODULES:
            item.add_marker(pytest.mark.xdist_group(_COM_GROUP))


def _module_name(item) -> str:
    """The item's test-module filename, or '' when it has none."""
    path = getattr(item, "path", None)
    return Path(str(path)).name if path is not None else ""
