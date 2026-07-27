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
import sys
from pathlib import Path

import pytest

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))


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


def pytest_collection_modifyitems(config, items) -> None:
    """Skip every `excel_com`-marked test unless the env var opts in."""
    if _excel_com_opted_in():
        return
    skip = pytest.mark.skip(
        reason=(
            f"Excel-COM test: set {_EXCEL_COM_ENV_VAR}=1 to run. Spawns a real "
            "visible Excel, whose Cap IQ Office Tools add-in can raise a modal "
            "dialog that stalls the run."
        )
    )
    for item in items:
        if "excel_com" in item.keywords:
            item.add_marker(skip)
