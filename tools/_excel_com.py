"""The one Excel-COM instance owner for the prep tooling.

Phase D deleted every COM path from the shipped plugin, including
`excel_to_powerpoint.excel_com_app` — which `add_template_named_ranges.py`'s
`--verify-excel` oracle imported. COM now lives only in `tools/`, so its single
owner lives here rather than being hand-rolled twice: that consolidation is what
v0.5.38 did for the four in-plugin call sites after they had already diverged on
the `Quit` guard, and there is no reason to re-earn the lesson.

Two properties, both established by measurement during Phase D and both easy to
get silently wrong:

**Early binding.** Late-bound `Worksheet.Copy` mis-marshals its optional
arguments — `Copy(After=…)` did not pass the keyword at all, so Excel saw a bare
`Copy()` (documented to copy into a NEW workbook) and did exactly that, leaving
the destination untouched and the caller none the wiser. With the type library
generated first every spelling works. `_assert_early_bound` refuses to proceed on
a dynamic wrapper, because a silent mis-copy is far worse than a stop.

**`DispatchEx`, never `Dispatch`/`EnsureDispatch`.** Those reuse an instance from
the running-object table. During Phase D one attached to a wedged automation Excel
left over from an earlier probe and every subsequent call returned `0x800ac472`.
`DispatchEx` always makes its own process — which also keeps the Cap IQ Office
Tools add-in out of it, since an automation-started Excel loads no
`OPEN=`-registered add-in. That is what makes it safe to open a Capital IQ
workbook here without the add-in touching `CIQWBGuid` / `CIQWBInfo`, the objection
`add_template_named_ranges.py` correctly raised against COM for template surgery.

The `Quit` is guarded, so a raising `Quit` cannot skip `CoUninitialize()` and
leak the apartment — the exact divergence v0.5.38 found across the four
in-plugin sites.
"""

from __future__ import annotations

import contextlib
import sys
from collections.abc import Iterator

#: Excel's type library. v1.9 covers Excel 2016 / 2019 / 365; older
#: majors/minors are tried in turn so this is not pinned to one install.
EXCEL_TYPELIB = "{00020813-0000-0000-C000-000000000046}"
_TYPELIB_VERSIONS = ((1, 9), (1, 8), (1, 7), (1, 6))


def _assert_early_bound(excel) -> None:
    """Stop unless `excel` is an early-bound wrapper (a dynamic one fails silently)."""
    if type(excel).__name__ == "CDispatch":
        raise SystemExit(
            "Excel COM is bound dynamically, and late-bound Worksheet.Copy silently\n"
            "copies to the wrong workbook. Generate the type library first:\n"
            '    python -c "from win32com.client import gencache; '
            f"gencache.EnsureModule('{EXCEL_TYPELIB}', 0, 1, 9)\"\n"
            "then re-run. (Delete %LOCALAPPDATA%\\\\Temp\\\\gen_py if it is stale.)"
        )


@contextlib.contextmanager
def excel_com_app(*, purpose: str, visible: bool = False) -> Iterator:
    """A fresh, early-bound Excel for `purpose`; always quit and torn down.

    `purpose` appears in the error when startup fails, so a failure says which
    prep step wanted Excel.
    """
    try:
        import pythoncom
        import win32com.client
        from win32com.client import gencache
    except ImportError as exc:  # pragma: no cover - dev-box only
        raise RuntimeError(
            f"pywin32 is required for {purpose}; install the dev extra "
            f"(pip install -e '.[dev]')"
        ) from exc

    for major, minor in _TYPELIB_VERSIONS:
        try:
            gencache.EnsureModule(EXCEL_TYPELIB, 0, major, minor)
            break
        except Exception:  # noqa: BLE001 - try the next version
            continue

    pythoncom.CoInitialize()
    excel = None
    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"could not start Excel for {purpose}: {exc}") from exc
        _assert_early_bound(excel)
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        yield excel
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception as exc:  # noqa: BLE001
                # Guarded so a raising Quit cannot skip CoUninitialize().
                print(f"  Excel Quit raised during {purpose}: {exc}", file=sys.stderr)
        pythoncom.CoUninitialize()


def disable_autosave(workbook) -> None:
    """Turn AutoSave off for an open workbook, where the property exists.

    Second line of defence behind opening only STAGED COPIES of repo files. The
    repo lives in a OneDrive-synced folder, where Excel 365 enables AutoSave by
    default and writes the file back regardless of `Close(SaveChanges=False)` —
    which silently re-saved all four source templates through Excel during Phase
    D development, undoing Phase C's byte-level preservation.
    """
    try:
        workbook.AutoSaveOn = False
    except Exception:  # noqa: BLE001, S110 - absent or not applicable
        pass
