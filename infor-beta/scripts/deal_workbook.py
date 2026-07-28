"""The deal's single workbook — one file, owned from stage one.

Phase D's centrepiece. Before this, every producing skill wrote a standalone
`.xlsx` and a final `workbook-aggregation` stage merged them: on Windows through
Excel COM, off-Windows through a best-effort openpyxl copy plus a LibreOffice
recalc. That merge is what produced the phase's whole tail of defects — the
LibreOffice `~` union-operator rewrite that made Excel repair-strip 44 formulas
on open, external-reference rebinding, order-dependent sheet renames, a
source-deletion gate to decide when the merge could be trusted, comment and
hyperlink carry-over, and the rule that a `financial-summary` LTM link resolves
*only* in the combined workbook. None of it exists here, because there is
nothing to merge: `pitch-<codename>.xlsx` is created at deal-init with its
template tabs already in place, and each stage writes its own tab.

Three consequences worth stating, because code elsewhere used to work around
their absence:

1. **Cross-tab links resolve immediately.** `financial-summary`'s
   `=INDEX('ltm-metrics'!…)` LTM lookup points at a tab in the same file the
   moment it is written, so nothing has to run "after aggregation".
2. **Tab names are fixed, not derived.** They are the constants below, matching
   the names the merge used to produce, so the deck assemblers and
   `financial_charts` address the same tabs as before.
3. **Writes are serialized.** The conductor dispatches a wave's stages
   concurrently, so `comps`, `precedents` and `financial-summary` can all reach
   for this file at once. `write_tab` takes an exclusive lock for the whole
   load → mutate → save cycle; two stages cannot interleave and lose one
   another's tab.

Template fidelity
-----------------
The five template tabs are copied in at init by copying the shipped
`INFOR Deal Workbook Template.xlsx`, which `tools/build_deal_workbook_template.py`
assembles from the four source templates. That template is a *pre-assembled*
artefact for a reason: openpyxl cannot move a sheet between workbooks, and a
cell-by-cell copy is precisely the lossy merge this phase deletes. So init is a
`shutil.copyfile` and nothing is reconstructed at runtime.

The Phase C `infor_` defined names are what make that safe. They are
**worksheet-scoped**, so they travel with their sheet and keep resolving after
the copy — `resolve_name_cell(ws, NAME_FX_RATE)` works on the `captable` tab of
a deal workbook exactly as it did on a standalone cap table. `write_tab` verifies
that on every write (see `_verify_tab_names`) rather than trusting it, because
`resolve_name_range` raises loudly and a stage failing three waves in is a worse
diagnostic than a stage failing at its first write.
"""

from __future__ import annotations

import os
import shutil
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from naming import safe_filename

# ─── Tab names ───────────────────────────────────────────────────────────────
# Verbatim the names the aggregator's merge produced (a single-sheet source took
# its skill key; the ownership pair kept its own sheet names), so every
# downstream reader — both deck assemblers, `financial_charts` — addresses the
# same tabs it always did.
TAB_CAPTABLE = "captable"
TAB_COMPS = "comps"
TAB_FINANCIAL_SUMMARY = "financial-summary"
TAB_LTM_METRICS = "ltm-metrics"
TAB_OWNERSHIP = "Ownership"
TAB_BLOOMBERG_OUTPUT = "Bloomberg Output"
TAB_PRECEDENTS = "precedents"

#: Canonical left-to-right order, matching the frozen production fixture
#: (`scripts/tests/fixtures/pitch-workbook.xlsx`) so a Phase D workbook opens
#: looking like the ones analysts have already been handed.
TAB_ORDER: tuple[str, ...] = (
    TAB_CAPTABLE,
    TAB_COMPS,
    TAB_FINANCIAL_SUMMARY,
    TAB_LTM_METRICS,
    TAB_OWNERSHIP,
    TAB_BLOOMBERG_OUTPUT,
    TAB_PRECEDENTS,
)

#: Tabs the shipped deal-workbook template carries (the rest are authored from
#: scratch by their producer, so they are created on first write).
TEMPLATE_TABS: frozenset[str] = frozenset(
    {TAB_CAPTABLE, TAB_COMPS, TAB_OWNERSHIP, TAB_BLOOMBERG_OUTPUT, TAB_PRECEDENTS}
)

#: Which tabs each deliverable needs. An earnings update has no comps /
#: precedents / ownership section, so those template tabs are dropped at init
#: rather than shipped empty in a client-facing workbook.
DELIVERABLE_TABS: dict[str, tuple[str, ...]] = {
    "pitch": TAB_ORDER,
    "earnings-update": (TAB_CAPTABLE, TAB_LTM_METRICS),
}

DEAL_WORKBOOK_TEMPLATE = "INFOR Deal Workbook Template.xlsx"


class DealWorkbookError(RuntimeError):
    """The deal workbook is missing, locked by a stuck writer, or malformed."""


# ─── Naming ──────────────────────────────────────────────────────────────────


def workbook_filename(deliverable_type: str, deal_name: str) -> str:
    """Return `<deliverable>-<deal name>.xlsx`.

    Unchanged from the aggregator's `combined_filename`, so a deal directory's
    workbook keeps the name analysts know: the deliverable prefix drops hyphens
    (`earnings-update` -> `earningsupdate`) while `pitch` stays `pitch`.
    """
    prefix = deliverable_type.replace("-", "").strip() or "deliverable"
    return f"{prefix}-{safe_filename(deal_name, default='Deal')}.xlsx"


def deal_workbook_path(
    deal_dir: Path | str, deliverable_type: str, deal_name: str
) -> Path:
    """Where the deal's workbook lives. Does not create it — see `init_deal_workbook`."""
    return Path(deal_dir).expanduser() / workbook_filename(deliverable_type, deal_name)


# ─── Init ────────────────────────────────────────────────────────────────────


def _template_path(template_name: str = DEAL_WORKBOOK_TEMPLATE) -> Path:
    """Resolve a shipped template through `CLAUDE_PLUGIN_ROOT`, as the skills do."""
    root = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", Path(__file__).resolve().parents[1]))
    candidate = root / "templates" / template_name
    if candidate.is_file():
        return candidate
    # A worktree / test run whose CLAUDE_PLUGIN_ROOT is unset or points elsewhere.
    fallback = Path(__file__).resolve().parents[1] / "templates" / template_name
    if fallback.is_file():
        return fallback
    raise DealWorkbookError(
        f"{template_name} not found under {candidate.parent} or {fallback.parent}. "
        f"It is assembled by tools/build_deal_workbook_template.py — re-run that "
        f"if a source template was re-saved."
    )


def init_deal_workbook(
    *,
    deal_dir: Path | str,
    deliverable_type: str,
    deal_name: str,
    overwrite: bool = False,
) -> Path:
    """Create the deal's workbook from the shipped template. Returns its path.

    Idempotent by default: an existing workbook is left exactly as it is, so
    re-running deal-init over a deal directory mid-run cannot discard the tabs
    stages have already written. Pass `overwrite=True` to start clean.

    Tabs the deliverable does not use are dropped (see `DELIVERABLE_TABS`), so
    an earnings-update workbook does not carry an empty comps section.
    """
    path = deal_workbook_path(deal_dir, deliverable_type, deal_name)
    if path.exists() and not overwrite:
        return path

    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(_template_path(), path)

    wanted = DELIVERABLE_TABS.get(deliverable_type)
    if wanted is not None:
        drop = [tab for tab in TEMPLATE_TABS if tab not in wanted]
        if drop:
            with _workbook_lock(path):
                wb = load_workbook(path)
                try:
                    for tab in drop:
                        if tab in wb.sheetnames:
                            del wb[tab]
                    wb.save(path)
                finally:
                    wb.close()
    return path


# ─── Serialized tab writes ───────────────────────────────────────────────────


@dataclass(frozen=True)
class TabSpec:
    """One tab's write.

    - ``write``: called as ``write(workbook, worksheet)`` with the deal workbook
      open and the target tab resolved. Everything cell-level stays in the
      producer module; this module owns only the file, the lock and the tab.
    - ``create``: the tab is authored from scratch rather than shipped by the
      template (`ltm-metrics`, `financial-summary`). A re-run replaces it, so a
      retried stage does not append a second copy or write over half of one.
    - ``verify_names``: `infor_` defined names that must resolve on the tab
      before ``write`` is called. Producers resolving through names pass theirs,
      turning a template that lost its names into an immediate, named failure.
    """

    write: Callable[[object, object], None]
    create: bool = False
    verify_names: tuple[str, ...] = ()


_LOCK_TIMEOUT_S = 120.0
_LOCK_POLL_S = 0.1
#: A lock older than this is from a writer that died (the conductor's stages are
#: single writes, not long sessions), so it is broken rather than waited out.
_LOCK_STALE_S = 300.0


class _workbook_lock:
    """Exclusive lock on one workbook, held for a whole load → mutate → save.

    A sibling `.lock` file created with `O_CREAT | O_EXCL` — uniform across
    Windows and Linux, unlike `msvcrt.locking` / `fcntl.flock`, and visible to a
    developer wondering why a stage is waiting. Waits up to `_LOCK_TIMEOUT_S`,
    breaks a lock left behind by a killed writer after `_LOCK_STALE_S`.
    """

    def __init__(self, path: Path):
        self._lock_path = Path(path).with_suffix(Path(path).suffix + ".lock")
        self._fd: int | None = None

    def __enter__(self) -> "_workbook_lock":
        deadline = time.monotonic() + _LOCK_TIMEOUT_S
        while True:
            try:
                self._fd = os.open(str(self._lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(self._fd, f"{os.getpid()}".encode())
                return self
            except FileExistsError:
                if self._break_if_stale():
                    continue
                if time.monotonic() >= deadline:
                    raise DealWorkbookError(
                        f"timed out after {_LOCK_TIMEOUT_S:.0f}s waiting for "
                        f"{self._lock_path.name}. Another stage is writing the deal "
                        f"workbook, or a killed one left the lock behind — delete "
                        f"the file to clear it."
                    ) from None
                time.sleep(_LOCK_POLL_S)

    def _break_if_stale(self) -> bool:
        try:
            age = time.time() - self._lock_path.stat().st_mtime
        except OSError:
            return True  # vanished — the holder released it
        if age <= _LOCK_STALE_S:
            return False
        try:
            self._lock_path.unlink()
        except OSError:
            return False
        return True

    def __exit__(self, *_exc) -> None:
        if self._fd is not None:
            try:
                os.close(self._fd)
            except OSError:
                pass
        try:
            self._lock_path.unlink()
        except OSError:
            pass


def _tab_index(tab_name: str, present: list[str]) -> int:
    """Where a newly created tab belongs, to keep `TAB_ORDER` left to right."""
    if tab_name not in TAB_ORDER:
        return len(present)
    wanted = TAB_ORDER.index(tab_name)
    for index, existing in enumerate(present):
        if existing in TAB_ORDER and TAB_ORDER.index(existing) > wanted:
            return index
    return len(present)


def _verify_tab_names(ws, tab_name: str, names: tuple[str, ...]) -> None:
    """Every named range the producer relies on must resolve on this tab."""
    if not names:
        return
    from template_layout import TemplateLayoutError, defined_name_ref

    missing = [name for name in names if defined_name_ref(ws, name) is None]
    if missing:
        raise TemplateLayoutError(
            f"deal workbook, tab {tab_name!r}: defined name(s) {', '.join(missing)} "
            f"do not resolve. The tab should carry them from "
            f"{DEAL_WORKBOOK_TEMPLATE}; re-run tools/build_deal_workbook_template.py "
            f"if a source template was re-saved."
        )


def write_tab(workbook_path: Path | str, tab_name: str, spec: TabSpec) -> None:
    """Apply `spec` to `tab_name` of the deal workbook, under an exclusive lock.

    The single mutation path. Serialized, so concurrent stages in one wave queue
    instead of clobbering each other; the whole load → mutate → save happens
    inside the lock, because openpyxl rewrites the entire file on save and a
    read that straddles another stage's save would drop that stage's tab.
    """
    path = Path(workbook_path)
    if not path.is_file():
        raise DealWorkbookError(
            f"{path} does not exist. The conductor creates the deal workbook at "
            f"deal-init via init_deal_workbook(); a stage must not create it."
        )

    with _workbook_lock(path):
        wb = load_workbook(path)
        try:
            if tab_name in wb.sheetnames:
                if spec.create:
                    del wb[tab_name]
                    ws = wb.create_sheet(tab_name, _tab_index(tab_name, wb.sheetnames))
                else:
                    ws = wb[tab_name]
            elif spec.create:
                ws = wb.create_sheet(tab_name, _tab_index(tab_name, wb.sheetnames))
            else:
                raise DealWorkbookError(
                    f"tab {tab_name!r} is not in {path.name} (has: "
                    f"{', '.join(wb.sheetnames)}). Template tabs are created at "
                    f"deal-init; pass TabSpec(create=True) for a tab authored "
                    f"from scratch."
                )

            _verify_tab_names(ws, tab_name, spec.verify_names)
            spec.write(wb, ws)
            wb.save(path)
        finally:
            wb.close()


def tab_names(workbook_path: Path | str) -> list[str]:
    """The workbook's tabs, in order. A cheap read for callers checking presence."""
    wb = load_workbook(Path(workbook_path), read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()
