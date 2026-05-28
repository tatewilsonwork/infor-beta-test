"""Aggregate every Excel workbook produced during a deliverable into one file.

The conductor runs the workbook-producing stages (captable, ltm-revenue,
comps, ...) as siblings, each emitting a standalone `.xlsx` under the deal's
`artefacts/`. This helper is the final consolidation stage: it merges those
workbooks into a single combined workbook named `<deliverable>-<deal name>.xlsx`
(e.g. `earningsupdate-Project Atlas.xlsx`, `pitch-Project Atlas.xlsx`), with
each source contributing its sheets under a tab named after the producing
skill. The individual source workbooks are deleted once the merge succeeds
(the combined file replaces them).

Two merge backends, mirroring `excel_to_powerpoint.py`:

  - **Excel COM** (Windows + Excel) — copies whole worksheet collections
    between open workbooks, so formulas, CapIQ links, charts, and formatting
    survive intact, and intra-workbook references stay internal.

  - **openpyxl** (Cowork / Linux / macOS, or Windows without Excel) — a
    best-effort cell-and-style copy. External data connections (CapIQ) and
    charts do NOT survive this path; use it only when COM is unavailable.

Tab naming: a single-sheet source becomes one tab named after the skill
(`captable`, `ltm-revenue`); a multi-sheet source contributes one tab per
sheet named `<skill>-<sheet>`. Excel's 31-char / forbidden-character / unique
constraints are enforced in both backends.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Mapping

# xlOpenXMLWorkbook — the .xlsx SaveAs file format for Excel COM.
_XL_OPEN_XML_WORKBOOK = 51
_EXCEL_SHEET_NAME_MAX = 31
_FORBIDDEN_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]+")


def _safe_file_stem(value: str) -> str:
    safe = re.sub(r"[/\\:*?\"<>|]+", "-", value).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe or "Deal"


def combined_filename(deliverable_type: str, deal_name: str) -> str:
    """Return `<deliverable>-<deal name>.xlsx`.

    The deliverable prefix drops hyphens so `earnings-update` reads as
    `earningsupdate` (matching the analyst-facing naming), while `pitch`
    stays `pitch`.
    """
    prefix = deliverable_type.replace("-", "").strip() or "deliverable"
    return f"{prefix}-{_safe_file_stem(deal_name)}.xlsx"


def _excel_safe_sheet_name(name: str) -> str:
    """Sanitize a tab name to Excel's rules: no []:*?/\\, <=31 chars, non-empty."""
    safe = _FORBIDDEN_SHEET_CHARS.sub("-", name).strip().strip("'")
    safe = safe[:_EXCEL_SHEET_NAME_MAX].strip()
    return safe or "Sheet"


def _unique_sheet_name(name: str, used: set[str]) -> str:
    """Return a sheet name unique within `used` (case-insensitive), <=31 chars."""
    candidate = _excel_safe_sheet_name(name)
    if candidate.casefold() not in used:
        used.add(candidate.casefold())
        return candidate
    for i in range(2, 1000):
        suffix = f" ({i})"
        trimmed = candidate[: _EXCEL_SHEET_NAME_MAX - len(suffix)].strip()
        attempt = f"{trimmed}{suffix}"
        if attempt.casefold() not in used:
            used.add(attempt.casefold())
            return attempt
    raise ValueError(f"could not derive a unique sheet name from {name!r}")


def _tab_name(skill: str, original_sheet: str, sheet_count: int) -> str:
    """Single-sheet source -> the skill name; multi-sheet -> `<skill>-<sheet>`."""
    return skill if sheet_count == 1 else f"{skill}-{original_sheet}"


def _resolve_sources(
    sources: Mapping[str, str | Path | None],
) -> tuple[list[tuple[str, Path]], list[str]]:
    """Drop None / missing-file entries; return (kept, skipped-skill-names)."""
    kept: list[tuple[str, Path]] = []
    skipped: list[str] = []
    for skill, raw in sources.items():
        if raw is None:
            skipped.append(skill)
            continue
        path = Path(str(raw)).expanduser()
        if not path.exists():
            skipped.append(skill)
            continue
        kept.append((skill, path.resolve()))
    return kept, skipped


def combine_workbooks(
    *,
    sources: Mapping[str, str | Path | None],
    output_dir: Path | str,
    deliverable_type: str,
    deal_name: str,
    delete_sources: bool = True,
) -> Path:
    """Merge the source workbooks into one combined `.xlsx` and return its path.

    `sources` maps a producing-skill name (used as the tab name) to that
    skill's workbook path. None values and non-existent paths are skipped, so
    optional upstream workbooks (e.g. a comps workbook that wasn't produced)
    can be wired in unconditionally. Raises ValueError if nothing is left to
    combine.

    When `delete_sources` is True (the default), the individual source files
    are removed after a successful merge — the combined workbook replaces them.
    """
    kept, _skipped = _resolve_sources(sources)
    if not kept:
        raise ValueError("no workbooks to combine — every source was None or missing")

    out_dir = Path(output_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / combined_filename(deliverable_type, deal_name)

    if sys.platform == "win32":
        try:
            _combine_via_com(kept, output_path)
        except RuntimeError:
            _combine_via_openpyxl(kept, output_path)
    else:
        _combine_via_openpyxl(kept, output_path)

    if delete_sources:
        for _skill, path in kept:
            if path.resolve() != output_path.resolve():
                path.unlink(missing_ok=True)

    return output_path


def _combine_via_com(sources: list[tuple[str, Path]], output_path: Path) -> None:
    """Merge with Excel COM, preserving formulas, links, charts, and formatting."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for COM-based workbook aggregation "
            "(Windows + Microsoft Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        combined = excel.Workbooks.Add()
        # Sheets present in the blank workbook; deleted once real sheets land.
        seed_sheets = [combined.Sheets(i + 1).Name for i in range(combined.Sheets.Count)]
        used_names: set[str] = set()

        for skill, path in sources:
            src = excel.Workbooks.Open(str(path), ReadOnly=True, UpdateLinks=0)
            try:
                original_names = [
                    src.Worksheets(i + 1).Name for i in range(src.Worksheets.Count)
                ]
                before = combined.Sheets.Count
                # Copy the whole worksheet collection in one operation so any
                # intra-workbook references resolve against the new copies.
                src.Worksheets.Copy(After=combined.Sheets(combined.Sheets.Count))
                for offset, original in enumerate(original_names):
                    sheet = combined.Sheets(before + 1 + offset)
                    sheet.Name = _unique_sheet_name(
                        _tab_name(skill, original, len(original_names)), used_names
                    )
            finally:
                src.Close(SaveChanges=False)

        # Drop the blank workbook's seed sheets now that real content exists.
        for name in seed_sheets:
            if combined.Sheets.Count > 1:
                combined.Sheets(name).Delete()

        if output_path.exists():
            output_path.unlink()
        combined.SaveAs(str(output_path), FileFormat=_XL_OPEN_XML_WORKBOOK)
        combined.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _combine_via_openpyxl(sources: list[tuple[str, Path]], output_path: Path) -> None:
    """Best-effort merge with openpyxl. CapIQ links and charts do NOT survive."""
    from openpyxl import Workbook, load_workbook

    combined = Workbook()
    combined.remove(combined.active)  # drop the default empty sheet
    used_names: set[str] = set()

    for skill, path in sources:
        src = load_workbook(path, data_only=False)
        sheet_count = len(src.sheetnames)
        for original in src.sheetnames:
            src_ws = src[original]
            title = _unique_sheet_name(
                _tab_name(skill, original, sheet_count), used_names
            )
            dst_ws = combined.create_sheet(title=title)
            _copy_sheet(src_ws, dst_ws)

    if not combined.sheetnames:
        combined.create_sheet(title="Sheet")
    combined.save(output_path)


def _copy_sheet(src_ws, dst_ws) -> None:
    """Copy cell values/formulas, styles, dimensions, merges between sheets."""
    from copy import copy

    for row in src_ws.iter_rows():
        for cell in row:
            dst = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                dst.font = copy(cell.font)
                dst.fill = copy(cell.fill)
                dst.border = copy(cell.border)
                dst.alignment = copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.protection = copy(cell.protection)

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    for key, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[key].width = dim.width
        dst_ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[key].height = dim.height
        dst_ws.row_dimensions[key].hidden = dim.hidden

    dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
