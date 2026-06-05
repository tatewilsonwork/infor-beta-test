"""Aggregate every Excel workbook produced during a deliverable into one file.

The conductor runs the workbook-producing stages (captable, ltm-metrics,
comps, ...) as siblings, each emitting a standalone `.xlsx` under the deal's
`artefacts/`. This helper is the final consolidation stage: it merges those
workbooks into a single combined workbook named `<deliverable>-<deal name>.xlsx`
(e.g. `earningsupdate-Project Atlas.xlsx`, `pitch-Project Atlas.xlsx`), with
each source contributing its sheets under a tab named after the producing
skill. The individual source workbooks are deleted once the merge succeeds
(the combined file replaces them).

Two merge backends, mirroring `excel_to_powerpoint.py`:

  - **Excel COM** (Windows + Excel) — when a `captable` source is present it is
    opened as the **base** workbook (and saved as the combined file), so the cap
    table's theme, CapIQ links and intra-workbook references survive intact; the
    other skills' sheets are copied in after it. (A blank workbook was used as
    the base before v0.5.9, which shifted theme colours and, on the openpyxl
    fallback, dropped CapIQ links — making the combined file hard to format and
    link.) CapIQ's very-hidden `__snloffice` helper sheet is dropped so it never
    surfaces as a tab. Each source's content sheets are copied **as a group in a
    single operation** so a source's intra-workbook cross-sheet references stay
    internal (the ownership `Ownership` sheet's hundreds of `='Bloomberg
    Output'!…` lookups would otherwise become external links to the soon-deleted
    source and resolve to `#REF`); the copy destination is given positionally,
    because the named `After=` argument is silently dropped by some Excel builds
    (they then copy into a brand-new workbook — a no-op append that, pre-v0.5.10,
    forced the openpyxl fallback and lost the theme).

  - **openpyxl** (Cowork / Linux / macOS, or Windows without Excel) — a
    best-effort cell-and-style copy. External data connections (CapIQ) and
    charts do NOT survive this path; use it only when COM is unavailable.

Theme: the combined workbook is stamped with the INFOR brand theme
(`templates/INFORFG.thmx`) on both backends — `ApplyTheme` under COM,
`loaded_theme` under openpyxl — so it carries INFOR colours/fonts even when the
merge base is a blank workbook (no cap table) or the openpyxl fallback runs.

Cross-tab links: once every workbook is one file, a relink pass rewrites the
skills' standalone scalar handoffs into live cross-tab formulas, so the
analyst's combined workbook stays internally linked — the cap table's LTM
Revenue / Adj. EBITDA cells (`D47`/`D48`) point at the `ltm-metrics` bridge
totals, and the ownership % denominator (`F35`) points at the cap table's basic
shares. See `_relink_cross_tab_openpyxl` / `_relink_cross_tab_com`.

Tab naming: a single-sheet source becomes one tab named after the skill
(`captable`, `ltm-metrics`); a multi-sheet source contributes one tab per
sheet named `<skill>-<sheet>`. Excel's 31-char / forbidden-character / unique
constraints are enforced in both backends.
"""

from __future__ import annotations

import re
import sys
import time
import zipfile
from pathlib import Path
from typing import Mapping

# xlOpenXMLWorkbook — the .xlsx SaveAs file format for Excel COM.
_XL_OPEN_XML_WORKBOOK = 51
_EXCEL_SHEET_NAME_MAX = 31
_FORBIDDEN_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]+")

# INFOR brand theme, applied to every combined workbook so it keeps INFOR
# colours/fonts regardless of merge base or backend. Ships beside the templates.
_THEME_FILENAME = "INFORFG.thmx"
# A .thmx is a zip; the theme XML Excel/openpyxl want lives at this member path.
_THMX_THEME_MEMBER = "theme/theme/theme1.xml"

# CapIQ's Excel add-in stows formula metadata in a very-hidden helper sheet
# named "__snloffice". Copied verbatim it surfaces as a garbled (CJK-looking)
# tab in the combined workbook, so it's excluded from aggregation. The add-in
# regenerates it on refresh, so dropping it from the merged file is harmless.
_CAPIQ_HELPER_SHEET = re.compile(r"^__snl", re.IGNORECASE)


def _is_capiq_helper_sheet(name: str) -> bool:
    return bool(_CAPIQ_HELPER_SHEET.match(name.strip()))


def _default_theme_path() -> Path:
    """Path to the shipped INFOR theme (`templates/INFORFG.thmx`)."""
    return Path(__file__).resolve().parent.parent / "templates" / _THEME_FILENAME


def _resolve_theme_path(theme_path: Path | str | None) -> Path | None:
    """Return an existing theme path (caller override or the shipped default), or
    None when no theme file is available — theming is then skipped."""
    candidate = Path(theme_path) if theme_path is not None else _default_theme_path()
    return candidate if candidate.exists() else None


def _extract_theme_xml(theme_path: Path) -> bytes | None:
    """Return the `theme1.xml` bytes from a `.thmx` package for openpyxl theme
    injection, or None if the file is missing / not a valid theme zip."""
    try:
        with zipfile.ZipFile(theme_path) as z:
            return z.read(_THMX_THEME_MEMBER)
    except (OSError, KeyError, zipfile.BadZipFile):
        return None


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
    """Single-sheet source -> the skill name; multi-sheet -> the original sheet
    names, unprefixed.

    A single-sheet workbook's sheet is usually generic (`Sheet1`, `Cap with
    Links`), so the skill name is the more useful tab label. A multi-sheet source
    carries self-describing sheet names (`Ownership`, `Bloomberg Output`) that the
    analyst expects to see verbatim — and prefixing them would force a rename that
    breaks the source's intra-workbook cross-sheet references (the ownership
    `Ownership` sheet's `='Bloomberg Output'!…` lookups -> `#REF`). So multi-sheet
    sources keep their sheet names; only the `_unique_sheet_name` collision guard
    can alter them."""
    return skill if sheet_count == 1 else original_sheet


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
    theme_path: Path | str | None = None,
) -> Path:
    """Merge the source workbooks into one combined `.xlsx` and return its path.

    `sources` maps a producing-skill name (used as the tab name) to that
    skill's workbook path. None values and non-existent paths are skipped, so
    optional upstream workbooks (e.g. a comps workbook that wasn't produced)
    can be wired in unconditionally. Raises ValueError if nothing is left to
    combine.

    When `delete_sources` is True (the default), the individual source files
    are removed after a successful merge — the combined workbook replaces them.

    `theme_path` overrides the brand theme stamped on the combined workbook;
    it defaults to the shipped `templates/INFORFG.thmx`. Theming is skipped
    silently if no theme file is found.
    """
    kept, _skipped = _resolve_sources(sources)
    if not kept:
        raise ValueError("no workbooks to combine — every source was None or missing")

    out_dir = Path(output_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / combined_filename(deliverable_type, deal_name)

    theme = _resolve_theme_path(theme_path)
    if sys.platform == "win32":
        try:
            _combine_via_com(kept, output_path, theme)
        except RuntimeError:
            _combine_via_openpyxl(kept, output_path, theme)
    else:
        _combine_via_openpyxl(kept, output_path, theme)

    if delete_sources:
        for _skill, path in kept:
            if path.resolve() != output_path.resolve():
                path.unlink(missing_ok=True)

    return output_path


# ─── Cross-tab relink ────────────────────────────────────────────────────────
# Once every workbook is merged into one file, these cells are rewritten from
# the skills' standalone scalar handoffs to live cross-tab formulas, so the
# analyst's combined workbook stays internally linked: edit the LTM bridge and
# the cap table follows; the ownership % tracks the cap table's share count.
_LTM_REVENUE_LABEL = "(=) LTM Revenue"
_LTM_EBITDA_LABELS = ("(=) LTM Adj. EBITDA", "(=) LTM EBITDA")
_CAP_LTM_REVENUE_CELL = "D47"   # cap table LTM Revenue (millions, F5 currency)
_CAP_LTM_EBITDA_CELL = "D48"    # cap table LTM Adj. EBITDA
_CAP_BASIC_SHARES_CELL = "F17"  # cap table basic shares outstanding (millions)
_OWN_DENOM_CELL = "F35"         # ownership % denominator (full units)
_OWN_SHARES_SCALE = 1_000_000   # cap table is in millions; ownership is full units


def _quote_sheet(name: str) -> str:
    """Single-quote a sheet name for use in a formula (handles spaces/hyphens)."""
    return "'" + name.replace("'", "''") + "'"


def _find_label_row_openpyxl(ws, prefixes) -> int | None:
    """Row (1-based) of the first col-A cell whose text starts with any prefix."""
    for row in ws.iter_rows(min_col=1, max_col=1):
        value = row[0].value
        if isinstance(value, str) and any(value.strip().startswith(p) for p in prefixes):
            return row[0].row
    return None


def _relink_cross_tab_openpyxl(combined, skill_to_tab: dict[str, str]) -> None:
    """Wire the cap table's LTM cells + ownership denominator to sibling tabs.

    No-op unless the relevant tabs exist. The LTM bridge total rows are dynamic
    (they depend on the segment/component counts), so they are located by their
    `(=) LTM Revenue` / `(=) LTM Adj. EBITDA` labels rather than a fixed cell.
    CapIQ links are already lost on the openpyxl path, so this only restores the
    cross-tab references; the COM path preserves the live CapIQ formulas too.
    """
    cap = skill_to_tab.get("captable")
    ltm = skill_to_tab.get("ltm-metrics")
    own = skill_to_tab.get("ownership")
    names = set(combined.sheetnames)
    if cap in names and ltm in names:
        ltm_ws = combined[ltm]
        rev = _find_label_row_openpyxl(ltm_ws, (_LTM_REVENUE_LABEL,))
        ebitda = _find_label_row_openpyxl(ltm_ws, _LTM_EBITDA_LABELS)
        q = _quote_sheet(ltm)
        if rev is not None:
            combined[cap][_CAP_LTM_REVENUE_CELL] = f"={q}!B{rev}*F7"
        if ebitda is not None:
            combined[cap][_CAP_LTM_EBITDA_CELL] = f"={q}!B{ebitda}*F7"
    if cap in names and own in names:
        combined[own][_OWN_DENOM_CELL] = (
            f"={_quote_sheet(cap)}!{_CAP_BASIC_SHARES_CELL}*{_OWN_SHARES_SCALE}"
        )


def _find_label_row_com(ws, prefixes) -> int | None:
    """COM counterpart of `_find_label_row_openpyxl`."""
    try:
        used = ws.UsedRange
        last = used.Row + used.Rows.Count - 1
    except Exception:
        last = 200
    for r in range(1, last + 1):
        value = ws.Cells(r, 1).Value
        if isinstance(value, str) and any(value.strip().startswith(p) for p in prefixes):
            return r
    return None


def _relink_cross_tab_com(combined, skill_to_tab: dict[str, str]) -> None:
    """COM counterpart of `_relink_cross_tab_openpyxl`; best-effort (never raises
    — a relink failure must not lose the successfully merged workbook)."""
    try:
        cap = skill_to_tab.get("captable")
        ltm = skill_to_tab.get("ltm-metrics")
        own = skill_to_tab.get("ownership")
        names = {combined.Sheets(i + 1).Name for i in range(combined.Sheets.Count)}
        if cap in names and ltm in names:
            ltm_ws = combined.Worksheets(ltm)
            rev = _find_label_row_com(ltm_ws, (_LTM_REVENUE_LABEL,))
            ebitda = _find_label_row_com(ltm_ws, _LTM_EBITDA_LABELS)
            cap_ws = combined.Worksheets(cap)
            q = _quote_sheet(ltm)
            if rev is not None:
                cap_ws.Range(_CAP_LTM_REVENUE_CELL).Formula = f"={q}!B{rev}*F7"
            if ebitda is not None:
                cap_ws.Range(_CAP_LTM_EBITDA_CELL).Formula = f"={q}!B{ebitda}*F7"
        if cap in names and own in names:
            combined.Worksheets(own).Range(_OWN_DENOM_CELL).Formula = (
                f"={_quote_sheet(cap)}!{_CAP_BASIC_SHARES_CELL}*{_OWN_SHARES_SCALE}"
            )
    except Exception:
        pass


def _open_workbook(excel, path: Path, *, read_only: bool):
    """Open a workbook robustly under COM.

    Works around two pywin32 quirks: `Workbooks.Open` sometimes returns None even
    though the workbook opened (grab the latest one — but only when the open count
    actually rose, so a genuine failure raises instead of silently handing back
    the wrong workbook), and a transient file-lock/COM race can fail the first
    open (one short retry absorbs it). Raises RuntimeError on real failure so the
    caller falls back to the openpyxl backend."""
    last_exc: Exception | None = None
    for attempt in range(2):
        before = excel.Workbooks.Count
        try:
            wb = excel.Workbooks.Open(str(path), ReadOnly=read_only, UpdateLinks=0)
        except Exception as exc:  # COM error — retry once, then surface
            last_exc = exc
            wb = None
        if wb is None and excel.Workbooks.Count > before:
            wb = excel.Workbooks(excel.Workbooks.Count)
        if wb is not None:
            return wb
        time.sleep(0.4 * (attempt + 1))
    raise RuntimeError(f"Excel failed to open workbook: {path} ({last_exc})")


def _rename_and_clean_base(combined, skill: str, used_names: set, skill_to_tab: dict) -> None:
    """Rename the base (cap table) workbook's content sheets to the skill tab name
    and drop any CapIQ helper sheet, so the base reads like any merged source."""
    original = [combined.Worksheets(i + 1).Name for i in range(combined.Worksheets.Count)]
    content_count = sum(1 for n in original if not _is_capiq_helper_sheet(n))
    drop: list = []
    for idx in range(1, combined.Sheets.Count + 1):
        sheet = combined.Sheets(idx)
        if _is_capiq_helper_sheet(sheet.Name):
            drop.append(sheet)
            continue
        new_name = _unique_sheet_name(_tab_name(skill, sheet.Name, content_count), used_names)
        sheet.Name = new_name
        skill_to_tab.setdefault(skill, new_name)
    for sheet in drop:
        sheet.Visible = -1  # xlSheetVisible — can't delete a very-hidden sheet
        sheet.Delete()


def _combine_via_com(
    sources: list[tuple[str, Path]], output_path: Path, theme_path: Path | None = None
) -> None:
    """Merge with Excel COM, preserving formulas, links, charts, and formatting.

    When a `captable` source is present it is opened as the BASE workbook (and
    saved as the combined file), so the cap table's theme, CapIQ links and
    intra-workbook references survive intact; the other skills' sheets are copied
    in after it. Without a cap table, a blank workbook is the base (legacy
    behaviour). After the merge, a best-effort cross-tab relink wires the cap
    table's LTM cells to the ltm-metrics tab and the ownership denominator to the
    cap table, and (when `theme_path` is given) the INFOR brand theme is stamped
    on so a blank-base merge doesn't ship the default Office theme.
    """
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

        used_names: set[str] = set()
        skill_to_tab: dict[str, str] = {}

        # The cap table (when present) is the base workbook so its formatting and
        # CapIQ links survive; the remaining sources are copied into it.
        base = next(((s, p) for s, p in sources if s == "captable"), None)
        if base is not None:
            rest = [entry for entry in sources if entry != base]
            combined = _open_workbook(excel, base[1], read_only=False)
            seed_sheets: list[str] = []
            _rename_and_clean_base(combined, base[0], used_names, skill_to_tab)
        else:
            rest = list(sources)
            combined = excel.Workbooks.Add()
            # Sheets present in the blank workbook; deleted once real sheets land.
            seed_sheets = [combined.Sheets(i + 1).Name for i in range(combined.Sheets.Count)]

        # Copy each source's content sheets into `combined`. All of a source's
        # sheets are copied in ONE operation so any intra-workbook cross-sheet
        # references survive as internal references (copying sheet-by-sheet turns
        # them into external links to the soon-deleted source -> #REF; this is the
        # ownership `Ownership` -> `Bloomberg Output` case). The destination MUST
        # be the active workbook or Excel copies into a NEW workbook instead of
        # appending here, so re-activate `combined` first; and the destination is
        # passed POSITIONALLY (Before=None, After=<last sheet>) because the named
        # `After=` form is silently dropped by some Excel builds, which then copy
        # into a stray new workbook (a no-op append). CapIQ `__snl*` helper sheets
        # are skipped at the source so they never surface as a tab.
        for skill, path in rest:
            src = _open_workbook(excel, path, read_only=True)
            try:
                content = [
                    src.Worksheets(i + 1).Name
                    for i in range(src.Worksheets.Count)
                    if not _is_capiq_helper_sheet(src.Worksheets(i + 1).Name)
                ]
                if not content:
                    continue
                before = combined.Sheets.Count
                combined.Activate()
                selector = content[0] if len(content) == 1 else content
                src.Worksheets(selector).Copy(None, combined.Sheets(combined.Sheets.Count))
                added = combined.Sheets.Count - before
                if added < len(content):
                    # Excel copied to a stray workbook instead of appending; bail
                    # to the openpyxl fallback rather than ship a partial file.
                    raise RuntimeError(
                        f"Excel did not append {len(content)} sheet(s) from {skill!r}"
                    )
                # Name each freshly-appended sheet from its OWN identity (not by
                # position), so source/destination sheet ordering can't misassign.
                for i in range(added):
                    sheet = combined.Sheets(before + i + 1)
                    new_name = _unique_sheet_name(
                        _tab_name(skill, sheet.Name, len(content)), used_names
                    )
                    if sheet.Name != new_name:
                        sheet.Name = new_name
                    skill_to_tab.setdefault(skill, sheet.Name)
            finally:
                src.Close(SaveChanges=False)

        # Drop the blank workbook's seed sheets now that real content exists.
        for name in seed_sheets:
            if combined.Sheets.Count > 1:
                combined.Sheets(name).Delete()

        # Wire the combined workbook's cross-tab links (best-effort).
        _relink_cross_tab_com(combined, skill_to_tab)

        # Stamp the INFOR brand theme so the combined workbook keeps INFOR
        # colours/fonts even when the base is a blank workbook (no cap table).
        # Best-effort: a theme failure must never lose the merged workbook.
        if theme_path is not None:
            try:
                combined.ApplyTheme(str(theme_path))
            except Exception:
                pass

        if output_path.exists():
            output_path.unlink()
        combined.SaveAs(str(output_path), FileFormat=_XL_OPEN_XML_WORKBOOK)
        combined.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _combine_via_openpyxl(
    sources: list[tuple[str, Path]], output_path: Path, theme_path: Path | None = None
) -> None:
    """Best-effort merge with openpyxl. CapIQ links and charts do NOT survive."""
    from openpyxl import Workbook, load_workbook

    combined = Workbook()
    combined.remove(combined.active)  # drop the default empty sheet
    used_names: set[str] = set()
    skill_to_tab: dict[str, str] = {}

    for skill, path in sources:
        src = load_workbook(path, data_only=False)
        content_sheets = [n for n in src.sheetnames if not _is_capiq_helper_sheet(n)]
        for original in content_sheets:
            src_ws = src[original]
            title = _unique_sheet_name(
                _tab_name(skill, original, len(content_sheets)), used_names
            )
            dst_ws = combined.create_sheet(title=title)
            _copy_sheet(src_ws, dst_ws)
            skill_to_tab.setdefault(skill, title)

    if not combined.sheetnames:
        combined.create_sheet(title="Sheet")
    _relink_cross_tab_openpyxl(combined, skill_to_tab)

    # Stamp the INFOR brand theme (a fresh openpyxl Workbook carries the default
    # Office theme, so copied cells' theme-colour refs would otherwise resolve
    # against Office colours). Best-effort: skip silently if the theme is missing.
    if theme_path is not None:
        theme_xml = _extract_theme_xml(Path(theme_path))
        if theme_xml is not None:
            combined.loaded_theme = theme_xml

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
