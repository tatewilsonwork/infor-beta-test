"""Build a standalone LTM metrics workbook for the earnings update.

One "LTM Metrics" tab stacks three blocks, separated by a blank spacer row:

1. **LTM Revenue overview** — revenue split by segment, with `% of total` and a
   total row. Feeds the overview slide's deferred LTM revenue pie.
2. **LTM Revenue bridge** — how the LTM revenue total is derived:
   `LTM = FY + current-year YTD − prior-year YTD`.
3. **LTM Adj. EBITDA bridge** — the same FY + YTD − prior-YTD bridge for
   Adjusted EBITDA (or unadjusted EBITDA when no Adj. figure is disclosed).
   Bridge only — no segment overview.

Arithmetic lives in cell formulas (% of total, the total row, the bridge sums)
so the workbook stays analyst-auditable, matching the cap-table convention.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

# INFOR mid-blue header fill / Palatino body, mirroring the deck brand.
_HEADER_FILL = PatternFill("solid", fgColor="1F3864")
_SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
_HEADER_FONT = Font(name="Palatino Linotype", size=11, bold=True, color="FFFFFF")
_TITLE_FONT = Font(name="Palatino Linotype", size=14, bold=True, color="1F3864")
_SECTION_FONT = Font(name="Palatino Linotype", size=11, bold=True, color="1F3864")
_META_FONT = Font(name="Palatino Linotype", size=10, italic=True, color="595959")
_BODY_FONT = Font(name="Palatino Linotype", size=11)
_TOTAL_FONT = Font(name="Palatino Linotype", size=11, bold=True)
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_MINUS = "−"  # typographic minus sign used in bridge labels


@dataclass(frozen=True)
class RevenueSegment:
    """One row of the LTM revenue overview."""

    name: str
    ltm_revenue: float


@dataclass(frozen=True)
class BridgeComponent:
    """One additive (or subtractive) line of an LTM bridge."""

    name: str
    value: float
    subtract: bool = False


def _safe_name(value: str) -> str:
    safe = re.sub(r"[/\\:*?\"<>|]+", "-", value).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe or "Company"


def _coerce_segments(
    segments: list[RevenueSegment] | list[tuple],
) -> list[RevenueSegment]:
    return [
        seg if isinstance(seg, RevenueSegment) else RevenueSegment(seg[0], float(seg[1]))
        for seg in segments
    ]


def _coerce_components(
    components: list[BridgeComponent] | list[tuple] | None,
) -> list[BridgeComponent]:
    if not components:
        return []
    out: list[BridgeComponent] = []
    for c in components:
        if isinstance(c, BridgeComponent):
            out.append(c)
        else:
            name, value = c[0], float(c[1])
            subtract = bool(c[2]) if len(c) > 2 else False
            out.append(BridgeComponent(name, value, subtract))
    return out


def bridge_total(
    components: list[BridgeComponent] | list[tuple] | None,
) -> float | None:
    """Sum a bridge's components (additive minus subtractive) for the handoff.

    Returns None when no components are supplied. The workbook derives the same
    total via a cell formula; this mirrors it so the typed stage handoff can pass
    the LTM revenue / EBITDA figure to a downstream stage (e.g. the cap table).
    The figure is in the same currency as the bridge components — i.e. the
    filing's reporting currency.
    """
    comps = _coerce_components(components)
    if not comps:
        return None
    return sum(-c.value if c.subtract else c.value for c in comps)


def _section(ws: Worksheet, row: int, text: str) -> None:
    for col in (1, 2, 3):
        cell = ws.cell(row=row, column=col)
        cell.fill = _SECTION_FILL
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = _SECTION_FONT


def _write_bridge(
    ws: Worksheet,
    *,
    start_row: int,
    section_title: str,
    result_label: str,
    currency: str,
    components: list[BridgeComponent],
) -> int:
    """Write a bridge block and return the row after its result row."""
    _section(ws, start_row, section_title)
    header_row = start_row + 1
    for col, text in enumerate(("Component", f"Amount ({currency})"), start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = _BORDER
        cell.alignment = Alignment(horizontal="center" if col > 1 else "left")

    first_data = header_row + 1
    for i, comp in enumerate(components):
        r = first_data + i
        prefix = f"({_MINUS}) " if comp.subtract else "(+) "
        ws.cell(row=r, column=1, value=prefix + comp.name).font = _BODY_FONT
        v = ws.cell(row=r, column=2, value=comp.value)
        v.font = _BODY_FONT
        v.number_format = "#,##0.0"
        ws.cell(row=r, column=1).border = _BORDER
        v.border = _BORDER

    last_data = first_data + len(components) - 1
    result_row = last_data + 1

    terms = "".join(
        f"-B{first_data + i}" if comp.subtract else f"+B{first_data + i}"
        for i, comp in enumerate(components)
    ).lstrip("+")

    # The label is a plain string. A leading "=" makes openpyxl store the cell
    # as a formula (Excel then renders it as "=@LTM Revenue"), so use the "(=)"
    # bridge glyph, mirroring the "(+)" / "(−)" component rows above.
    ws.cell(row=result_row, column=1, value=f"(=) {result_label}").font = _TOTAL_FONT
    rv = ws.cell(row=result_row, column=2, value=f"={terms}")
    rv.font = _TOTAL_FONT
    rv.number_format = "#,##0.0"
    for col in (1, 2):
        ws.cell(row=result_row, column=col).border = _BORDER

    return result_row + 1


def build_ltm_metrics_workbook(
    *,
    company_name: str,
    period_label: str,
    currency: str,
    segmentation_basis: str,
    segments: list[RevenueSegment] | list[tuple],
    revenue_bridge: list[BridgeComponent] | list[tuple] | None = None,
    ebitda_bridge: list[BridgeComponent] | list[tuple] | None = None,
    ebitda_label: str = "LTM Adj. EBITDA",
    output_dir: Path | str,
    file_stem: str | None = None,
) -> Path:
    """Write an LTM metrics .xlsx (overview + bridges) and return its path.

    `segments` and the bridge components may be dataclasses or plain tuples.
    `segmentation_basis` is a human label such as "Service line" or "Geography".
    `ebitda_label` is "LTM Adj. EBITDA" by default; pass "LTM EBITDA" when no
    Adjusted figure is disclosed. Either bridge may be omitted.
    """
    rows = _coerce_segments(segments)
    if not rows:
        raise ValueError("LTM metrics workbook requires at least one revenue segment")
    rev_components = _coerce_components(revenue_bridge)
    ebitda_components = _coerce_components(ebitda_bridge)

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = file_stem or f"{_safe_name(company_name)} - LTM Metrics"
    output_path = out_dir / f"{stem}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "LTM Metrics"
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"{company_name} — LTM Metrics"
    ws["A1"].font = _TITLE_FONT
    ws["A2"] = f"Period: {period_label}"
    ws["A3"] = f"Revenue segmentation: {segmentation_basis}"
    ws["A4"] = f"Figures in {currency}"
    for cell in ("A2", "A3", "A4"):
        ws[cell].font = _META_FONT

    # --- LTM Revenue overview (row 5 left blank as a spacer after the meta block) ---
    section_row = 6
    _section(ws, section_row, "LTM Revenue Overview")
    header_row = section_row + 1
    for col, text in enumerate(("Segment", f"LTM Revenue ({currency})", "% of Total"), start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = _BORDER
        cell.alignment = Alignment(horizontal="center" if col > 1 else "left")

    first_data = header_row + 1
    for i, seg in enumerate(rows):
        r = first_data + i
        ws.cell(row=r, column=1, value=seg.name).font = _BODY_FONT
        v = ws.cell(row=r, column=2, value=seg.ltm_revenue)
        v.font = _BODY_FONT
        v.number_format = "#,##0.0"
        ws.cell(row=r, column=1).border = _BORDER
        v.border = _BORDER

    last_data = first_data + len(rows) - 1
    total_row = last_data + 1
    total_ref = f"B{total_row}"

    for i in range(len(rows)):
        r = first_data + i
        p = ws.cell(row=r, column=3, value=f"=B{r}/{total_ref}")
        p.font = _BODY_FONT
        p.number_format = "0.0%"
        p.border = _BORDER
        p.alignment = Alignment(horizontal="center")

    ws.cell(row=total_row, column=1, value="Total").font = _TOTAL_FONT
    tv = ws.cell(row=total_row, column=2, value=f"=SUM(B{first_data}:B{last_data})")
    tv.font = _TOTAL_FONT
    tv.number_format = "#,##0.0"
    tp = ws.cell(row=total_row, column=3, value=f"=SUM(C{first_data}:C{last_data})")
    tp.font = _TOTAL_FONT
    tp.number_format = "0.0%"
    tp.alignment = Alignment(horizontal="center")
    for col in (1, 2, 3):
        ws.cell(row=total_row, column=col).border = _BORDER

    # --- Bridges, each preceded by a blank spacer row ---
    next_row = total_row + 1
    if rev_components:
        next_row = _write_bridge(
            ws,
            start_row=next_row + 1,
            section_title="LTM Revenue Bridge",
            result_label="LTM Revenue",
            currency=currency,
            components=rev_components,
        )
    if ebitda_components:
        next_row = _write_bridge(
            ws,
            start_row=next_row + 1,
            section_title=f"{ebitda_label} Bridge",
            result_label=ebitda_label,
            currency=currency,
            components=ebitda_components,
        )

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 14

    wb.save(output_path)
    return output_path
