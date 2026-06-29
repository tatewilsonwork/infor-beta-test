"""Build the standalone Financial Summary data tab for the pitch deck.

One **"Financial Summary"** tab holds the data behind the deck's Financial Summary
slide (slide-library entry 8): the **4 metrics the `financial-summary` skill selected**,
each with its **last 5 fiscal years** plus an **LTM** column. The layout is deliberately
*chart-ready* so a later step can drop native Excel charts on top with no reshaping.

Chart-ready layout contract (rely on this from the future chart step)::

    A1 : "<company> — Financial Summary"                         (title)
    A2 : "<currency note>"  e.g. "Figures in US$MM unless noted"  (meta)
    A3 : "<period note>"    e.g. "FY = fiscal year; LTM = trailing twelve months as of Q3 2026"
    row 4: (blank spacer)
    row 5 (HEADER — a single contiguous period axis):
        A5 = "Metric"   B5..F5 = the 5 fiscal-year labels, oldest -> newest
        G5 = "LTM"      (only when the LTM column is shown — see suppression below)
        <last col> = "Units"
    rows 6-9 (DATA — one metric per row, four rows):
        A = metric label (identical to the deck tile label)
        B..F = five NUMERIC fiscal-year values, chronological
        LTM cell:
            * flow metric  -> =INDEX('ltm-metrics'!$B:$B, MATCH("(=) <result_label>", 'ltm-metrics'!$A:$A, 0))
              (a label-keyed link into the post-aggregation `ltm-metrics` tab; #N/A in
              the standalone file, resolves in the combined pitch workbook — exactly like
              the cap table's CapIQ formulas)
            * non-flow metric (balance / ratio that has no LTM bridge) -> the latest
              reported value, written as a literal number (point-in-time fallback)
        <last col> = the metric's units string (e.g. "US$MM", "%")

LTM suppression: when the most recent filing is a fiscal year-end 10-K with no later
interim 10-Q stub, LTM == the latest fiscal year. The caller passes ``show_ltm=False``;
the LTM column is dropped and, for flow metrics, the **most-recent fiscal-year cell**
carries the `ltm-metrics` link instead (req 5).

The data block (B5 down through the last metric row, across every period column) carries
**no merged cells** and **numeric** value cells, so the chart step can select the header
row + one metric row and chart it directly. Arithmetic / LTM derivation lives on the
`ltm-metrics` tab and is *linked* here — Excel does the math, not the LLM.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from naming import safe_filename

# INFOR mid-blue header fill / Palatino body, mirroring `ltm_metrics.py` so the
# folded-in tab matches the rest of the combined workbook.
_HEADER_FILL = PatternFill("solid", fgColor="1F3864")
_HEADER_FONT = Font(name="Palatino Linotype", size=11, bold=True, color="FFFFFF")
_TITLE_FONT = Font(name="Palatino Linotype", size=14, bold=True, color="1F3864")
_META_FONT = Font(name="Palatino Linotype", size=10, italic=True, color="595959")
_BODY_FONT = Font(name="Palatino Linotype", size=11)
_LABEL_FONT = Font(name="Palatino Linotype", size=11, bold=True)
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_SHEET_TITLE = "Financial Summary"
# Tab name the `ltm-metrics` workbook takes inside the combined pitch workbook
# (the aggregator renames a single-sheet source to its skill key). The LTM link
# must target this post-aggregation name, not the standalone "LTM Metrics" sheet.
_DEFAULT_LTM_SHEET = "ltm-metrics"

_METRIC_COUNT = 4  # the deck shows exactly four tiles
_VALUE_FORMAT = "#,##0.0"


@dataclass(frozen=True)
class MetricSeries:
    """One metric row of the Financial Summary tab.

    - ``label``: the metric NAME (e.g. ``"Revenue"``). Doubles as the deck tile
      label and column A. No amounts / currency tokens.
    - ``units``: the unit string for this metric's values (e.g. ``"US$MM"``,
      ``"%"``). Constant down the row, shown in the Units column.
    - ``fiscal_values``: the five fiscal-year values, **chronological (oldest ->
      newest)**; length must match the workbook's ``fiscal_labels``.
    - ``result_label``: for a **flow** metric, the exact `ltm-metrics` bridge
      result label (e.g. ``"LTM Revenue"``) — the LTM cell links to its
      ``(=) <result_label>`` row. ``None`` for a non-flow metric.
    - ``ltm_value``: for a **non-flow** metric (``result_label is None``), the
      latest reported value used as the point-in-time "LTM" figure. Ignored when
      ``result_label`` is set.
    """

    label: str
    units: str
    fiscal_values: list[float] = field(default_factory=list)
    result_label: str | None = None
    ltm_value: float | None = None


def _normalize_metric(metric: "MetricSeries | dict") -> MetricSeries:
    if isinstance(metric, MetricSeries):
        return metric
    return MetricSeries(
        label=metric["label"],
        units=metric.get("units", ""),
        fiscal_values=[float(v) for v in metric.get("fiscal_values", [])],
        result_label=metric.get("result_label"),
        ltm_value=(None if metric.get("ltm_value") is None else float(metric["ltm_value"])),
    )


def _quote_sheet(name: str) -> str:
    """Single-quote a sheet name for use in a formula (handles the hyphen in
    ``ltm-metrics`` and any embedded apostrophe)."""
    return "'" + name.replace("'", "''") + "'"


def _ltm_link(ltm_sheet: str, result_label: str) -> str:
    """A label-keyed lookup of a bridge total on the post-aggregation LTM tab.

    Resolves in the combined workbook (where the ``ltm-metrics`` tab exists);
    surfaces ``#N/A`` / ``#REF!`` in the standalone file, like the cap table's
    CapIQ formulas. Keyed on the bridge's ``(=) <result_label>`` row so it
    survives the bridge's variable row position.
    """
    sheet = _quote_sheet(ltm_sheet)
    key = f"(=) {result_label}"
    return f'=INDEX({sheet}!$B:$B, MATCH("{key}", {sheet}!$A:$A, 0))'


def build_financial_summary_workbook(
    *,
    company_name: str,
    currency_note: str,
    period_note: str,
    fiscal_labels: list[str],
    metrics: "list[MetricSeries | dict]",
    show_ltm: bool = True,
    ltm_sheet_name: str = _DEFAULT_LTM_SHEET,
    output_dir: Path | str,
    file_stem: str | None = None,
) -> Path:
    """Write the chart-ready Financial Summary .xlsx and return its path.

    ``fiscal_labels`` are the five fiscal-year column headers in chronological
    order (oldest -> newest). ``metrics`` must hold exactly four ``MetricSeries``
    (or dicts), each supplying one value per fiscal label. When ``show_ltm`` is
    True an LTM column is appended; flow metrics (those with a ``result_label``)
    link to the ``ltm-metrics`` tab and non-flow metrics show their ``ltm_value``.
    When ``show_ltm`` is False the LTM column is dropped and a flow metric's
    most-recent fiscal-year cell carries the LTM-tab link instead (req 5).

    Raises ValueError on the wrong metric count, a fiscal-value count that does
    not match ``fiscal_labels``, or a flow metric missing its ``result_label``
    while a non-flow metric has neither a link nor an ``ltm_value``.
    """
    rows = [_normalize_metric(m) for m in metrics]
    if len(rows) != _METRIC_COUNT:
        raise ValueError(
            f"Financial Summary holds exactly {_METRIC_COUNT} metrics; got {len(rows)}"
        )
    if not fiscal_labels:
        raise ValueError("at least one fiscal-year label is required")
    n_fy = len(fiscal_labels)
    for m in rows:
        if len(m.fiscal_values) != n_fy:
            raise ValueError(
                f"metric {m.label!r} has {len(m.fiscal_values)} fiscal values; "
                f"expected {n_fy} to match the fiscal-year labels"
            )
        if not m.label.strip():
            raise ValueError("metric label cannot be blank")
        if m.result_label is None and m.ltm_value is None and show_ltm:
            raise ValueError(
                f"non-flow metric {m.label!r} needs an ltm_value (it has no "
                f"result_label to link to the ltm-metrics tab)"
            )

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = file_stem or f"{safe_filename(company_name, default='Company')} - Financial Summary"
    output_path = out_dir / f"{stem}.xlsx"

    wb = Workbook()
    ws: Worksheet = wb.active
    ws.title = _SHEET_TITLE
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"{company_name} — Financial Summary"
    ws["A1"].font = _TITLE_FONT
    ws["A2"] = currency_note
    ws["A3"] = period_note
    for cell in ("A2", "A3"):
        ws[cell].font = _META_FONT

    # --- Header row 5: Metric | <FY labels> | [LTM] | Units ------------------
    header_row = 5
    first_data = header_row + 1

    period_labels = list(fiscal_labels) + (["LTM"] if show_ltm else [])
    headers = ["Metric", *period_labels, "Units"]
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = _BORDER
        cell.alignment = Alignment(horizontal="center" if col > 1 else "left")

    units_col = 1 + len(period_labels) + 1  # Metric + periods + Units
    # The most-recent fiscal-year column (last FY, before any LTM column).
    last_fy_col = 1 + n_fy
    ltm_col = last_fy_col + 1 if show_ltm else None

    # --- Data rows: one metric per row ---------------------------------------
    for i, m in enumerate(rows):
        r = first_data + i
        label_cell = ws.cell(row=r, column=1, value=m.label)
        label_cell.font = _LABEL_FONT
        label_cell.border = _BORDER

        for j, value in enumerate(m.fiscal_values):
            c = 2 + j
            cell = ws.cell(row=r, column=c, value=value)
            cell.font = _BODY_FONT
            cell.number_format = _VALUE_FORMAT
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal="center")

        if show_ltm:
            cell = ws.cell(row=r, column=ltm_col)
            if m.result_label is not None:
                cell.value = _ltm_link(ltm_sheet_name, m.result_label)
            else:
                cell.value = m.ltm_value
            cell.font = _LABEL_FONT
            cell.number_format = _VALUE_FORMAT
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal="center")
        elif m.result_label is not None:
            # Suppression: LTM == latest FY, so the most-recent FY cell carries
            # the link to the LTM tab instead of the literal value (req 5).
            cell = ws.cell(row=r, column=last_fy_col)
            cell.value = _ltm_link(ltm_sheet_name, m.result_label)
            cell.number_format = _VALUE_FORMAT

        units_cell = ws.cell(row=r, column=units_col, value=m.units)
        units_cell.font = _META_FONT
        units_cell.border = _BORDER
        units_cell.alignment = Alignment(horizontal="center")

    # Column widths: metric label wide, period/units columns even.
    ws.column_dimensions["A"].width = 34
    for col in range(2, units_col):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(units_col)].width = 10

    wb.save(output_path)
    return output_path
