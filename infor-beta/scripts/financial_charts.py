"""Financial Summary charts — build the deck's four metric charts and place them.

Post-aggregation stage (`financial-charts`). The four metric charts are built on
the **combined** pitch workbook's `financial-summary` tab — the only place each
flow metric's ``=INDEX('ltm-metrics'!…)`` LTM link resolves, because the
`ltm-metrics` tab co-exists there after `workbook-aggregation` folds it in. (The
standalone Financial Summary file can't be charted: its LTM cells stay ``#N/A``
until aggregation.) The charts are then rendered and dropped into the Financial
Summary slide's four chart placeholders, stretched to each placeholder's box —
the same picture-into-placeholder pattern as the cap-table / ownership insertions.

Two backends mirror ``excel_to_powerpoint.py`` / ``slide_render.py``:

  - **Excel COM** (Windows + Excel) — native clustered-column charts are built on
    the tab and exported to PNG via ``Chart.Export``; a full recalc first resolves
    the LTM links. The charts persist on the tab and CapIQ links / other-tab
    formulas survive (COM does not round-trip the workbook through openpyxl).

  - **openpyxl + LibreOffice headless** (Cowork / Linux / macOS) — openpyxl writes
    the native charts on the tab; each PNG is rendered from a single-chart temp
    workbook built off LibreOffice-recalculated values. Best-effort fidelity.

INFOR chart formatting (the only formatting that matters): Palatino Linotype 9 pt
black (data labels + category-axis labels); no chart title; no chart border; the
category (horizontal) axis rendered as a **solid black baseline line of visible
width** (a hairline / width-less line is dropped by the LibreOffice PNG render, so
the width is explicit); no major gridlines; the value (vertical) axis hidden — no
line, no label; gap width 50%; data labels on every bar at Outside End; all bars
filled RGB(70, 86, 110) = hex ``46566E``.

The same module also builds the overview slide's **LTM revenue pie**
(``render_ltm_revenue_pie_into_deck``): a by-segment pie over the combined
workbook's ``ltm-metrics`` tab "LTM Revenue Overview" block. The pie series is the
**"% of Total" column** (the ``=B/Btotal`` fraction), so its data labels show the
segment share (value-only, ``#,##0.0%`` format) rather than the dollar amounts; the
legend at the TOP carries the segment names, no title/border, and slice fills from
the INFOR theme accent palette (``pptx_helpers.INFOR_ACCENTS``). It rides the same
post-aggregation stage and is dropped into the overview slide's "Rectangle 4"
placeholder.
"""

from __future__ import annotations

import io
import os
import shutil
import sys
import tempfile
from pathlib import Path

from pptx import Presentation

from pptx_helpers import INFOR_ACCENTS

# --- INFOR chart constants ---------------------------------------------------
_BAR_RGB_HEX = "46566E"  # RGB(70, 86, 110) as openpyxl RRGGBB
# Excel COM `.RGB` wants R + G*256 + B*65536 (i.e. 0x6E5646 = BGR byte order).
_BAR_RGB_COM = 70 + 86 * 256 + 110 * 65536
_FONT_NAME = "Palatino Linotype"
_FONT_SIZE_PT = 9
_FONT_SIZE_HUNDREDTHS = _FONT_SIZE_PT * 100  # openpyxl drawing fonts use 1/100 pt
_GAP_WIDTH = 50
_VALUE_FORMAT = "#,##0.0"

# Category (horizontal) axis baseline: an explicit, visible solid-black line.
# Without an explicit width the openpyxl ``<a:ln>`` defaults to a hairline that the
# LibreOffice PNG render drops entirely, leaving the bars with no baseline (Issue
# 2). 12700 EMU = 1.0 pt for the openpyxl path; the COM path takes the point value.
_AXIS_LINE_WIDTH_EMU = 12700
_AXIS_LINE_WEIGHT_PT = 1.0

# --- tab / slide geometry ----------------------------------------------------
_SHEET_DEFAULT = "financial-summary"
_HEADER_ROW = 5
# Financial Summary slide in the assembled pitch deck (the earnings slide at raw
# library index 7 is deleted, so the FS slide lands at slides[7]).
_SLIDE_INDEX_DEFAULT = 7
# Chart placeholder shape name -> the tab data row whose metric it charts. The
# 2x2 grid matches the slide tiles: #1 top-left, #2 top-right, #3 / #4 below.
_PLACEHOLDER_MAP: list[tuple[str, int]] = [
    ("Rectangle 17", 6),  # Metric #1 (label tile Rectangle 13)
    ("Rectangle 7", 7),   # Metric #2 (label tile Rectangle 12)
    ("Rectangle 19", 8),  # Metric #3 (label tile Rectangle 15)
    ("Rectangle 18", 9),  # Metric #4 (label tile Rectangle 14)
]
# Placeholder box: 4.53" x 2.51". Used to size the exported chart so its aspect
# matches the slide box (the picture is stretched to the box either way).
_CHART_W_PT = 4.53 * 72
_CHART_H_PT = 2.51 * 72
_CHART_W_CM = 4.53 * 2.54
_CHART_H_CM = 2.51 * 2.54

# --- LTM revenue pie (overview slide) ----------------------------------------
# The pie is built on the combined workbook's `ltm-metrics` tab — its "LTM Revenue
# Overview" block carries literal segment × LTM-revenue values — and dropped into
# the overview slide's wide/short "[Pie Chart Placeholder]" (Rectangle 4). The
# data live on the same combined workbook the FS charts use, so this rides the
# post-aggregation `financial-charts` stage rather than a parallel path.
_PIE_SHEET_DEFAULT = "ltm-metrics"
_PIE_SECTION_LABEL = "LTM Revenue Overview"
_PIE_TOTAL_LABEL = "Total"
_PIE_SEGMENT_COL = 1  # column A — segment names (categories / legend)
_PIE_VALUE_COL = 2    # column B — LTM revenue $ amounts (Python fraction-compute reads this)
_PIE_PCT_COL = 3      # column C — "% of Total" (=B/Btotal fraction); the chart series
# Data-label number format for the pie: the series is the column-C fraction, so a
# value-only "%" format renders e.g. 0.452 as "45.2%".
_PIE_LABEL_FORMAT = '#,##0.0%_);(#,##0.0%);"--"'
# Overview slide in the assembled pitch deck (slides[6] after the earnings slide
# at raw library index 7 is deleted).
_OVERVIEW_SLIDE_INDEX = 6
_PIE_PLACEHOLDER = "Rectangle 4"
# Placeholder box: 4.51" x 1.77" (wide and short) — size the exported pie to the
# same aspect so the stretched picture is not distorted.
_PIE_W_PT = 4.51 * 72
_PIE_H_PT = 1.77 * 72
_PIE_W_CM = 4.51 * 2.54
_PIE_H_CM = 1.77 * 2.54

# --- Excel COM enums ---------------------------------------------------------
_XL_NORMAL = -4143
_XL_COLUMN_CLUSTERED = 51
_XL_PIE = 5
_XL_CATEGORY = 1
_XL_VALUE = 2
_XL_LABEL_OUTSIDE_END = 2
_XL_LABEL_BEST_FIT = 5
_XL_LEGEND_TOP = -4160
_MSO_FALSE = 0
_MSO_TRUE = -1


def period_axis_columns(ws) -> tuple[int, int]:
    """Return the (first, last) 1-based column of the period axis on row 5.

    The period header runs ``B5 .. <col before "Units">``: ``B5:G5`` when the LTM
    column is shown, ``B5:F5`` when it is suppressed. Reading it dynamically lets
    a 5- or 6-column tab chart correctly. ``ws`` is an openpyxl worksheet.
    """
    units_col = None
    for col in range(2, ws.max_column + 1):
        if ws.cell(row=_HEADER_ROW, column=col).value == "Units":
            units_col = col
            break
    if units_col is None or units_col < 3:
        raise ValueError(
            "could not locate the 'Units' header on row 5; the Financial Summary "
            "tab is not in the expected chart-ready layout"
        )
    return 2, units_col - 1


def _find_label_row_openpyxl(ws, prefixes) -> int | None:
    """Row (1-based) of the first col-A cell whose text starts with any prefix.

    Mirrors ``workbook_aggregator._find_label_row_openpyxl`` — the established way
    to locate a labelled block on these tabs without hardcoding row numbers.
    """
    for row in ws.iter_rows(min_col=1, max_col=1):
        value = row[0].value
        if isinstance(value, str) and any(value.strip().startswith(p) for p in prefixes):
            return row[0].row
    return None


def ltm_revenue_overview_range(ws) -> tuple[int, int] | None:
    """Return the (first, last) 1-based data row of the "LTM Revenue Overview" block.

    The block runs: section title row, header row (section + 1), then one row per
    segment starting at section + 2, ending just above the "Total" row. Segment
    names live in column A and LTM revenue values in column B. The Total row is
    excluded. ``ws`` is an openpyxl worksheet. Returns ``None`` when the block (or
    its Total row) is not found.
    """
    section = _find_label_row_openpyxl(ws, (_PIE_SECTION_LABEL,))
    if section is None:
        return None
    first_data = section + 2  # skip the section title + the header row
    total = _find_label_row_openpyxl_from(ws, (_PIE_TOTAL_LABEL,), start=first_data)
    if total is None or total <= first_data:
        return None
    return first_data, total - 1


def _find_label_row_openpyxl_from(ws, prefixes, *, start: int) -> int | None:
    """Like ``_find_label_row_openpyxl`` but only scanning rows >= ``start``."""
    for row in range(start, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and any(value.strip().startswith(p) for p in prefixes):
            return row
    return None


def _find_label_row_com(ws, prefixes, *, start: int = 1) -> int | None:
    """COM counterpart of :func:`_find_label_row_openpyxl_from`."""
    try:
        used = ws.UsedRange
        last = used.Row + used.Rows.Count - 1
    except Exception:
        last = 200
    for r in range(start, last + 1):
        value = ws.Cells(r, 1).Value
        if isinstance(value, str) and any(value.strip().startswith(p) for p in prefixes):
            return r
    return None


def _hex_to_com_bgr(hex_rgb: str) -> int:
    """Convert an ``RRGGBB`` hex string to the int Excel COM ``.RGB`` expects
    (``R + G*256 + B*65536`` — i.e. BGR byte order)."""
    r = int(hex_rgb[0:2], 16)
    g = int(hex_rgb[2:4], 16)
    b = int(hex_rgb[4:6], 16)
    return r + g * 256 + b * 65536


# ---------------------------------------------------------------------------
# Public orchestrator
# ---------------------------------------------------------------------------
def render_financial_summary_charts_into_deck(
    *,
    deck_path: Path | str,
    combined_workbook_path: Path | str,
    sheet_name: str = _SHEET_DEFAULT,
    slide_index: int = _SLIDE_INDEX_DEFAULT,
    output_path: Path | str | None = None,
) -> Path | None:
    """Build the four Financial Summary charts and place them into the deck.

    Returns the output deck path, or ``None`` when the slide is left with its
    placeholders — either because the combined workbook has no ``financial-summary``
    tab (the financial-summary stage produced nothing) **or** because the native
    charts were persisted to the workbook but their PNGs could not be rendered for
    the deck (LibreOffice unavailable — the graceful-degradation path, Issue 1).
    In both cases the slide keeps its placeholders, like the ownership null path.
    """
    deck = Path(deck_path).resolve()
    workbook = Path(combined_workbook_path).resolve()
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    if not workbook.exists():
        raise FileNotFoundError(f"combined workbook not found: {workbook}")

    axis = _resolve_axis(workbook, sheet_name)
    if axis is None:
        return None
    first_col, last_col = axis

    if sys.platform == "win32":
        try:
            pngs = _build_charts_com(workbook, sheet_name, first_col, last_col)
        except RuntimeError:
            # Excel COM unavailable on this Windows box — fall through.
            pngs = _build_charts_openpyxl_libreoffice(workbook, sheet_name, first_col, last_col)
    else:
        pngs = _build_charts_openpyxl_libreoffice(workbook, sheet_name, first_col, last_col)

    if not pngs or any(row not in pngs for _name, row in _PLACEHOLDER_MAP):
        # Native charts were persisted to the workbook, but their PNGs could not be
        # rendered (LibreOffice unavailable). Leave the slide placeholders rather
        # than crashing — the durable artefact (workbook charts) is already saved.
        return None

    out = Path(output_path).resolve() if output_path is not None else deck
    return insert_pngs_into_placeholders(
        deck_path=deck,
        slide_index=slide_index,
        pngs_by_placeholder={name: pngs[row] for name, row in _PLACEHOLDER_MAP},
        output_path=out,
    )


def _resolve_axis(workbook: Path, sheet_name: str) -> tuple[int, int] | None:
    """Read the period-axis columns from the combined workbook, or None if the
    ``financial-summary`` tab is absent."""
    from openpyxl import load_workbook

    wb = load_workbook(workbook, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return None
        return period_axis_columns(wb[sheet_name])
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# LTM revenue pie (overview slide) — public orchestrator
# ---------------------------------------------------------------------------
def render_ltm_revenue_pie_into_deck(
    *,
    deck_path: Path | str,
    combined_workbook_path: Path | str,
    sheet_name: str = _PIE_SHEET_DEFAULT,
    slide_index: int = _OVERVIEW_SLIDE_INDEX,
    placeholder_name: str = _PIE_PLACEHOLDER,
    output_path: Path | str | None = None,
) -> Path | None:
    """Build the LTM-revenue-by-segment pie and place it on the overview slide.

    The pie is built on the combined workbook's ``ltm-metrics`` tab (its "LTM
    Revenue Overview" block carries literal segment × LTM-revenue values) and
    dropped into the overview slide's "[Pie Chart Placeholder]" (``Rectangle 4``).

    Returns the output deck path, or ``None`` when the slide keeps its placeholder
    — either because the combined workbook has no ``ltm-metrics`` tab / no "LTM
    Revenue Overview" block, **or** because the native pie was persisted to the
    workbook but its PNG could not be rendered for the deck (LibreOffice unavailable
    — the graceful-degradation path). Mirrors the FS / ownership null paths.
    """
    deck = Path(deck_path).resolve()
    workbook = Path(combined_workbook_path).resolve()
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    if not workbook.exists():
        raise FileNotFoundError(f"combined workbook not found: {workbook}")

    rng = _resolve_pie_range(workbook, sheet_name)
    if rng is None:
        return None
    first_row, last_row = rng

    if sys.platform == "win32":
        try:
            png = _build_pie_com(workbook, sheet_name, first_row, last_row)
        except RuntimeError:
            # Excel COM unavailable on this Windows box — fall through.
            png = _build_pie_openpyxl_libreoffice(workbook, sheet_name, first_row, last_row)
    else:
        png = _build_pie_openpyxl_libreoffice(workbook, sheet_name, first_row, last_row)

    if png is None:
        # Native pie persisted to the workbook, but its PNG could not be rendered
        # (LibreOffice unavailable). Leave the overview placeholder rather than
        # crashing — the durable artefact (workbook pie) is already saved.
        return None

    out = Path(output_path).resolve() if output_path is not None else deck
    return insert_pngs_into_placeholders(
        deck_path=deck,
        slide_index=slide_index,
        pngs_by_placeholder={placeholder_name: png},
        output_path=out,
    )


def _resolve_pie_range(workbook: Path, sheet_name: str) -> tuple[int, int] | None:
    """Read the "LTM Revenue Overview" data-row span from the combined workbook,
    or None if the ``ltm-metrics`` tab / block is absent."""
    from openpyxl import load_workbook

    wb = load_workbook(workbook, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return None
        return ltm_revenue_overview_range(wb[sheet_name])
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Placeholder insertion (mirrors excel_to_powerpoint.insert_excel_into_placeholder)
# ---------------------------------------------------------------------------
def insert_pngs_into_placeholders(
    *,
    deck_path: Path | str,
    slide_index: int,
    pngs_by_placeholder: dict[str, bytes],
    output_path: Path | str | None = None,
) -> Path:
    """Replace each named placeholder on a slide with a picture, in one pass.

    Each picture is added at the placeholder's exact left/top/width/height (the
    picture is stretched to the box), and the placeholder shape is removed.
    """
    deck = Path(deck_path).resolve()
    prs = Presentation(deck)
    slide = prs.slides[slide_index]
    for name, png in pngs_by_placeholder.items():
        placeholder = next((s for s in slide.shapes if s.name == name), None)
        if placeholder is None:
            raise KeyError(f"placeholder {name!r} not found on slide {slide_index + 1}")
        left, top, width, height = (
            placeholder.left,
            placeholder.top,
            placeholder.width,
            placeholder.height,
        )
        placeholder._element.getparent().remove(placeholder._element)
        buf = png if hasattr(png, "read") else io.BytesIO(png)
        slide.shapes.add_picture(buf, left, top, width=width, height=height)
    out = Path(output_path).resolve() if output_path is not None else deck
    prs.save(out)
    return out


def insert_png_into_placeholder(
    *,
    deck_path: Path | str,
    slide_index: int,
    placeholder_name: str,
    png_bytes: bytes,
    output_path: Path | str | None = None,
) -> Path:
    """Single-placeholder convenience wrapper over :func:`insert_pngs_into_placeholders`."""
    return insert_pngs_into_placeholders(
        deck_path=deck_path,
        slide_index=slide_index,
        pngs_by_placeholder={placeholder_name: png_bytes},
        output_path=output_path,
    )


# ---------------------------------------------------------------------------
# Excel COM backend (Windows) — full fidelity, charts persist, links survive
# ---------------------------------------------------------------------------
def _build_charts_com(
    workbook: Path, sheet_name: str, first_col: int, last_col: int
) -> dict[int, bytes]:
    """Build the four native charts on the tab via Excel COM and export each PNG."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for COM-based chart building (Windows + Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    tmp_paths: list[str] = []
    pngs: dict[int, bytes] = {}
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        # Chart.Export is reliable when the instance can render; run visible but
        # parked far off-screen so the window never pops in front of the analyst
        # (mirrors excel_to_powerpoint._excel_com_range_to_png).
        excel.Visible = True
        excel.DisplayAlerts = False
        try:
            excel.WindowState = _XL_NORMAL
            excel.Top, excel.Left = 4000, 6000
        except Exception:
            pass
        wb = excel.Workbooks.Open(str(workbook), ReadOnly=False, UpdateLinks=0)
        try:
            # Resolve the INDEX/MATCH LTM links (and any other formulas) before
            # snapshotting; a flaky recalc must never abort the build.
            try:
                excel.CalculateFull()
            except Exception:
                pass

            ws = wb.Worksheets(sheet_name)
            try:
                excel.ActiveWindow.ScrollRow = 1
                excel.ActiveWindow.ScrollColumn = 1
            except Exception:
                pass
            # Every chart is built + exported at the same on-screen "scratch" spot
            # (row 1) and only then moved to its 2x2 grid slot below the data. A
            # chart parked below the rendered viewport exports as a blank/0-byte
            # PNG, so exporting in-place would silently lose the lower-row charts.
            scratch_left = ws.Cells(1, 2).Left
            scratch_top = ws.Cells(1, 1).Top
            grid_top = ws.Cells(11, 1).Top
            for idx, (_placeholder, data_row) in enumerate(_PLACEHOLDER_MAP):
                chart_obj = ws.ChartObjects().Add(
                    Left=scratch_left, Top=scratch_top, Width=_CHART_W_PT, Height=_CHART_H_PT
                )
                chart = chart_obj.Chart
                chart.ChartType = _XL_COLUMN_CLUSTERED
                while chart.SeriesCollection().Count > 0:
                    chart.SeriesCollection(1).Delete()
                series = chart.SeriesCollection().NewSeries()
                series.Values = ws.Range(
                    ws.Cells(data_row, first_col), ws.Cells(data_row, last_col)
                )
                series.XValues = ws.Range(
                    ws.Cells(_HEADER_ROW, first_col), ws.Cells(_HEADER_ROW, last_col)
                )
                _format_com_chart(chart, series)

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                    tmp = f.name
                tmp_paths.append(tmp)
                chart.Export(Filename=tmp, FilterName="PNG")
                with open(tmp, "rb") as fh:
                    pngs[data_row] = fh.read()

                # Park the (now-exported) chart in its grid slot for persistence.
                col = idx % 2
                row = idx // 2
                chart_obj.Left = scratch_left + col * (_CHART_W_PT + 18)
                chart_obj.Top = grid_top + row * (_CHART_H_PT + 18)

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
        return pngs
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
        for tmp in tmp_paths:
            if os.path.exists(tmp):
                os.unlink(tmp)


def _com_strip_chart_border(chart) -> None:
    """Remove the chart-area and plot-area outlines so no border frames the picture.

    The legacy ``ChartArea.Border.LineStyle = 0`` alone does not suppress the
    modern chart-area outline in current Excel — that line is drawn by
    ``ChartArea.Format.Line``, which defaults visible. We clear both, plus the
    plot-area outline, defensively (a pie has no plot-area line to clear).
    """
    try:
        chart.ChartArea.Format.Line.Visible = _MSO_FALSE
    except Exception:
        pass
    try:
        chart.ChartArea.Border.LineStyle = 0
    except Exception:
        pass
    try:
        chart.PlotArea.Format.Line.Visible = _MSO_FALSE
    except Exception:
        pass


def _format_com_chart(chart, series) -> None:
    """Apply the INFOR chart formatting to a COM chart + its single series."""
    chart.HasTitle = False
    chart.HasLegend = False
    _com_strip_chart_border(chart)

    # Gap width 50% (clustered-column group).
    chart.ChartGroups(1).GapWidth = _GAP_WIDTH

    # All bars filled RGB(70, 86, 110).
    series.Format.Fill.Visible = True
    series.Format.Fill.ForeColor.RGB = _BAR_RGB_COM

    # Data labels on every bar, Outside End, Palatino 9 black.
    series.HasDataLabels = True
    labels = series.DataLabels()
    labels.Position = _XL_LABEL_OUTSIDE_END
    labels.NumberFormat = _VALUE_FORMAT
    labels.Font.Name = _FONT_NAME
    labels.Font.Size = _FONT_SIZE_PT
    labels.Font.Color = 0  # black

    # Category (horizontal) axis: Palatino 9 black, no title, a visible solid-black
    # baseline line (the default is a gray line — set it explicitly black with a
    # visible weight so the bars sit on a clear baseline).
    cat_axis = chart.Axes(_XL_CATEGORY)
    cat_axis.HasTitle = False
    cat_axis.TickLabels.Font.Name = _FONT_NAME
    cat_axis.TickLabels.Font.Size = _FONT_SIZE_PT
    cat_axis.TickLabels.Font.Color = 0
    cat_axis.Format.Line.Visible = _MSO_TRUE
    cat_axis.Format.Line.ForeColor.RGB = 0  # black
    try:
        cat_axis.Format.Line.Weight = _AXIS_LINE_WEIGHT_PT  # visible width (points)
    except Exception:
        pass

    # Value (vertical) axis + gridlines: hidden entirely.
    value_axis = chart.Axes(_XL_VALUE)
    value_axis.HasTitle = False
    value_axis.HasMajorGridlines = False
    value_axis.Delete()


def _build_pie_com(workbook: Path, sheet_name: str, first_row: int, last_row: int) -> bytes:
    """Build the LTM-revenue pie on the ``ltm-metrics`` tab via Excel COM, export PNG.

    Mirrors :func:`_build_charts_com`: the chart is built + exported at an on-screen
    scratch spot (a chart parked below the rendered viewport exports blank), then
    parked below the data and saved so it persists on the tab.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for COM-based chart building (Windows + Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    tmp_paths: list[str] = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        try:
            excel.WindowState = _XL_NORMAL
            excel.Top, excel.Left = 4000, 6000
        except Exception:
            pass
        wb = excel.Workbooks.Open(str(workbook), ReadOnly=False, UpdateLinks=0)
        try:
            try:
                excel.CalculateFull()
            except Exception:
                pass
            ws = wb.Worksheets(sheet_name)
            try:
                excel.ActiveWindow.ScrollRow = 1
                excel.ActiveWindow.ScrollColumn = 1
            except Exception:
                pass
            scratch_left = ws.Cells(1, 2).Left
            scratch_top = ws.Cells(1, 1).Top
            chart_obj = ws.ChartObjects().Add(
                Left=scratch_left, Top=scratch_top, Width=_PIE_W_PT, Height=_PIE_H_PT
            )
            chart = chart_obj.Chart
            chart.ChartType = _XL_PIE
            while chart.SeriesCollection().Count > 0:
                chart.SeriesCollection(1).Delete()
            series = chart.SeriesCollection().NewSeries()
            # Chart the column-C "% of Total" fraction (resolved by the CalculateFull
            # above), not the column-B $ amounts — the value labels then read the share.
            series.Values = ws.Range(
                ws.Cells(first_row, _PIE_PCT_COL), ws.Cells(last_row, _PIE_PCT_COL)
            )
            series.XValues = ws.Range(
                ws.Cells(first_row, _PIE_SEGMENT_COL), ws.Cells(last_row, _PIE_SEGMENT_COL)
            )
            _format_com_pie(chart, series, last_row - first_row + 1)

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                tmp = f.name
            tmp_paths.append(tmp)
            chart.Export(Filename=tmp, FilterName="PNG")
            with open(tmp, "rb") as fh:
                png = fh.read()

            # Park the (now-exported) chart below the data for persistence.
            try:
                chart_obj.Left = scratch_left
                chart_obj.Top = ws.Cells(last_row + 3, 1).Top
            except Exception:
                pass

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
        return png
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
        for tmp in tmp_paths:
            if os.path.exists(tmp):
                os.unlink(tmp)


def _format_com_pie(chart, series, n_points: int) -> None:
    """Apply INFOR pie formatting: legend at top, no title/border, accent slice fills."""
    chart.HasTitle = False
    _com_strip_chart_border(chart)

    # Legend at the TOP, Palatino 9 black.
    chart.HasLegend = True
    legend = chart.Legend
    legend.Position = _XL_LEGEND_TOP
    try:
        legend.Font.Name = _FONT_NAME
        legend.Font.Size = _FONT_SIZE_PT
        legend.Font.Color = 0
    except Exception:
        pass

    # Per-slice fills from the INFOR theme accents, in order, cycled past six.
    for i in range(1, n_points + 1):
        try:
            point = series.Points(i)
            point.Format.Fill.Visible = True
            point.Format.Fill.ForeColor.RGB = _hex_to_com_bgr(
                INFOR_ACCENTS[(i - 1) % len(INFOR_ACCENTS)]
            )
        except Exception:
            pass

    # Value-only data labels, Palatino 9. The series is the column-C "% of Total"
    # fraction, so ShowValue + a "%" number format renders e.g. 0.452 as "45.2%"
    # (ShowPercentage is off — we want the cell's own value, not a recomputed share).
    # Set the flags directly on the DataLabels object: under late-binding COM,
    # ApplyDataLabels(ShowValue=...) keyword args silently do nothing.
    try:
        series.HasDataLabels = True
        labels = series.DataLabels()
        labels.ShowValue = True
        labels.ShowPercentage = False
        labels.ShowCategoryName = False
        labels.ShowSeriesName = False
        labels.ShowLegendKey = False
        labels.NumberFormat = _PIE_LABEL_FORMAT
        labels.Position = _XL_LABEL_BEST_FIT
        labels.Font.Name = _FONT_NAME
        labels.Font.Size = _FONT_SIZE_PT
        labels.Font.Color = 0
    except Exception:
        pass


# ---------------------------------------------------------------------------
# openpyxl + LibreOffice backend (off-Windows) — best-effort
# ---------------------------------------------------------------------------
def _build_charts_openpyxl_libreoffice(
    workbook: Path, sheet_name: str, first_col: int, last_col: int
) -> dict[int, bytes]:
    """Persist native openpyxl charts on the tab, then render each PNG.

    The combined workbook off-Windows was already built by openpyxl (CapIQ links
    do not survive that path), so loading + re-saving it here costs nothing extra.
    PNGs are rendered from single-chart temp workbooks built off
    LibreOffice-recalculated values so the LTM bar is correct.

    Chart persistence on the tab must **not** depend on LibreOffice being present:
    the native charts are added and the workbook saved **first**, then the PNG
    render is attempted. If ``soffice``/``libreoffice`` (or ``pypdfium2``) is
    missing, the workbook charts have already been saved and an empty ``{}`` is
    returned so the caller degrades gracefully (leaves the deck placeholders)
    rather than aborting the whole stage (Issue 1).
    """
    from openpyxl import load_workbook

    wb = load_workbook(workbook)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"sheet {sheet_name!r} not found in {workbook}")
    ws = wb[sheet_name]
    anchors = ["B11", "I11", "B27", "I27"]
    for (_placeholder, data_row), anchor in zip(_PLACEHOLDER_MAP, anchors):
        ws.add_chart(_make_openpyxl_chart(ws, data_row, first_col, last_col), anchor)
    # Persist the native charts on the tab FIRST — independent of LibreOffice.
    wb.save(workbook)

    try:
        resolved = _libreoffice_recalc_values(workbook, sheet_name, first_col, last_col)
        labels = resolved["labels"]
        pngs: dict[int, bytes] = {}
        for _placeholder, data_row in _PLACEHOLDER_MAP:
            pngs[data_row] = _render_single_chart_png(labels, resolved[data_row])
        return pngs
    except RuntimeError as exc:
        # LibreOffice / pypdfium2 unavailable for the deck-image step. The native
        # charts are already saved on the tab; degrade gracefully instead of
        # aborting the stage — the caller leaves the slide placeholders.
        print(
            f"[financial-charts] deck-image render skipped ({exc}); the four native "
            f"charts are saved on the '{sheet_name}' tab of {workbook.name}.",
            file=sys.stderr,
        )
        return {}


def _make_openpyxl_chart(ws, data_row: int, first_col: int, last_col: int):
    """Build one INFOR-formatted clustered-column BarChart for a metric row."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.gapWidth = _GAP_WIDTH
    chart.title = None
    chart.legend = None
    chart.width = _CHART_W_CM
    chart.height = _CHART_H_CM

    data = Reference(ws, min_col=first_col, max_col=last_col, min_row=data_row, max_row=data_row)
    cats = Reference(ws, min_col=first_col, max_col=last_col, min_row=_HEADER_ROW, max_row=_HEADER_ROW)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)

    # All bars filled RGB(70, 86, 110).
    chart.series[0].graphicalProperties.solidFill = _BAR_RGB_HEX

    # Data labels on every bar, Outside End, Palatino 9 black.
    labels = DataLabelList()
    labels.showVal = True
    labels.dLblPos = "outEnd"
    labels.numFmt = _VALUE_FORMAT
    labels.txPr = _palatino_text()
    chart.dataLabels = labels

    # Category axis: Palatino 9 black; value axis + gridlines hidden.
    chart.x_axis.delete = False
    chart.x_axis.txPr = _palatino_text()
    chart.y_axis.delete = True
    chart.y_axis.majorGridlines = None
    _openpyxl_no_border_black_axis(chart)

    return chart


def _openpyxl_no_border_black_axis(chart) -> None:
    """No chart-area border + a visible solid-black category-axis baseline.

    openpyxl best-effort mirror of the COM border/axis fix: clears the chart-area
    outline and sets the category-axis line explicitly black **with a visible
    width**. A width-less ``<a:ln>`` renders as a hairline that LibreOffice drops
    on export, so the bars float with no baseline (Issue 2) — pin the width.
    """
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties

    chart.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.x_axis.spPr = GraphicalProperties(
        ln=LineProperties(w=_AXIS_LINE_WIDTH_EMU, solidFill="000000")
    )


def _palatino_text():
    """A RichText carrying Palatino Linotype 9 pt black default run properties."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (
        CharacterProperties,
        Font as DrawingFont,
        Paragraph,
        ParagraphProperties,
    )

    cp = CharacterProperties(
        latin=DrawingFont(typeface=_FONT_NAME),
        sz=_FONT_SIZE_HUNDREDTHS,
        solidFill="000000",
    )
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _libreoffice_recalc_values(
    workbook: Path, sheet_name: str, first_col: int, last_col: int
) -> dict:
    """Recalc the combined workbook with LibreOffice and read the period labels +
    each metric row's resolved values (so the LTM cell is no longer a formula)."""
    from openpyxl import load_workbook

    from excel_to_powerpoint import _soffice_convert  # reuse the recalc-on-load helper

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice/libreoffice) not found on PATH; required for the "
            "non-Windows Financial Summary chart renderer. Install LibreOffice or "
            "run the conductor on a Windows machine with Excel."
        )
    with tempfile.TemporaryDirectory() as tmp_dir:
        _soffice_convert(soffice, workbook, "xlsx:Calc MS Excel 2007 XML", Path(tmp_dir))
        recalced = Path(tmp_dir) / f"{workbook.stem}.xlsx"
        if not recalced.exists():
            raise RuntimeError("LibreOffice produced no recalculated workbook")
        wb = load_workbook(recalced, data_only=True)
        ws = wb[sheet_name]
        out: dict = {
            "labels": [
                ws.cell(row=_HEADER_ROW, column=c).value for c in range(first_col, last_col + 1)
            ]
        }
        for _placeholder, data_row in _PLACEHOLDER_MAP:
            out[data_row] = [
                ws.cell(row=data_row, column=c).value for c in range(first_col, last_col + 1)
            ]
        wb.close()
        return out


def _render_single_chart_png(labels: list, values: list) -> bytes:
    """Render one INFOR-formatted chart to PNG from literal values via LibreOffice.

    Builds a throwaway single-chart workbook (data parked above the print area,
    print area set to the chart footprint), converts it to PDF, and trims the
    white margins so only the chart survives.
    """
    from openpyxl import Workbook

    from excel_to_powerpoint import _soffice_convert, _trim_white_margins

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice/libreoffice) not found on PATH")
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError("pypdfium2 is required for the non-Windows chart renderer") from exc

    wb = Workbook()
    ws = wb.active
    n = len(labels)
    for j in range(n):
        ws.cell(row=1, column=2 + j, value=labels[j])
        ws.cell(row=2, column=2 + j, value=values[j])
    # Anchor the chart well below the data and print only its footprint, so the
    # data cells in rows 1-2 do not bleed into the rendered image.
    ws.add_chart(_make_single_value_chart(ws, n), "A5")
    ws.print_area = "A5:N28"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    for side in ("left", "right", "top", "bottom"):
        setattr(ws.page_margins, side, 0.1)

    with tempfile.TemporaryDirectory() as tmp_dir:
        src = Path(tmp_dir) / "chart.xlsx"
        wb.save(src)
        _soffice_convert(soffice, src, "pdf", Path(tmp_dir))
        pdf_path = Path(tmp_dir) / "chart.pdf"
        if not pdf_path.exists():
            raise RuntimeError("LibreOffice produced no PDF for the chart")
        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            pil = pdf[0].render(scale=200 / 72).to_pil().convert("RGB")
            pil = _trim_white_margins(pil)
            buf = io.BytesIO()
            pil.save(buf, format="PNG")
            return buf.getvalue()
        finally:
            pdf.close()


def _make_single_value_chart(ws, n: int):
    """A chart over a literal one-row data block (row 2, header row 1) at B..."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.gapWidth = _GAP_WIDTH
    chart.title = None
    chart.legend = None
    chart.width = _CHART_W_CM
    chart.height = _CHART_H_CM
    data = Reference(ws, min_col=2, max_col=1 + n, min_row=2, max_row=2)
    cats = Reference(ws, min_col=2, max_col=1 + n, min_row=1, max_row=1)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = _BAR_RGB_HEX
    labels = DataLabelList()
    labels.showVal = True
    labels.dLblPos = "outEnd"
    labels.numFmt = _VALUE_FORMAT
    labels.txPr = _palatino_text()
    chart.dataLabels = labels
    chart.x_axis.delete = False
    chart.x_axis.txPr = _palatino_text()
    chart.y_axis.delete = True
    chart.y_axis.majorGridlines = None
    _openpyxl_no_border_black_axis(chart)
    return chart


# ---------------------------------------------------------------------------
# openpyxl + LibreOffice backend for the LTM revenue pie — best-effort
# ---------------------------------------------------------------------------
def _fractions_from_amounts(amounts: list) -> list:
    """Convert a list of $ amounts to fractions of their total (None-safe).

    The off-Windows PNG render charts literal cells, but the pie's "% of Total"
    column on the real tab is the formula ``=B/Btotal`` that openpyxl cannot
    evaluate. So for the throwaway render workbook we recompute the same fractions
    in Python from the column-B ``$`` literals, so the rendered slices and value
    labels match the native pie's column-C series.
    """
    numeric = [a for a in amounts if isinstance(a, (int, float))]
    total = sum(numeric)
    if not total:
        return [None for _ in amounts]
    return [(a / total if isinstance(a, (int, float)) else None) for a in amounts]


def _build_pie_openpyxl_libreoffice(
    workbook: Path, sheet_name: str, first_row: int, last_row: int
) -> bytes | None:
    """Persist a native openpyxl pie on the ltm-metrics tab, then render its PNG.

    The native pie's series is the **"% of Total" column** (column C, a ``=B/Btotal``
    formula). openpyxl cannot evaluate that formula, so the throwaway render workbook
    charts the equivalent fractions recomputed in Python from the column-B ``$``
    literals (``frac = b / sum(b)``) — the slices and value labels then match the
    native pie without a LibreOffice recalc of the source tab.

    Like the FS chart path, the native pie is added and the workbook saved **first**
    so its persistence does not depend on LibreOffice; the PNG render is attempted
    after. If ``soffice``/``libreoffice`` (or ``pypdfium2``) is missing, the pie has
    already been saved on the tab and ``None`` is returned so the caller leaves the
    overview placeholder instead of aborting the stage (Issue 3 parity with Issue 1).
    """
    from openpyxl import load_workbook

    wb = load_workbook(workbook)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"sheet {sheet_name!r} not found in {workbook}")
    ws = wb[sheet_name]
    n = last_row - first_row + 1
    ws.add_chart(
        _make_openpyxl_pie(ws, first_row, last_row, n),
        ws.cell(row=last_row + 3, column=1).coordinate,
    )
    labels = [ws.cell(row=r, column=_PIE_SEGMENT_COL).value for r in range(first_row, last_row + 1)]
    amounts = [ws.cell(row=r, column=_PIE_VALUE_COL).value for r in range(first_row, last_row + 1)]
    # Persist the native pie on the tab FIRST — independent of LibreOffice.
    wb.save(workbook)

    try:
        # The render workbook charts literal cells, but the real tab's "% of Total"
        # column is a formula openpyxl can't evaluate — recompute the fractions from
        # the $ literals so the rendered slices/labels match the native pie's column-C
        # series.
        return _render_single_pie_png(labels, _fractions_from_amounts(amounts))
    except RuntimeError as exc:
        print(
            f"[financial-charts] overview pie render skipped ({exc}); the native "
            f"pie is saved on the '{sheet_name}' tab of {workbook.name}.",
            file=sys.stderr,
        )
        return None


def _style_openpyxl_pie(chart, n_points: int) -> None:
    """INFOR pie formatting: no title, accent slice fills, legend at top, value-only
    "%" labels (the series is the column-C "% of Total" fraction, so a value label
    with a "%" number format reads e.g. "45.2%")."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties

    chart.title = None
    chart.width = _PIE_W_CM
    chart.height = _PIE_H_CM
    chart.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=INFOR_ACCENTS[i % len(INFOR_ACCENTS)]))
        for i in range(n_points)
    ]
    chart.legend.position = "t"
    labels = DataLabelList()
    labels.showVal = True
    labels.showPercent = False
    labels.showCatName = False
    labels.showSerName = False
    labels.showLegendKey = False
    labels.numFmt = _PIE_LABEL_FORMAT
    labels.txPr = _palatino_text()
    chart.dataLabels = labels


def _make_openpyxl_pie(ws, first_row: int, last_row: int, n_points: int):
    """Build an INFOR-formatted PieChart over the "LTM Revenue Overview" block.

    The series is the **"% of Total" column** (column C, the ``=B/Btotal`` fraction)
    so the data labels read the segment share; Excel evaluates the formula. Slice
    geometry is identical to charting the $ amounts. Categories (legend) = the
    Segment column.
    """
    from openpyxl.chart import PieChart, Reference

    chart = PieChart()
    data = Reference(
        ws, min_col=_PIE_PCT_COL, max_col=_PIE_PCT_COL, min_row=first_row, max_row=last_row
    )
    cats = Reference(
        ws, min_col=_PIE_SEGMENT_COL, max_col=_PIE_SEGMENT_COL, min_row=first_row, max_row=last_row
    )
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    _style_openpyxl_pie(chart, n_points)
    return chart


def _make_single_pie_chart(ws, n: int):
    """A pie over a literal block: categories A1:A{n}, values B1:B{n}."""
    from openpyxl.chart import PieChart, Reference

    chart = PieChart()
    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=n)
    cats = Reference(ws, min_col=1, max_col=1, min_row=1, max_row=n)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    _style_openpyxl_pie(chart, n)
    return chart


def _render_single_pie_png(labels: list, values: list) -> bytes:
    """Render the INFOR pie to PNG from literal labels/values via LibreOffice.

    Mirrors :func:`_render_single_chart_png`: data parked above the print area, the
    print area set to the chart footprint, converted to PDF, white margins trimmed.
    """
    from openpyxl import Workbook

    from excel_to_powerpoint import _soffice_convert, _trim_white_margins

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice/libreoffice) not found on PATH")
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError("pypdfium2 is required for the non-Windows chart renderer") from exc

    wb = Workbook()
    ws = wb.active
    n = len(labels)
    for i in range(n):
        ws.cell(row=1 + i, column=1, value=labels[i])
        ws.cell(row=1 + i, column=2, value=values[i])
    # Anchor the chart away from the data and print only its footprint.
    ws.add_chart(_make_single_pie_chart(ws, n), "D2")
    ws.print_area = "D2:N28"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    for side in ("left", "right", "top", "bottom"):
        setattr(ws.page_margins, side, 0.1)

    with tempfile.TemporaryDirectory() as tmp_dir:
        src = Path(tmp_dir) / "pie.xlsx"
        wb.save(src)
        _soffice_convert(soffice, src, "pdf", Path(tmp_dir))
        pdf_path = Path(tmp_dir) / "pie.pdf"
        if not pdf_path.exists():
            raise RuntimeError("LibreOffice produced no PDF for the pie chart")
        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            pil = pdf[0].render(scale=200 / 72).to_pil().convert("RGB")
            pil = _trim_white_margins(pil)
            buf = io.BytesIO()
            pil.save(buf, format="PNG")
            return buf.getvalue()
        finally:
            pdf.close()
