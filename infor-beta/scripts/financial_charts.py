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
black (data labels + category-axis labels); no chart title; no major gridlines;
the value (vertical) axis hidden — no line, no label; gap width 50%; data labels
on every bar at Outside End; all bars filled RGB(70, 86, 110) = hex ``46566E``.
"""

from __future__ import annotations

import io
import os
import shutil
import sys
import tempfile
from pathlib import Path

from pptx import Presentation

# --- INFOR chart constants ---------------------------------------------------
_BAR_RGB_HEX = "46566E"  # RGB(70, 86, 110) as openpyxl RRGGBB
# Excel COM `.RGB` wants R + G*256 + B*65536 (i.e. 0x6E5646 = BGR byte order).
_BAR_RGB_COM = 70 + 86 * 256 + 110 * 65536
_FONT_NAME = "Palatino Linotype"
_FONT_SIZE_PT = 9
_FONT_SIZE_HUNDREDTHS = _FONT_SIZE_PT * 100  # openpyxl drawing fonts use 1/100 pt
_GAP_WIDTH = 50
_VALUE_FORMAT = "#,##0.0"

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

# --- Excel COM enums ---------------------------------------------------------
_XL_NORMAL = -4143
_XL_COLUMN_CLUSTERED = 51
_XL_CATEGORY = 1
_XL_VALUE = 2
_XL_LABEL_OUTSIDE_END = 2


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

    Returns the output deck path, or ``None`` when the combined workbook has no
    ``financial-summary`` tab (e.g. the financial-summary stage produced nothing)
    — in that case the slide is left with its placeholders, like the ownership
    null path.
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


def _format_com_chart(chart, series) -> None:
    """Apply the INFOR chart formatting to a COM chart + its single series."""
    chart.HasTitle = False
    chart.HasLegend = False
    chart.ChartArea.Border.LineStyle = 0  # no border box around the picture

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

    # Category (horizontal) axis: Palatino 9 black, no title.
    cat_axis = chart.Axes(_XL_CATEGORY)
    cat_axis.HasTitle = False
    cat_axis.TickLabels.Font.Name = _FONT_NAME
    cat_axis.TickLabels.Font.Size = _FONT_SIZE_PT
    cat_axis.TickLabels.Font.Color = 0

    # Value (vertical) axis + gridlines: hidden entirely.
    value_axis = chart.Axes(_XL_VALUE)
    value_axis.HasTitle = False
    value_axis.HasMajorGridlines = False
    value_axis.Delete()


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
    """
    from openpyxl import load_workbook

    wb = load_workbook(workbook)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"sheet {sheet_name!r} not found in {workbook}")
    ws = wb[sheet_name]
    anchors = ["B11", "I11", "B27", "I27"]
    for (_placeholder, data_row), anchor in zip(_PLACEHOLDER_MAP, anchors):
        ws.add_chart(_make_openpyxl_chart(ws, data_row, first_col, last_col), anchor)
    wb.save(workbook)

    resolved = _libreoffice_recalc_values(workbook, sheet_name, first_col, last_col)
    labels = resolved["labels"]
    pngs: dict[int, bytes] = {}
    for _placeholder, data_row in _PLACEHOLDER_MAP:
        pngs[data_row] = _render_single_chart_png(labels, resolved[data_row])
    return pngs


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

    return chart


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
    return chart
