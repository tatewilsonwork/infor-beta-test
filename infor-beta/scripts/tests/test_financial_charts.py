"""Unit tests for the Financial Summary chart builder + placeholder insertion.

These cover the cross-platform pieces — period-axis detection, the openpyxl
chart formatting, and PNG-into-placeholder insertion. The Excel COM and
LibreOffice render backends are environmental (per repo convention) and are
exercised by the skill's mandatory QA render, not here.
"""

from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from financial_summary_workbook import MetricSeries, build_financial_summary_workbook
from financial_charts import (
    _make_openpyxl_chart,
    insert_png_into_placeholder,
    period_axis_columns,
)

# A valid 1x1 transparent PNG — avoids a PIL dependency in the test.
_PNG_1X1 = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06"
    b"\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00"
    b"\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
)

_FISCAL = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]


def _metrics() -> list[MetricSeries]:
    return [
        MetricSeries("Revenue", "US$MM", [3100.0, 3450.0, 3820.0, 4180.0, 4520.0],
                     result_label="LTM Revenue"),
        MetricSeries("Gross Profit", "US$MM", [1240.0, 1410.0, 1570.0, 1740.0, 1900.0],
                     result_label="LTM Gross Profit"),
        MetricSeries("Adjusted EBITDA", "US$MM", [820.0, 940.0, 1080.0, 1210.0, 1330.0],
                     result_label="LTM Adj. EBITDA"),
        MetricSeries("Net Income", "US$MM", [410.0, 470.0, 540.0, 610.0, 680.0],
                     result_label="LTM Net Income"),
    ]


def _fs_workbook(tmp_path: Path, **overrides) -> Path:
    kwargs = dict(
        company_name="SampleCo",
        currency_note="Figures in US$MM unless noted",
        period_note="FY = fiscal year; LTM = trailing twelve months as of Q3 2026",
        fiscal_labels=_FISCAL,
        metrics=_metrics(),
        output_dir=tmp_path,
    )
    kwargs.update(overrides)
    return build_financial_summary_workbook(**kwargs)


def test_period_axis_columns_with_ltm(tmp_path: Path):
    ws = load_workbook(_fs_workbook(tmp_path)).active
    # Metric | FY1..FY5 | LTM | Units -> period axis B..G.
    assert period_axis_columns(ws) == (2, 7)


def test_period_axis_columns_suppressed_ltm(tmp_path: Path):
    ws = load_workbook(_fs_workbook(tmp_path, show_ltm=False)).active
    # Metric | FY1..FY5 | Units -> period axis B..F.
    assert period_axis_columns(ws) == (2, 6)


def test_make_openpyxl_chart_applies_infor_formatting(tmp_path: Path):
    ws = load_workbook(_fs_workbook(tmp_path)).active
    chart = _make_openpyxl_chart(ws, data_row=6, first_col=2, last_col=7)

    assert chart.type == "col"
    assert chart.grouping == "clustered"
    assert chart.gapWidth == 50
    assert chart.title is None
    assert chart.legend is None
    # One series, filled INFOR blue (openpyxl wraps a hex string in a ColorChoice).
    assert len(chart.series) == 1
    fill = chart.series[0].graphicalProperties.solidFill
    assert getattr(fill, "srgbClr", fill) == "46566E"
    # Data labels on every bar, Outside End.
    assert chart.dataLabels.showVal is True
    assert chart.dataLabels.dLblPos == "outEnd"
    # Value axis hidden (no line/label), no gridlines; category axis kept.
    assert chart.y_axis.delete is True
    assert chart.y_axis.majorGridlines is None
    assert chart.x_axis.delete is False


def test_add_openpyxl_charts_yields_four_charts(tmp_path: Path):
    # Building one chart per metric row produces four independent charts.
    wb = load_workbook(_fs_workbook(tmp_path))
    ws = wb.active
    charts = [_make_openpyxl_chart(ws, data_row=r, first_col=2, last_col=7) for r in (6, 7, 8, 9)]
    assert len(charts) == 4
    assert all(c.gapWidth == 50 and c.y_axis.delete is True for c in charts)


def _deck_with_placeholder(tmp_path: Path) -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    box = slide.shapes.add_textbox(Inches(0.35), Inches(1.51), Inches(4.53), Inches(2.51))
    box.name = "Rectangle 17"
    box.text_frame.text = "[Placeholder for Metric #1 Chart]"
    path = tmp_path / "deck.pptx"
    prs.save(path)
    return path


def test_insert_png_replaces_placeholder_with_picture(tmp_path: Path):
    deck = _deck_with_placeholder(tmp_path)
    out = insert_png_into_placeholder(
        deck_path=deck,
        slide_index=0,
        placeholder_name="Rectangle 17",
        png_bytes=_PNG_1X1,
        output_path=deck,
    )
    prs = Presentation(out)
    slide = prs.slides[0]
    names = [s.name for s in slide.shapes]
    assert "Rectangle 17" not in names
    pictures = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(pictures) == 1
    pic = pictures[0]
    # Picture lands at the placeholder's exact box.
    assert pic.left == Inches(0.35)
    assert pic.top == Inches(1.51)
    assert pic.width == Inches(4.53)
    assert pic.height == Inches(2.51)


def test_insert_missing_placeholder_raises(tmp_path: Path):
    deck = _deck_with_placeholder(tmp_path)
    import pytest

    with pytest.raises(KeyError):
        insert_png_into_placeholder(
            deck_path=deck,
            slide_index=0,
            placeholder_name="Rectangle 999",
            png_bytes=_PNG_1X1,
            output_path=deck,
        )
