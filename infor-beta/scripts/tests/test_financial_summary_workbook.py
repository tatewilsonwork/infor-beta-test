"""Unit tests for the Financial Summary workbook helper."""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from financial_summary_workbook import MetricSeries, build_financial_summary_workbook

_FISCAL = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]


def _metrics() -> list[MetricSeries]:
    return [
        MetricSeries("Revenue", "US$MM", [3100.0, 3450.0, 3820.0, 4180.0, 4520.0],
                     result_label="LTM Revenue"),
        MetricSeries("Gross Profit", "US$MM", [1240.0, 1410.0, 1570.0, 1740.0, 1900.0],
                     result_label="LTM Gross Profit"),
        MetricSeries("Adjusted EBITDA", "US$MM", [820.0, 940.0, 1080.0, 1210.0, 1330.0],
                     result_label="LTM Adj. EBITDA"),
        # Non-flow metric: point-in-time latest value, no ltm-metrics bridge.
        MetricSeries("Combined Loan Balances", "US$MM", [9000.0, 10000.0, 11000.0, 12000.0, 12500.0],
                     ltm_value=12500.0),
    ]


def _build(tmp_path: Path, **overrides) -> Path:
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


def test_writes_expected_file_and_sheet(tmp_path: Path):
    path = _build(tmp_path)
    assert path.exists()
    assert path.name == "SampleCo - Financial Summary.xlsx"
    assert load_workbook(path).active.title == "Financial Summary"


def test_header_row_is_contiguous_period_axis(tmp_path: Path):
    ws = load_workbook(_build(tmp_path)).active
    assert ws["A5"].value == "Metric"
    # Five fiscal years oldest -> newest in B..F, LTM in G, Units in H.
    assert [ws.cell(row=5, column=c).value for c in range(2, 7)] == _FISCAL
    assert ws["G5"].value == "LTM"
    assert ws["H5"].value == "Units"


def test_metric_rows_are_numeric_with_units(tmp_path: Path):
    ws = load_workbook(_build(tmp_path)).active
    assert ws["A6"].value == "Revenue"
    assert [ws.cell(row=6, column=c).value for c in range(2, 7)] == [3100.0, 3450.0, 3820.0, 4180.0, 4520.0]
    assert all(isinstance(ws.cell(row=6, column=c).value, (int, float)) for c in range(2, 7))
    assert ws["H6"].value == "US$MM"


def test_flow_metric_ltm_links_label_keyed_to_ltm_metrics_tab(tmp_path: Path):
    ws = load_workbook(_build(tmp_path)).active
    # The LTM cell is a label-keyed lookup into the post-aggregation 'ltm-metrics' tab.
    assert ws["G6"].value == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", 'ltm-metrics'!$A:$A, 0))"
    )
    assert ws["G8"].value == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Adj. EBITDA\", 'ltm-metrics'!$A:$A, 0))"
    )


def test_non_flow_metric_ltm_is_literal_value(tmp_path: Path):
    ws = load_workbook(_build(tmp_path)).active
    # Combined Loan Balances is the 4th metric -> row 9; its LTM is a point-in-time literal.
    assert ws["A9"].value == "Combined Loan Balances"
    assert ws["G9"].value == 12500.0


def test_suppression_drops_ltm_column_and_links_latest_fy(tmp_path: Path):
    ws = load_workbook(_build(tmp_path, show_ltm=False)).active
    # No LTM column: header runs Metric | FY1..FY5 | Units (Units lands in G).
    assert ws["G5"].value == "Units"
    row5 = [ws.cell(row=5, column=c).value for c in range(1, 8)]
    assert "LTM" not in row5
    # Flow metric: the most-recent FY cell (F) carries the ltm-metrics link.
    assert ws["F6"].value == (
        "=INDEX('ltm-metrics'!$B:$B, MATCH(\"(=) LTM Revenue\", 'ltm-metrics'!$A:$A, 0))"
    )
    # Non-flow metric keeps its literal most-recent FY value.
    assert ws["F9"].value == 12500.0


def test_value_cells_use_currency_number_format(tmp_path: Path):
    # v0.5.19: metric value cells carry the "$#,##0.0" currency format (FY values,
    # the flow-metric LTM link cell, and the non-flow LTM literal).
    ws = load_workbook(_build(tmp_path)).active
    expected = '$#,##0.0_);($#,##0.0);"--"'
    for c in range(2, 7):  # FY values B..F on the Revenue row
        assert ws.cell(row=6, column=c).number_format == expected
    assert ws["G6"].number_format == expected  # flow-metric LTM link cell
    assert ws["G9"].number_format == expected  # non-flow LTM literal


def test_no_merged_cells_in_data_block(tmp_path: Path):
    ws = load_workbook(_build(tmp_path)).active
    assert list(ws.merged_cells.ranges) == []


def test_accepts_plain_dicts(tmp_path: Path):
    path = _build(
        tmp_path,
        metrics=[
            {"label": "Revenue", "units": "US$MM", "fiscal_values": [1, 2, 3, 4, 5], "result_label": "LTM Revenue"},
            {"label": "Gross Profit", "units": "US$MM", "fiscal_values": [1, 2, 3, 4, 5], "result_label": "LTM Gross Profit"},
            {"label": "Adjusted EBITDA", "units": "US$MM", "fiscal_values": [1, 2, 3, 4, 5], "result_label": "LTM Adj. EBITDA"},
            {"label": "Net Income", "units": "US$MM", "fiscal_values": [1, 2, 3, 4, 5], "result_label": "LTM Net Income"},
        ],
    )
    ws = load_workbook(path).active
    assert ws["A6"].value == "Revenue"


def test_wrong_metric_count_raises(tmp_path: Path):
    with pytest.raises(ValueError):
        _build(tmp_path, metrics=_metrics()[:3])


def test_fiscal_value_count_mismatch_raises(tmp_path: Path):
    bad = _metrics()
    bad[0] = MetricSeries("Revenue", "US$MM", [1.0, 2.0, 3.0], result_label="LTM Revenue")
    with pytest.raises(ValueError):
        _build(tmp_path, metrics=bad)


def test_non_flow_without_ltm_value_raises_when_ltm_shown(tmp_path: Path):
    bad = _metrics()
    bad[3] = MetricSeries("Combined Loan Balances", "US$MM", [1.0, 2.0, 3.0, 4.0, 5.0])
    with pytest.raises(ValueError):
        _build(tmp_path, metrics=bad)


def test_combined_metric_fiscal_value_written_as_formula(tmp_path: Path):
    # A combined metric (loans + advances) passes each FY value as a formula of
    # its reported components, never pre-summed; openpyxl stores it as a formula.
    metrics = _metrics()
    metrics[3] = MetricSeries(
        "Combined Loan & Advance Bal.",
        "US$MM",
        ["=9000+800", "=10000+850", "=11000+900", "=12000+950", "=12500+1000"],
        ltm_value="=12500+1000",
    )
    ws = load_workbook(_build(tmp_path, metrics=metrics)).active
    assert ws["B9"].data_type == "f"
    assert ws["B9"].value == "=9000+800"
    # Non-flow combined LTM is likewise a formula.
    assert ws["G9"].data_type == "f"
    assert ws["G9"].value == "=12500+1000"


def test_bare_text_metric_value_raises(tmp_path: Path):
    bad = _metrics()
    bad[0] = MetricSeries("Revenue", "US$MM", [3100.0, 3450.0, 3820.0, 4180.0, "n/a"],
                          result_label="LTM Revenue")
    with pytest.raises(ValueError):
        _build(tmp_path, metrics=bad)
