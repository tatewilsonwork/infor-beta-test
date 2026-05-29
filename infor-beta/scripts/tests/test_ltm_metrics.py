"""Unit tests for the LTM metrics workbook helper."""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from ltm_metrics import BridgeComponent, RevenueSegment, build_ltm_metrics_workbook


def _build(tmp_path: Path, **overrides) -> Path:
    kwargs = dict(
        company_name="SampleCo",
        period_label="LTM ended March 31, 2026",
        currency="US$MM",
        segmentation_basis="Service line",
        segments=[
            RevenueSegment("Cloud Services & Subscriptions", 1932.0),
            RevenueSegment("Customer Support", 1480.0),
            RevenueSegment("License", 360.0),
            RevenueSegment("Professional Service & Other", 290.0),
        ],
        revenue_bridge=[
            BridgeComponent("FY2025 Revenue", 5400.0),
            BridgeComponent("Q3 2026 YTD Revenue", 3050.0),
            BridgeComponent("Q3 2025 YTD Revenue", 2388.0, subtract=True),
        ],
        ebitda_bridge=[
            BridgeComponent("FY2025 Adj. EBITDA", 1820.0),
            BridgeComponent("Q3 2026 YTD Adj. EBITDA", 1040.0),
            BridgeComponent("Q3 2025 YTD Adj. EBITDA", 815.0, subtract=True),
        ],
        output_dir=tmp_path,
    )
    kwargs.update(overrides)
    return build_ltm_metrics_workbook(**kwargs)


def test_build_ltm_metrics_workbook_writes_expected_file(tmp_path: Path):
    path = _build(tmp_path)
    assert path.exists()
    assert path.name == "SampleCo - LTM Metrics.xlsx"
    assert load_workbook(path).active.title == "LTM Metrics"


def test_revenue_overview_uses_formulas_for_totals_and_percentages(tmp_path: Path):
    path = _build(tmp_path)
    ws = load_workbook(path).active  # formulas, not computed values

    # Section row 6, header row 7, data rows 8-11, total row 12.
    assert ws["A6"].value == "LTM Revenue Overview"
    assert ws["A7"].value == "Segment"
    assert ws["B8"].value == 1932.0
    # % of total references the total cell so Excel recomputes.
    assert ws["C8"].value == "=B8/B12"
    assert ws["B12"].value == "=SUM(B8:B11)"
    assert ws["C12"].value == "=SUM(C8:C11)"
    assert ws["A12"].value == "Total"


def test_revenue_bridge_sums_fy_plus_ytd_minus_prior(tmp_path: Path):
    path = _build(tmp_path)
    ws = load_workbook(path).active

    # Overview total is row 12; row 13 is a blank spacer, bridge starts at 14.
    assert ws["A14"].value == "LTM Revenue Bridge"
    assert ws["A15"].value == "Component"
    assert ws["A16"].value == "(+) FY2025 Revenue"
    assert ws["B16"].value == 5400.0
    assert ws["A18"].value == "(−) Q3 2025 YTD Revenue"
    # LTM = FY + current YTD − prior YTD, referencing the component cells.
    assert ws["A19"].value == "= LTM Revenue"
    assert ws["B19"].value == "=B16+B17-B18"


def test_ebitda_bridge_only_and_label(tmp_path: Path):
    path = _build(tmp_path)
    ws = load_workbook(path).active

    # Revenue result is row 19; row 20 is a blank spacer, EBITDA starts at 21.
    assert ws["A21"].value == "LTM Adj. EBITDA Bridge"
    assert ws["A26"].value == "= LTM Adj. EBITDA"
    assert ws["B26"].value == "=B23+B24-B25"


def test_ebitda_label_falls_back_to_unadjusted(tmp_path: Path):
    path = _build(tmp_path, ebitda_label="LTM EBITDA")
    ws = load_workbook(path).active
    assert ws["A21"].value == "LTM EBITDA Bridge"
    assert ws["A26"].value == "= LTM EBITDA"


def test_segments_and_bridges_accept_plain_tuples(tmp_path: Path):
    path = _build(
        tmp_path,
        segmentation_basis="Geography",
        segments=[("Americas", 2500.0), ("EMEA", 1100.0), ("Asia-Pacific", 462.0)],
        revenue_bridge=[
            ("FY2025 Revenue", 3500.0),
            ("Q3 2026 YTD Revenue", 2700.0),
            ("Q3 2025 YTD Revenue", 2138.0, True),
        ],
    )
    ws = load_workbook(path).active
    assert ws["A3"].value == "Revenue segmentation: Geography"
    # 3 segments -> data rows 8-10, total row 11.
    assert ws["B11"].value == "=SUM(B8:B10)"
    # Spacer row 12, then bridge: section 13, header 14, data 15-17, result 18.
    assert ws["A15"].value == "(+) FY2025 Revenue"
    assert ws["B18"].value == "=B15+B16-B17"


def test_bridges_optional(tmp_path: Path):
    path = _build(tmp_path, revenue_bridge=None, ebitda_bridge=None)
    ws = load_workbook(path).active
    # Only the overview block is present; no bridge section rows.
    assert ws["A6"].value == "LTM Revenue Overview"
    assert ws["A14"].value is None


def test_empty_segments_raise(tmp_path: Path):
    with pytest.raises(ValueError):
        _build(tmp_path, segments=[])
