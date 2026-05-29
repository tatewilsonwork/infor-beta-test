"""Tests for the cap-table Excel→PowerPoint renderer's LibreOffice recalc path.

The Windows COM path forces ``excel.CalculateFull()``; the non-Windows
(Cowork/Linux) path must force LibreOffice to recalculate the openpyxl-authored,
manual-calc workbook on load, or every formula cell prints as 0/blank.
"""

import shutil
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from excel_to_powerpoint import _soffice_convert, _write_lo_recalc_profile


def test_write_lo_recalc_profile_forces_always_recalc(tmp_path: Path):
    """The throwaway profile must set OOXML recalc-on-load to 0 (Always)."""
    uri = _write_lo_recalc_profile(tmp_path)

    xcu = tmp_path / "lo_profile" / "user" / "registrymodifications.xcu"
    assert xcu.exists(), "registrymodifications.xcu must live under <profile>/user/"
    text = xcu.read_text(encoding="utf-8")
    assert "OOXMLRecalcMode" in text
    # 0 = Always; 1 = Never and 2 = Prompt would both skip the recalc headless.
    assert "<value>0</value>" in text
    assert uri.startswith("file:") and uri.endswith("lo_profile")


def _build_cacheless_cap_table(path: Path) -> None:
    """A manual-calc cap table with formula cells and NO cached values, plus a
    CapIQ ``_xll.*`` cell LibreOffice cannot resolve."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Cap with Links"
    ws["F16"] = 30.0  # share price (hardcoded)
    ws["F17"] = 100.0  # basic shares (hardcoded)
    ws["F18"] = "=F16*F17"  # basic market cap -> 3000
    ws["F28"] = 500.0  # net debt (hardcoded)
    ws["F31"] = "=F18+F28"  # Enterprise Value -> 3500
    ws["D47"] = 1200.0  # LTM revenue, hardcoded by ltm-metrics
    ws["D48"] = 300.0  # LTM Adj. EBITDA, hardcoded
    ws["D34"] = '=IFERROR(D47,"n/a ")'  # -> 1200
    ws["D35"] = '=IFERROR(D48,"n/a ")'  # -> 300
    # A CapIQ add-in cell: unknown to LibreOffice -> #NAME?, must degrade not crash.
    ws["E47"] = '=_xll.SNL.Clients.Office.Excel.Functions.SPG("X","IQ_REV")'
    ws["E34"] = '=IFERROR(E47,"n/a ")'  # -> "n/a "
    wb.calculation.calcMode = "manual"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


@pytest.mark.skipif(sys.platform == "win32", reason="Windows uses the Excel COM path")
def test_libreoffice_recalc_populates_inworkbook_formulas(tmp_path: Path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        pytest.skip("LibreOffice (soffice/libreoffice) not installed")

    src = tmp_path / "cap_cacheless.xlsx"
    _build_cacheless_cap_table(src)
    # Precondition: openpyxl wrote no cached value, so without recalc this is blank.
    assert load_workbook(src, data_only=True)["Cap with Links"]["F31"].value is None

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    _soffice_convert(soffice, src, "xlsx:Calc MS Excel 2007 XML", out_dir)

    ws = load_workbook(out_dir / "cap_cacheless.xlsx", data_only=True)["Cap with Links"]
    # In-workbook arithmetic now computes through Enterprise Value.
    assert ws["F18"].value == pytest.approx(3000.0)
    assert ws["F31"].value == pytest.approx(3500.0)
    # The hardcoded LTM column (fed by ltm-metrics) resolves through IFERROR.
    assert ws["D34"].value == pytest.approx(1200.0)
    assert ws["D35"].value == pytest.approx(300.0)
    # CapIQ cell could not resolve -> degraded to n/a; the recalc did not crash.
    assert "n/a" in str(ws["E34"].value)
