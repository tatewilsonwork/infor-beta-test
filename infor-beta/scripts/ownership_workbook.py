"""Populate the INFOR ownership workbook from SEDI insider holdings.

Companion-workbook builder for the pitch deck's insider-ownership slide,
mirroring ``ltm_metrics.py``. Parsing a SEDI "Insider Information by Issuer"
report — which insiders are current, their relationship codes, common-share
tranches, dates, and roles — is the ``ownership`` skill's LLM/web work; this
module is the deterministic write. Given the issuer's **current** insiders and
its total basic shares outstanding, it fills the template's Select-Insiders
data block (rows 39-65, columns B/F/G/J) and the % denominator (``F35``),
leaving the formula-driven display block (``B4:G17``), the include flags (col
H), the adjusted-# formulas (col I), and the institutional / Bloomberg side
untouched.

The display block ranks the top 12 insiders by common shares via
``LARGE``/``XLOOKUP`` over the I (``=H*F``) helper column, so only an insider's
common-share count drives the slide — options, RSUs, PSUs and DSUs are excluded
by design (they are not common shares). The template's top-12 display assumes
at least 12 insiders hold a positive common-share balance; insiders holding 0
common shares are still written (the analyst can toggle col H to exclude them).
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

_SHEET = "Ownership"
_DATA_FIRST_ROW = 39
_DATA_LAST_ROW = 65  # rows 39-65 -> 27 insider slots
_MAX_INSIDERS = _DATA_LAST_ROW - _DATA_FIRST_ROW + 1
_TOTAL_SHARES_CELL = "F35"
_DATE_FORMAT = "yyyy-mm-dd"  # SEDI reports dates as ISO 'YYYY-MM-DD'
_BLUE = Font(color="0000FF")  # hardcoded-value convention (matches captable)

_COL_SEDI_NAME = "B"
_COL_BASIC = "F"
_COL_DATE = "G"
_COL_ADJ_NAME = "J"


@dataclass
class InsiderHolding:
    """One current insider's common-share holding from a SEDI report.

    - ``sedi_name``: the raw SEDI name ("Last, First Middle") -> col B.
    - ``adjusted_name``: presentation name + role, e.g.
      ``"Mark Barrenechea (CEO & Director)"`` -> col J (what the slide shows).
    - ``common_shares``: the common-share count, or a list of per-registered-
      holder tranche counts (direct, RRSP, Holdco, ...) which is written as a
      sum formula -> col F. Common shares ONLY — never options/RSU/PSU/DSU.
    - ``most_recent_date``: the latest common-share transaction date -> col G.
    """

    sedi_name: str
    adjusted_name: str
    common_shares: int | list[int]
    most_recent_date: date | str | None = None


def _coerce_date(value: date | str | None) -> date | None:
    if value is None or isinstance(value, date):
        return value
    text = str(value).strip()
    return date.fromisoformat(text) if text else None


def _basic_shares_cell_value(common_shares: int | list[int]):
    """A plain int for a single tranche; an Excel sum formula for several.

    Per the analyst workflow, multiple common-share tranches are summed *in the
    cell* (``=193000+0+0``) rather than pre-totalled, so the workbook documents
    every registered holding and stays auditable.
    """
    if isinstance(common_shares, (list, tuple)):
        tranches = [int(t) for t in common_shares]
        if not tranches:
            return None
        if len(tranches) == 1:
            return tranches[0]
        return "=" + "+".join(str(t) for t in tranches)
    return int(common_shares)


def _normalize(insider: "InsiderHolding | dict") -> InsiderHolding:
    if isinstance(insider, InsiderHolding):
        return insider
    return InsiderHolding(
        sedi_name=insider["sedi_name"],
        adjusted_name=insider["adjusted_name"],
        common_shares=insider["common_shares"],
        most_recent_date=insider.get("most_recent_date"),
    )


def read_basic_shares_from_cap_table(captable_path: Path | str) -> int | None:
    """Return total basic shares outstanding (full units) from a cap table, or None.

    Sources ``F35`` for the ownership workbook from the companion cap table.
    Reads the cap table's Section VII basic-share **input** rows (``Cap with
    Links`` col F, rows 168-185, in millions) and sums them, converting to full
    units. The captable skill writes those as hardcoded literals (not formulas),
    so they read reliably with openpyxl regardless of recalc state — unlike the
    Section VII total ``F186`` (a ``SUM`` formula whose cached value openpyxl may
    not see). Sub-event rows (e.g. a buyback's negative row) net in correctly
    because they live in the same column. Returns None when the file is
    missing/unreadable or the sum is non-positive, so the caller can leave
    ``F35`` blank for the analyst.
    """
    try:
        wb = load_workbook(Path(captable_path), data_only=False)
    except Exception:
        return None
    ws = wb["Cap with Links"] if "Cap with Links" in wb.sheetnames else wb.active
    total_millions = 0.0
    found = False
    for row in range(168, 186):  # Section VII basic-share inputs (rows 168-185)
        value = ws[f"F{row}"].value
        if isinstance(value, (int, float)):
            total_millions += float(value)
            found = True
    if not found or total_millions <= 0:
        return None
    return round(total_millions * 1_000_000)


def build_ownership_workbook(
    *,
    template_path: Path | str,
    insiders: "list[InsiderHolding | dict]",
    total_shares_outstanding: int | None,
    output_path: Path | str,
) -> Path:
    """Fill the ownership template's insider block and return the saved path.

    ``insiders`` lists the issuer's **current** insiders only — those whose
    SEDI "Ceased to be Insider" is "Not Applicable". ``total_shares_outstanding``
    is the issuer's basic shares outstanding in **full units** (not millions),
    typically sourced from the companion cap table; it drives the % column via
    ``F35``. Pass ``None`` to leave ``F35`` for the analyst to fill.
    """
    insiders = [_normalize(i) for i in insiders]
    if len(insiders) > _MAX_INSIDERS:
        raise ValueError(
            f"ownership template holds {_MAX_INSIDERS} insider rows "
            f"({_DATA_FIRST_ROW}-{_DATA_LAST_ROW}); got {len(insiders)} current insiders"
        )

    template = Path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"ownership template not found: {template}")
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, out)

    wb = load_workbook(out)  # preserve formulas (no data_only)
    if _SHEET not in wb.sheetnames:
        raise KeyError(f"sheet {_SHEET!r} not found in ownership template (have {wb.sheetnames})")
    ws = wb[_SHEET]

    for offset, insider in enumerate(insiders):
        row = _DATA_FIRST_ROW + offset
        name_cell = ws[f"{_COL_SEDI_NAME}{row}"]
        name_cell.value = insider.sedi_name
        name_cell.font = _BLUE

        basic_cell = ws[f"{_COL_BASIC}{row}"]
        basic_cell.value = _basic_shares_cell_value(insider.common_shares)
        basic_cell.font = _BLUE

        date_cell = ws[f"{_COL_DATE}{row}"]
        coerced = _coerce_date(insider.most_recent_date)
        if coerced is not None:
            date_cell.value = coerced
            date_cell.number_format = _DATE_FORMAT
            date_cell.font = _BLUE

        adj_cell = ws[f"{_COL_ADJ_NAME}{row}"]
        adj_cell.value = insider.adjusted_name
        adj_cell.font = _BLUE

    if total_shares_outstanding is not None:
        total_cell = ws[_TOTAL_SHARES_CELL]
        total_cell.value = int(total_shares_outstanding)
        total_cell.font = _BLUE
        total_cell.comment = Comment(
            "Total basic shares outstanding (full units) - from the companion cap table.",
            "INFOR",
        )

    wb.save(out)
    return out
