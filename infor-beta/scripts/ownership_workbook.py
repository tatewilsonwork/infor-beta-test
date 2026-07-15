"""Populate the INFOR ownership workbook from SEDI insider holdings.

Companion-workbook builder for the pitch deck's insider-ownership slide,
mirroring ``ltm_metrics.py``. Parsing a SEDI "Insider Information by Issuer"
report — which insiders are current, their relationship codes, common-share
tranches, dates, and roles — is the ``ownership`` skill's LLM/web work; this
module is the deterministic write. Given the issuer's **current** insiders and
its total basic shares outstanding, it fills the template's Select-Insiders
data block (rows 39-65, columns B/F/G/J) and the % denominator (``F35``),
leaving the formula-driven display block (``B4:G17``), the include flags (col
H), and the adjusted-# formulas (col I) untouched.

When the analyst also attaches a **Bloomberg ownership export** (the BBG Excel
add-in "Ownership" template — a ``Summary View`` sheet with holder rows from
row 14), the builder fills the institutional side too: the export's holder
rows are copied into the template's ``Bloomberg Output`` tab (whose layout
mirrors the BBG Summary View, so the Ownership tab's pre-wired rows 68-185
``XLOOKUP`` link up unchanged), each Bloomberg holder that duplicates a SEDI
insider gets its Include flag (col H) set to 0 — **the SEDI figure always
wins; only the Bloomberg duplicate is excluded** — every populated Bloomberg
row gets an adjusted display name (col J, legal suffixes stripped), and the
unused link rows are neutralised so the Select-Institutions ``LARGE`` block
computes. Without an export the institutional side stays untouched (today's
insider-only behaviour).

The display block ranks the top 12 insiders by common shares via
``LARGE``/``XLOOKUP`` over the I (``=H*F``) helper column, so only an insider's
common-share count drives the slide — options, RSUs, PSUs and DSUs are excluded
by design (they are not common shares). The template's top-12 display assumes
at least 12 insiders hold a positive common-share balance; insiders holding 0
common shares are still written (the analyst can toggle col H to exclude them).
"""

from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

from template_layout import (
    CAP_TABLE_SECTION_VII_ANCHORS,
    CAP_TABLE_SHEET,
    OWNERSHIP_BBG_LINK_ANCHORS,
    OWNERSHIP_BBG_TEMPLATE_ANCHORS,
    OWNERSHIP_INSIDER_BLOCK_ANCHORS,
    OWNERSHIP_TEMPLATE,
    OWNERSHIP_TOTAL_SHARES_ANCHORS,
    verify_anchors,
)

_SHEET = "Ownership"
_DATA_FIRST_ROW = 39
_DATA_LAST_ROW = 65  # rows 39-65 -> 27 insider slots
_MAX_INSIDERS = _DATA_LAST_ROW - _DATA_FIRST_ROW + 1
_TOTAL_SHARES_CELL = "F35"
_DATE_FORMAT = "yyyy-mm-dd"  # SEDI reports dates as ISO 'YYYY-MM-DD'

_COL_SEDI_NAME = "B"
_COL_BASIC = "F"
_COL_DATE = "G"
_COL_ADJ_NAME = "J"
_COL_INCLUDE = "H"

# ── Bloomberg institutional side ─────────────────────────────────────────────
# The template's 'Bloomberg Output' tab mirrors the BBG Excel add-in's
# "Summary View" layout: info cells (E7 name, E9 ticker, I9 view, E11/I11
# sort), row-13 column headers, holder rows from row 14. The Ownership tab's
# rows 68-185 are pre-wired against C14:C131 (name), L (position), N (filing
# date), so writing the export's rows into the same coordinates links the
# Select-Institutions block up with no formula work.
_BBG_SHEET = "Bloomberg Output"
_BBG_HEADER_ROW = 13
_BBG_FIRST_ROW = 14
_BBG_LAST_ROW = 131  # C14:C131 -> 118 holder slots (Ownership rows 68-185)
_MAX_BBG_HOLDERS = _BBG_LAST_ROW - _BBG_FIRST_ROW + 1
# Columns copied per holder row (col B keeps the template's own 1-118 numbering).
_BBG_COPY_COLS = (
    "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q",
    "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC",
)
_BBG_INFO_CELLS = ("E7", "E9", "I9", "E11", "I11")  # name, ticker, view, sort by, order
# Row-13 headers used to recognise / validate the BBG Summary View layout.
_BBG_EXPECTED_HEADERS = {"C": "Holder Name", "L": "Position", "N": "Filing Date", "R": "Insider Status"}
_OWN_BBG_FIRST_ROW = 68
_OWN_BBG_LAST_ROW = 185


def _blue(cell) -> Font:
    """Blue font for a hardcoded input cell, preserving the template's typeface.

    The hardcoded-value convention is blue text (matches the cap table), but a
    bare ``Font(color="0000FF")`` carries no name/size, so it resets the cell to
    PowerPoint/Excel's Calibri 11 default and drops the template's Palatino. Read
    the cell's existing font and re-emit it blue so name/size/weight survive.
    """
    f = cell.font
    return Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic, color="0000FF")


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
    ws = wb[CAP_TABLE_SHEET] if CAP_TABLE_SHEET in wb.sheetnames else wb.active
    # A readable-but-shifted cap table must raise, not silently sum the wrong
    # window: verify the Section VII sentinels before reading F168:F185.
    verify_anchors(ws, CAP_TABLE_SECTION_VII_ANCHORS, template=Path(captable_path).name)
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


@dataclass
class BloombergHolder:
    """One holder row from a Bloomberg ownership export's Summary View.

    ``values`` / ``number_formats`` map column letters (C..AC) to the source
    cell's value / number format, so the row can be copied into the template's
    ``Bloomberg Output`` tab verbatim.
    """

    name: str
    position: float | int | None = None
    insider_status: str | None = None  # "Y" = insider, "N-P" etc. = institution
    institution_type: str | None = None
    filing_date: date | datetime | None = None
    values: dict[str, object] = field(default_factory=dict)
    number_formats: dict[str, str] = field(default_factory=dict)


@dataclass
class BloombergExport:
    """A parsed Bloomberg ownership export (BBG Excel add-in, Summary View)."""

    company_name: str | None
    ticker: str | None
    holders: list[BloombergHolder]
    info: dict[str, object] = field(default_factory=dict)  # info-cell coordinate -> value


def _find_summary_view(wb):
    """The sheet carrying the BBG Summary View layout ('Holder Name' in C13)."""
    candidates = list(wb.worksheets)
    if "Summary View" in wb.sheetnames:
        candidates.insert(0, wb["Summary View"])
    for ws in candidates:
        if ws.cell(row=_BBG_HEADER_ROW, column=3).value == "Holder Name":
            return ws
    return None


def read_bloomberg_export(path: Path | str) -> BloombergExport:
    """Parse a Bloomberg ownership export (.xlsm/.xlsx from the BBG Excel add-in).

    Locates the Summary View sheet (``'Holder Name'`` in ``C13``), validates the
    column layout the ownership template's link formulas assume, and returns the
    holder rows (values + number formats for columns C..AC). Raises with a clear
    message when the file does not look like a BBG ownership Summary View, so the
    skill can tell the analyst what to re-attach.
    """
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"Bloomberg ownership export not found: {src}")
    wb = load_workbook(src, data_only=True)  # cached values for any add-in formulas
    ws = _find_summary_view(wb)
    if ws is None:
        raise ValueError(
            f"{src.name} does not contain a Bloomberg ownership Summary View "
            "('Holder Name' expected in C13) — attach the BBG Excel add-in "
            "Ownership export (Summary View)"
        )
    mismatched = {
        col: ws[f"{col}{_BBG_HEADER_ROW}"].value
        for col, expected in _BBG_EXPECTED_HEADERS.items()
        if ws[f"{col}{_BBG_HEADER_ROW}"].value != expected
    }
    if mismatched:
        raise ValueError(
            f"{src.name} Summary View columns differ from the layout the ownership "
            f"template links against (row {_BBG_HEADER_ROW}: expected "
            f"{_BBG_EXPECTED_HEADERS}, found {mismatched})"
        )

    holders: list[BloombergHolder] = []
    row = _BBG_FIRST_ROW
    while True:
        name = ws[f"C{row}"].value
        if name is None or str(name).strip() == "":
            break
        values: dict[str, object] = {}
        formats: dict[str, str] = {}
        for col in _BBG_COPY_COLS:
            cell = ws[f"{col}{row}"]
            if cell.value is not None:
                values[col] = cell.value
                formats[col] = cell.number_format
        holders.append(
            BloombergHolder(
                name=str(name).strip(),
                position=values.get("L"),
                insider_status=values.get("R"),
                institution_type=values.get("S"),
                filing_date=values.get("N"),
                values=values,
                number_formats=formats,
            )
        )
        row += 1

    info = {coord: ws[coord].value for coord in _BBG_INFO_CELLS if ws[coord].value is not None}
    return BloombergExport(
        company_name=ws["E7"].value,
        ticker=ws["E9"].value,
        holders=holders,
        info=info,
    )


# Trailing legal-form tokens stripped from an institution's display name. A
# token preceded by '&' is part of the brand, not a suffix ("Kelso & Co LP" ->
# "Kelso & Co", never "Kelso"). Stripping repeats ("Vanguard Group Inc" ->
# "Vanguard") but never empties the name.
_LEGAL_SUFFIX_TOKENS = {
    "inc", "incorporated", "corp", "corporation", "co", "company",
    "ltd", "limited", "llc", "lp", "llp", "ulc", "plc",
    "sa", "nv", "ag", "se", "spa",
    "group", "partners", "holdings", "holding",
}


def strip_legal_suffixes(name: str) -> str:
    """Default adjusted name: 'T Rowe Price Group Inc' -> 'T Rowe Price'.

    The skill agent may override individual names where house style differs
    (e.g. restoring 'T. Rowe Price' punctuation); this is the deterministic
    baseline. Person names pass through unchanged (their tokens are not legal
    suffixes).
    """
    tokens = str(name).split()
    while len(tokens) > 1:
        tail = tokens[-1].rstrip(".,").lower()
        if tail in _LEGAL_SUFFIX_TOKENS and tokens[-2] != "&":
            tokens.pop()
        else:
            break
    return " ".join(tokens) if tokens else str(name)


_NON_WORD_RE = re.compile(r"[^\w\s]", re.UNICODE)


def _name_tokens(name: str) -> list[str]:
    return [t for t in _NON_WORD_RE.sub(" ", str(name)).lower().split() if t]


def _token_matches(a: str, b: str) -> bool:
    """Exact token match, or an initial matching a full token's first letter."""
    if a == b:
        return True
    if len(a) == 1 or len(b) == 1:
        return a[0] == b[0]
    return False


def _givens_compatible(a: list[str], b: list[str]) -> bool:
    """Given-name token lists agree (order-free, initials tolerated).

    Both sides must carry *something* (surname-only never matches), full names
    must share at least one exact token when both sides have full names
    ("Mary" vs "M Christine" is rejected; "Eric" vs "Eric D" accepted), and
    every token of the shorter list must find a distinct flexible match in the
    longer ("Smith M Christine" matches SEDI "Smith, M Christine").
    """
    if not a or not b:
        return False
    a_full = {t for t in a if len(t) > 1}
    b_full = {t for t in b if len(t) > 1}
    if a_full and b_full and not (a_full & b_full):
        return False
    short, rest = (list(a), list(b)) if len(a) <= len(b) else (list(b), list(a))
    for t in short:
        matched = next((u for u in rest if _token_matches(t, u)), None)
        if matched is None:
            return False
        rest.remove(matched)
    return True


def bloomberg_matches_sedi(bbg_name: str, sedi_name: str) -> bool:
    """True when a Bloomberg holder duplicates a SEDI insider.

    Two paths: (1) corporate / exact — the suffix-stripped, punctuation-free
    token lists are identical (covers holdco insiders and letter-perfect person
    matches); (2) person — SEDI prints ``Last, First Middle`` and Bloomberg
    prints ``Last First Middle``, so the SEDI surname must prefix the Bloomberg
    tokens and the given names must agree per ``_givens_compatible``.
    """
    if _name_tokens(strip_legal_suffixes(bbg_name)) == _name_tokens(
        strip_legal_suffixes(sedi_name.replace(",", " "))
    ):
        return True
    if "," not in sedi_name:
        return False
    last, _, given = sedi_name.partition(",")
    surname = _name_tokens(last)
    btoks = _name_tokens(bbg_name)
    k = len(surname)
    if k == 0 or len(btoks) <= k or btoks[:k] != surname:
        return False
    return _givens_compatible(_name_tokens(given), btoks[k:])


def match_bloomberg_to_sedi(
    bbg_names: "list[str]", sedi_names: "list[str]"
) -> dict[str, str]:
    """Map each Bloomberg holder name to the SEDI insider it duplicates.

    Deterministic and conservative — the skill reviews this report (and any
    Bloomberg rows flagged Insider Status "Y" that did NOT match) and can
    correct individual rows via ``build_ownership_workbook``'s
    ``bloomberg_include_overrides``.
    """
    matches: dict[str, str] = {}
    for bbg in bbg_names:
        for sedi in sedi_names:
            if bloomberg_matches_sedi(bbg, sedi):
                matches[bbg] = sedi
                break
    return matches


def _write_bloomberg_side(
    wb,
    export: BloombergExport,
    *,
    sedi_names: "list[str]",
    adjusted_overrides: dict[str, str],
    include_overrides: dict[str, int],
) -> None:
    """Fill 'Bloomberg Output' + the Ownership tab's H/J columns (rows 68-185)."""
    if _BBG_SHEET not in wb.sheetnames:
        raise KeyError(f"sheet {_BBG_SHEET!r} not found in ownership template (have {wb.sheetnames})")
    ws_bbg = wb[_BBG_SHEET]
    ws_own = wb[_SHEET]
    # The export side is validated in read_bloomberg_export; this validates the
    # template side the rows land on — the 'Bloomberg Output' header row the
    # holder rows are copied under, and the Ownership tab's pre-wired link rows
    # 68–185 (whose H/J columns are written and B/F/G neutralised below).
    verify_anchors(ws_bbg, OWNERSHIP_BBG_TEMPLATE_ANCHORS, template=OWNERSHIP_TEMPLATE)
    verify_anchors(ws_own, OWNERSHIP_BBG_LINK_ANCHORS, template=OWNERSHIP_TEMPLATE)

    holders = export.holders
    if len(holders) > _MAX_BBG_HOLDERS:
        print(
            f"ownership: Bloomberg export has {len(holders)} holders; the template "
            f"holds {_MAX_BBG_HOLDERS} (C{_BBG_FIRST_ROW}:C{_BBG_LAST_ROW}) — writing the "
            "first 118 (the Summary View is sorted by position, so the tail is smallest)",
            file=sys.stderr,
        )
        holders = holders[:_MAX_BBG_HOLDERS]

    for coord, value in export.info.items():
        ws_bbg[coord] = value

    matches = match_bloomberg_to_sedi([h.name for h in holders], sedi_names)

    for offset, holder in enumerate(holders):
        bbg_row = _BBG_FIRST_ROW + offset
        for col, value in holder.values.items():
            cell = ws_bbg[f"{col}{bbg_row}"]
            cell.value = value
            fmt = holder.number_formats.get(col)
            if fmt and fmt != "General":
                cell.number_format = fmt

        own_row = _OWN_BBG_FIRST_ROW + offset
        include = include_overrides.get(holder.name, 0 if holder.name in matches else 1)
        if include != 1:
            include_cell = ws_own[f"{_COL_INCLUDE}{own_row}"]
            include_cell.value = 0
            include_cell.font = _blue(include_cell)
            reason = (
                f"Excluded - duplicate of SEDI insider '{matches[holder.name]}'; "
                "the SEDI figure (Select Insiders block) is kept."
                if holder.name in matches
                else "Excluded by the ownership skill (analyst/agent override)."
            )
            include_cell.comment = Comment(reason, "INFOR")

        adj_cell = ws_own[f"{_COL_ADJ_NAME}{own_row}"]
        adj_cell.value = adjusted_overrides.get(holder.name) or strip_legal_suffixes(holder.name)
        adj_cell.font = _blue(adj_cell)

    # Neutralise the unused link rows so the Select-Institutions LARGE range
    # stays numeric: an un-fed 'Bloomberg Output' XLOOKUP evaluates to #N/A,
    # which would poison LARGE($I$68:$I$185, k). Clearing B/F/G and zeroing H
    # leaves I (=H*F) at 0.
    for own_row in range(_OWN_BBG_FIRST_ROW + len(holders), _OWN_BBG_LAST_ROW + 1):
        for col in (_COL_SEDI_NAME, _COL_BASIC, _COL_DATE):
            ws_own[f"{col}{own_row}"] = None
        ws_own[f"{_COL_INCLUDE}{own_row}"] = 0


def build_ownership_workbook(
    *,
    template_path: Path | str,
    insiders: "list[InsiderHolding | dict]",
    total_shares_outstanding: int | None,
    output_path: Path | str,
    bloomberg_export_path: Path | str | None = None,
    bloomberg_adjusted_names: dict[str, str] | None = None,
    bloomberg_include_overrides: dict[str, int] | None = None,
) -> Path:
    """Fill the ownership template's insider block and return the saved path.

    ``insiders`` lists the issuer's **current** insiders only — those whose
    SEDI "Ceased to be Insider" is "Not Applicable". ``total_shares_outstanding``
    is the issuer's basic shares outstanding in **full units** (not millions),
    typically sourced from the companion cap table; it drives the % column via
    ``F35``. Pass ``None`` to leave ``F35`` for the analyst to fill.

    ``bloomberg_export_path`` (optional) is the analyst-attached Bloomberg
    ownership export; when supplied the builder also fills the ``Bloomberg
    Output`` tab and the Ownership tab's institutional H/J columns (see module
    docstring). ``bloomberg_adjusted_names`` overrides the default
    suffix-stripped display name per Bloomberg holder name;
    ``bloomberg_include_overrides`` (holder name -> 0/1) overrides the
    computed SEDI-duplicate exclusion for individual rows.
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
    # Verify the template layout before writing the hardcoded addresses blind:
    # the insider block rows 39–65 (header row 38 + the row-67 lower bound) and
    # the F35 % denominator.
    verify_anchors(
        ws,
        OWNERSHIP_INSIDER_BLOCK_ANCHORS + OWNERSHIP_TOTAL_SHARES_ANCHORS,
        template=OWNERSHIP_TEMPLATE,
    )

    for offset, insider in enumerate(insiders):
        row = _DATA_FIRST_ROW + offset
        name_cell = ws[f"{_COL_SEDI_NAME}{row}"]
        name_cell.value = insider.sedi_name
        name_cell.font = _blue(name_cell)

        basic_cell = ws[f"{_COL_BASIC}{row}"]
        basic_cell.value = _basic_shares_cell_value(insider.common_shares)
        basic_cell.font = _blue(basic_cell)

        date_cell = ws[f"{_COL_DATE}{row}"]
        coerced = _coerce_date(insider.most_recent_date)
        if coerced is not None:
            date_cell.value = coerced
            date_cell.number_format = _DATE_FORMAT
            date_cell.font = _blue(date_cell)

        adj_cell = ws[f"{_COL_ADJ_NAME}{row}"]
        adj_cell.value = insider.adjusted_name
        adj_cell.font = _blue(adj_cell)

    if total_shares_outstanding is not None:
        total_cell = ws[_TOTAL_SHARES_CELL]
        total_cell.value = int(total_shares_outstanding)
        # Leave F35's font untouched — the template ships it Palatino (bold), and
        # the aggregator later relinks it to the cap table's basic shares. Setting
        # a bare Font here would reset it to Calibri 11.
        total_cell.comment = Comment(
            "Total basic shares outstanding (full units) - from the companion cap table.",
            "INFOR",
        )

    if bloomberg_export_path is not None:
        _write_bloomberg_side(
            wb,
            read_bloomberg_export(bloomberg_export_path),
            sedi_names=[i.sedi_name for i in insiders],
            adjusted_overrides=bloomberg_adjusted_names or {},
            include_overrides=bloomberg_include_overrides or {},
        )

    wb.save(out)
    return out
