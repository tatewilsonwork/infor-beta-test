"""Populate the deal workbook's `comps` tab from selected public comparables.

Fills the tab behind the pitch deck's comparable-companies ("public comps" /
trading comps) slide, mirroring ``ltm_metrics.py`` and ``ownership_workbook.py``.
Choosing the three verticals, the six public peers per vertical, their Capital IQ
tickers, and the one-line descriptions is the ``comps`` skill's LLM/web work;
this module is the deterministic write.

The tab holds three vertical (peer-group) blocks. Each
block has a label cell, six ticker rows and six description rows; every other
column is a Capital IQ array formula (``_xll.SNL…SPG($B10, …)``) that resolves
off the ticker in column B once the analyst opens the workbook with the Capital
IQ add-in active. This builder writes ONLY the three input fields — the vertical
label, the CapIQ ticker (column B) and the description (column AA) — and never
touches the CapIQ formulas or the group-average / global-statistic rows.

Where those fields ARE comes from the tab, not from this file: each
block's label cell and row span are read off its ``infor_comps_groupN_label`` /
``infor_comps_groupN_block`` defined names (``D9`` + rows 10-15, ``D19`` +
20-25, ``D29`` + 30-35 as shipped). A template whose verticals were moved, or
resized to a different number of peers, needs no code change.

Capital IQ cannot be refreshed in this environment, so the tab ships with its
formulas un-evaluated; the analyst refreshes them in Excel. The comps sheet
round-trips cleanly through openpyxl (CapIQ array formulas preserved, file stays
Excel-openable), which is what makes it safe for `write_tab` to load and re-save
the whole deal workbook on every stage's write; a regression test guards that the
formulas survive.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import range_boundaries

from deal_workbook import TAB_COMPS, TabSpec, write_tab
from template_layout import (
    COMPS_SHEET,
    COMPS_TEMPLATE,
    COMPS_WRITE_NAMES,
    NAME_COMPS_GROUP_BLOCKS,
    NAME_COMPS_GROUP_LABELS,
    resolve_name_cell,
    resolve_name_range,
    verify_names,
)

_SHEET = COMPS_SHEET
_MAX_VERTICALS = 3
_MAX_DESCRIPTION_CHARS = 50  # column AA width ~50; longer overflows visually in the cell

_COL_TICKER = "B"
_COL_DESCRIPTION = "AA"


def _vertical_slots(ws) -> list[tuple[str, range]]:
    """``(label cell, data rows)`` per vertical, resolved from the defined names.

    Both the label cell and the block extent come from the template rather than
    from constants: the shipped blocks are ``D9`` + rows 10-15, ``D19`` + 20-25,
    ``D29`` + 30-35, but a template whose verticals were moved or resized needs
    no code change, because the row span is read off ``infor_comps_groupN_block``.
    """
    slots: list[tuple[str, range]] = []
    for label_name, block_name in zip(NAME_COMPS_GROUP_LABELS, NAME_COMPS_GROUP_BLOCKS):
        label = resolve_name_cell(ws, label_name, template=COMPS_TEMPLATE)
        block = resolve_name_range(ws, block_name, template=COMPS_TEMPLATE)
        _, first_row, _, last_row = range_boundaries(block)
        slots.append((label, range(first_row, last_row + 1)))
    return slots


@dataclass
class CompCompany:
    """One public comparable: its Capital IQ ticker and a one-line description.

    - ``ticker``: CapIQ ``Exchange:Ticker`` identifier (e.g. ``"NasdaqGS:MSFT"``,
      ``"TSX:RY"``, ``"NYSE:JPM"``) -> column B. Every downstream metric column's
      array formula keys off this, so the exchange prefix is required.
    - ``description``: a <=50-char note on what the company does / sells ->
      column AA. No geography (the exchange prefix already signals it).
    """

    ticker: str
    description: str = ""


@dataclass
class Vertical:
    """One peer group ("vertical"): a label plus up to six comparables.

    - ``name``: the vertical / peer-group label -> ``D9`` / ``D19`` / ``D29``.
    - ``companies``: up to six ``CompCompany`` entries -> the block's rows.
    """

    name: str
    companies: list[CompCompany] = field(default_factory=list)


def _normalize_company(company: "CompCompany | dict") -> CompCompany:
    if isinstance(company, CompCompany):
        return company
    return CompCompany(ticker=company["ticker"], description=company.get("description", ""))


def _normalize_vertical(vertical: "Vertical | dict") -> Vertical:
    if isinstance(vertical, Vertical):
        companies = vertical.companies
        name = vertical.name
    else:
        companies = vertical.get("companies", [])
        name = vertical["name"]
    return Vertical(name=name, companies=[_normalize_company(c) for c in companies])


def build_comps_workbook(
    *,
    verticals: "list[Vertical | dict]",
    deal_workbook: Path | str,
) -> Path:
    """Fill the deal workbook's `comps` tab vertical blocks; return the workbook path.

    Since Phase D there is no standalone comps file and no template to copy: the
    `comps` tab is already in the deal workbook, carried in from
    `INFOR Deal Workbook Template.xlsx` at deal-init with its CapIQ array
    formulas and its `infor_comps_*` defined names intact.

    ``verticals`` lists up to three peer groups, each with up to six public
    comparables. Writes only the vertical labels, the CapIQ tickers (column B)
    and the descriptions (column AA); the CapIQ array formulas and the statistic
    rows are left untouched so they resolve when the analyst refreshes Capital
    IQ. Unfilled verticals keep the template's ``[Group #N]`` placeholder. Raises
    ValueError on too many verticals / companies, an empty ticker, or an
    over-length description.
    """
    verticals = [_normalize_vertical(v) for v in verticals]
    if not verticals:
        raise ValueError("no verticals supplied — expected 1–3 peer groups")
    if len(verticals) > _MAX_VERTICALS:
        raise ValueError(
            f"comps template holds {_MAX_VERTICALS} vertical blocks; got {len(verticals)}"
        )
    for vertical in verticals:
        for company in vertical.companies:
            if not str(company.ticker).strip():
                raise ValueError(f"empty ticker in vertical {vertical.name!r}")
            if len(company.description) > _MAX_DESCRIPTION_CHARS:
                raise ValueError(
                    f"description for {company.ticker!r} is {len(company.description)} "
                    f"chars; max {_MAX_DESCRIPTION_CHARS} (column AA overflows beyond that)"
                )

    def _write(_wb, ws) -> None:
        # Verify every vertical's label + block name resolves before writing
        # anything, so a tab that lost them halts here rather than half-filled.
        verify_names(ws, COMPS_WRITE_NAMES, template=COMPS_TEMPLATE)

        for vertical, (label_cell, rows) in zip(verticals, _vertical_slots(ws)):
            if len(vertical.companies) > len(rows):
                raise ValueError(
                    f"vertical {vertical.name!r} has {len(vertical.companies)} companies; "
                    f"the template holds {len(rows)} rows per vertical"
                )
            ws[label_cell] = vertical.name
            for row, company in zip(rows, vertical.companies):
                ws[f"{_COL_TICKER}{row}"] = str(company.ticker).strip()
                description = company.description.strip()
                if description:
                    ws[f"{_COL_DESCRIPTION}{row}"] = description

    write_tab(
        deal_workbook,
        TAB_COMPS,
        TabSpec(write=_write, verify_names=tuple(NAME_COMPS_GROUP_BLOCKS)),
    )
    return Path(deal_workbook)
