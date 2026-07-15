"""Runtime verification of the shipped templates' hardcoded layout.

The Excel and PowerPoint writers address the shipped templates by hardcoded
cell addresses and slide indices (the cap table's ``D47``/``D48`` LTM cells,
the ownership insider block at rows 39–65, the earnings assembler's kept
library slide indices, …). A re-saved template with an inserted row, or a
re-ordered slide library, would previously fail *silently* — the writer puts a
number in the wrong cell or clones the wrong slide, and the error surfaces as
a wrong figure on a client slide.

This module is the central layout map: every load-bearing address is paired
with a **sentinel** — the label text the shipped template carries in a stable
anchor cell next to it (discovered by opening the shipped templates, never
invented). Writers call the verify helpers below *before* reading or writing a
hardcoded address; a mismatch raises :class:`TemplateLayoutError` naming the
template, the protected address, what was expected and what was found — so a
layout shift halts the run instead of shipping a wrong number.

Style model: the Bloomberg-export ``C13 = "Holder Name"`` header check in
``ownership_workbook.read_bloomberg_export`` — cheap, targeted cell reads (no
full-sheet scans), with a message that tells the analyst what to fix.

The slide-library side mirrors the pitch flow's existing self-discovery of
Financial Summary slides by their ``Rectangle 17`` marker shape
(``financial_charts._find_financial_summary_slides``): each library slide the
assemblers clone, delete, or fill is verified by a marker shape + expected
text before it is touched. ``OVERVIEW_SLIDE_INDEX`` (the assembled pitch
deck's public-company overview slide) lives here so ``pitch_deck_assembler``
and ``financial_charts`` share one definition.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


class TemplateLayoutError(RuntimeError):
    """A shipped template's layout no longer matches the hardcoded addresses.

    Raised by the verify helpers when a sentinel anchor cell / marker shape
    does not carry its expected label — i.e. the template was re-saved with
    rows/columns inserted, or the slide library was re-ordered. The message
    names the template, the protected address, the anchor, what was expected,
    and what was actually found.
    """


# ─── Cell anchors ─────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class CellAnchor:
    """One protected address and the sentinel label that anchors it.

    - ``target``: the load-bearing cell/range/rows being protected (what a
      writer is about to read or write blind).
    - ``label_addr``: the nearby cell that carries the sentinel label in the
      shipped template.
    - ``expected``: the sentinel text (verbatim from the shipped template).
    - ``contains``: substring match instead of exact — for anchors whose cell
      holds a formula string (e.g. the cap table's ``B16`` share-price label
      is a ``CONCATENATE`` formula containing ``"Share Price"``).
    """

    target: str
    label_addr: str
    expected: str
    contains: bool = False


def _anchor_text(value) -> str | None:
    """Normalise a cell value for sentinel comparison.

    Plain strings (labels and formula strings) pass through; openpyxl
    ``ArrayFormula`` objects contribute their formula text; everything else
    (numbers, None) is returned as-is / None.
    """
    if value is None or isinstance(value, str):
        return value
    text = getattr(value, "text", None)  # openpyxl ArrayFormula
    if isinstance(text, str):
        return text
    return str(value)


def _anchor_matches(anchor: CellAnchor, found: str | None) -> bool:
    if found is None:
        return False
    if anchor.contains:
        return anchor.expected in found
    return found.strip() == anchor.expected


def verify_cell_anchor(ws, anchor: CellAnchor, *, template: str) -> None:
    """Verify one sentinel anchor on an open openpyxl worksheet; raise on mismatch."""
    verify_anchors(ws, (anchor,), template=template)


def verify_anchors(ws, anchors, *, template: str) -> None:
    """Verify sentinel anchors on an open openpyxl worksheet.

    Checks every anchor and raises one :class:`TemplateLayoutError` listing
    all mismatches (so a shifted template reports every displaced address at
    once, not just the first).
    """
    problems: list[str] = []
    for anchor in anchors:
        found = _anchor_text(ws[anchor.label_addr].value)
        if not _anchor_matches(anchor, found):
            mode = "containing" if anchor.contains else "="
            problems.append(
                f"{anchor.target} expects anchor {anchor.label_addr} {mode} "
                f"{anchor.expected!r} but found {found!r}"
            )
    if problems:
        raise TemplateLayoutError(
            f"{template}, sheet {ws.title!r}: layout verification failed — "
            + "; ".join(problems)
            + ". The template layout has shifted; refusing to read/write the "
            "hardcoded addresses blind."
        )


def verify_workbook_anchors(
    workbook_path: Path | str, *, sheet: str, anchors, template: str | None = None
) -> None:
    """Open a workbook and verify sentinel anchors on one sheet.

    Convenience wrapper for callers holding only a path (the assemblers'
    picture-range inserts, the aggregator's relink pre-flight). ``template``
    defaults to the workbook's filename. Raises :class:`TemplateLayoutError`
    when the sheet is missing or an anchor mismatches; file-level errors
    (missing/corrupt file) propagate as their native exceptions.
    """
    from openpyxl import load_workbook

    path = Path(workbook_path)
    name = template or path.name
    wb = load_workbook(path, data_only=False)
    try:
        if sheet not in wb.sheetnames:
            raise TemplateLayoutError(
                f"{name}: layout verification failed — expected sheet {sheet!r} "
                f"(have {wb.sheetnames})"
            )
        verify_anchors(wb[sheet], anchors, template=name)
    finally:
        wb.close()


# ─── Cap table (INFOR Cap Table Template.xlsx, sheet 'Cap with Links') ───────

CAP_TABLE_TEMPLATE = "INFOR Cap Table Template.xlsx"
CAP_TABLE_SHEET = "Cap with Links"

# Header input cells the captable skill writes (Steps 3/3b) and the aggregator
# reads/references at relink time.
CAP_TABLE_TICKER_ANCHOR = CellAnchor("F3", "B3", "Ticker:")
CAP_TABLE_OUTPUT_CCY_ANCHOR = CellAnchor("F5", "B5", "Output Currency:")
CAP_TABLE_FX_ANCHOR = CellAnchor("F7", "B7", "FX Rate:")
# B16 is a CONCATENATE formula ('Share Price (<date>)'), so substring-match it.
CAP_TABLE_SHARE_PRICE_ANCHOR = CellAnchor("F16", "B16", "Share Price", contains=True)
CAP_TABLE_BASIC_SHARES_ANCHOR = CellAnchor("F17", "B17", "Basic Shares Outstanding")

CAP_TABLE_HEADER_ANCHORS = (
    CAP_TABLE_TICKER_ANCHOR,
    CAP_TABLE_OUTPUT_CCY_ANCHOR,
    CAP_TABLE_FX_ANCHOR,
    CAP_TABLE_SHARE_PRICE_ANCHOR,
    CAP_TABLE_BASIC_SHARES_ANCHOR,
)

# LTM valuation inputs (captable Step 6b; relinked to the ltm-metrics tab by
# the aggregator). D33 is the Financial Metrics block's LTM column header.
CAP_TABLE_LTM_ANCHORS = (
    CellAnchor("D47", "B47", "Revenue"),
    CellAnchor("D48", "B48", "Adj. EBITDA"),
    CellAnchor("D47:D48", "D33", "LTM"),
)

# Slide picture range (both deck assemblers paste 'Cap with Links'!B15:F40).
# B15 is the range's own top-left label; B40 pins the bottom row.
CAP_TABLE_PICTURE_RANGE = "B15:F40"
CAP_TABLE_PICTURE_ANCHORS = (
    CellAnchor(CAP_TABLE_PICTURE_RANGE, "B15", "Company Ticker:"),
    CellAnchor(CAP_TABLE_PICTURE_RANGE, "B40", "EV / Adj. EBITDA"),
)

# Section VII basic-share input rows (read by
# ownership_workbook.read_basic_shares_from_cap_table as F168:F185).
CAP_TABLE_SECTION_VII_ROWS = (168, 185)
CAP_TABLE_SECTION_VII_ANCHORS = (
    CellAnchor("F168:F185", "B166", "VII. BASIC SHARES OUTSTANDING"),
    CellAnchor("F168:F185", "B167", "Description"),
    CellAnchor("F168:F185", "F167", "Amount"),
    CellAnchor("F168:F185", "B186", "Total Basic Shares Outstanding"),
)


def verify_cap_table_before_write(ws) -> None:
    """Verify every cap-table region the captable skill writes blind.

    Called by the captable skill (SKILL.md Step 3) on the freshly copied
    template before any cell write: the header inputs (F3/F5/F7/F16), the LTM
    valuation cells (D47/D48), and the Section VII block the section writes
    land around. Raises :class:`TemplateLayoutError` if the template layout
    has shifted.
    """
    verify_anchors(
        ws,
        CAP_TABLE_HEADER_ANCHORS + CAP_TABLE_LTM_ANCHORS + CAP_TABLE_SECTION_VII_ANCHORS,
        template=CAP_TABLE_TEMPLATE,
    )


# ─── Ownership (INFOR Ownership Template.xlsx) ────────────────────────────────

OWNERSHIP_TEMPLATE = "INFOR Ownership Template.xlsx"
OWNERSHIP_SHEET = "Ownership"
OWNERSHIP_BBG_SHEET = "Bloomberg Output"

# Select-Insiders data block (rows 39–65, cols B/F/G/H/J): anchored by its
# header row 38 and bounded below by the row-67 Bloomberg link-block header.
OWNERSHIP_INSIDER_BLOCK_ANCHORS = (
    CellAnchor("B39:J65", "B38", "SEDI Name"),
    CellAnchor("B39:J65", "F38", "Basic"),
    CellAnchor("B39:J65", "G38", "Date"),
    CellAnchor("B39:J65", "H38", "Include"),
    CellAnchor("B39:J65", "J38", "Adj. Name"),
    CellAnchor("B39:J65", "B67", "From Bloomberg"),
)

# % denominator (written by the ownership skill; relinked to the cap table's
# basic shares by the aggregator).
OWNERSHIP_TOTAL_SHARES_ANCHORS = (
    CellAnchor("F35", "B35", "Total Basic Shares Outstanding"),
)

# Bloomberg link rows 68–185 (cols H/J written; B/F/G neutralised on unused
# rows): anchored by the row-67 header and the first/last pre-wired link
# formulas, which pin the 118-row extent against 'Bloomberg Output'!C14:C131.
OWNERSHIP_BBG_LINK_ANCHORS = (
    CellAnchor("H68:J185", "B67", "From Bloomberg"),
    CellAnchor("H68:J185", "B68", "'Bloomberg Output'!C14", contains=True),
    CellAnchor("H68:J185", "B185", "'Bloomberg Output'!C131", contains=True),
)

# The template's own 'Bloomberg Output' tab mirrors the BBG add-in Summary
# View; the export-side header check in read_bloomberg_export validates the
# attachment, this validates the template copy the holder rows land on.
OWNERSHIP_BBG_TEMPLATE_ANCHORS = (
    CellAnchor("C14:AC131", "C13", "Holder Name"),
    CellAnchor("C14:AC131", "L13", "Position"),
    CellAnchor("C14:AC131", "N13", "Filing Date"),
    CellAnchor("C14:AC131", "R13", "Insider Status"),
)

# Slide picture ranges (pitch deck ownership slide).
OWNERSHIP_INSIDERS_PICTURE_RANGE = "B4:G17"
OWNERSHIP_INSIDERS_PICTURE_ANCHORS = (
    CellAnchor(OWNERSHIP_INSIDERS_PICTURE_RANGE, "B4", "Select Insiders"),
    CellAnchor(OWNERSHIP_INSIDERS_PICTURE_RANGE, "B17", "Subtotal"),
)
OWNERSHIP_INSTITUTIONS_PICTURE_RANGE = "B19:G35"
OWNERSHIP_INSTITUTIONS_PICTURE_ANCHORS = (
    CellAnchor(OWNERSHIP_INSTITUTIONS_PICTURE_RANGE, "B19", "Select Institutions"),
    CellAnchor(OWNERSHIP_INSTITUTIONS_PICTURE_RANGE, "B35", "Total Basic Shares Outstanding"),
)


# ─── Comps (INFOR Comps Template.xlsx, sheet 'Comps') ─────────────────────────

COMPS_TEMPLATE = "INFOR Comps Template.xlsx"
COMPS_SHEET = "Comps"

# The three vertical blocks (label D9/D19/D29; tickers B10:B15/B20:B25/B30:B35;
# descriptions AA10…): anchored by the row-7 column headers above the first
# block and the 'Group Average' row that closes each 6-row block. The ticker
# column B itself has no adjacent header label in the template — its position
# is pinned transitively by the D/AA headers and the block bounds.
COMPS_BLOCK_ANCHORS = (
    CellAnchor("B10:AA15", "D7", "Company"),
    CellAnchor("B10:AA15", "AA7", "Description"),
    CellAnchor("B10:AA15", "D17", "Group Average"),
    CellAnchor("B20:AA25", "D27", "Group Average"),
    CellAnchor("B30:AA35", "D37", "Group Average"),
)

# Output-currency cell the aggregator relinks to the cap table's F5.
COMPS_OUTPUT_CCY_ANCHORS = (CellAnchor("F3", "E3", "Currency:"),)


# ─── Precedents (INFOR Precedents Template.xlsx, sheet 'Precedents') ──────────

PRECEDENTS_TEMPLATE = "INFOR Precedents Template.xlsx"
PRECEDENTS_SHEET = "Precedents"

# Output-currency cell (written by the builder; relinked by the aggregator).
PRECEDENTS_OUTPUT_CCY_ANCHORS = (CellAnchor("C2", "B2", "Output:"),)

# The two transaction blocks (rows 8–13 / 17–22) and the input columns the
# builder writes, anchored by the row-4/5 header labels and the 'Group
# Average' row that closes each block. Column AI (HQ code) has its own row-5
# header; column H is left empty by design (no anchor needed).
PRECEDENTS_BLOCK_ANCHORS = (
    CellAnchor("B8:B22", "B5", "Currency"),
    CellAnchor("E8:E22", "E5", "Ann. Date"),
    CellAnchor("F8:F22", "F5", "Target"),
    CellAnchor("G8:G22", "G5", "Acquiror"),
    CellAnchor("I8:I22", "I4", "TEV"),
    CellAnchor("I8:I22", "I5", "Source FX"),
    CellAnchor("K8:L22", "K4", "Revenue - Source FX"),
    CellAnchor("M8:N22", "M4", "Net Income - Source FX"),
    CellAnchor("O8:P22", "O4", "Adj. EBITDA - Source FX"),
    CellAnchor("Q8:R22", "Q4", "Book Value - Source FX"),
    CellAnchor("S8:V22", "S4", "EV / Revenue"),
    CellAnchor("S8:V22", "U4", "EV / EBITDA"),
    CellAnchor("W8:X22", "W4", "P/ E"),
    CellAnchor("Y8:Y22", "Y4", "P / B"),
    CellAnchor("Z8:Z22", "Z4", "P / TBV"),
    CellAnchor("AB8:AG22", "AB4", "TEV"),
    CellAnchor("AB8:AG22", "AG4", "P / TBV"),
    CellAnchor("AI8:AI22", "AI5", "HQ"),
    CellAnchor("B8:AI13", "E14", "Group Average"),
    CellAnchor("B17:AI22", "E23", "Group Average"),
)


# ─── Slide library (INFOR Slide Library.pptx) ─────────────────────────────────

SLIDE_LIBRARY_TEMPLATE = "INFOR Slide Library.pptx"
SLIDE_LIBRARY_SLIDE_COUNT = 17

# Zero-based index of the public-company overview slide in the ASSEMBLED pitch
# deck (= raw library index 6; unchanged by the pitch flow's deletion of the
# earnings slide at raw index 7). Shared by pitch_deck_assembler (fill + cap
# table insert) and financial_charts (LTM revenue pie insert) — previously two
# independent definitions.
OVERVIEW_SLIDE_INDEX = 6


@dataclass(frozen=True)
class SlideMarker:
    """A library slide's identifying shape + expected text fragment.

    Extends the pitch flow's existing FS-slide self-discovery pattern (the
    ``Rectangle 17`` metric-chart placeholder in ``financial_charts``): a slide
    is verified by a named shape whose text contains a fragment unique to that
    slide concept in the shipped library.
    """

    shape_name: str
    contains: str
    description: str


MARKER_COVER = SlideMarker("Title 1", "[CLIENT NAME]", "cover")
MARKER_OVERVIEW = SlideMarker("Title 6", "Introduction to [Client Name]", "public-company overview")
MARKER_EARNINGS_SUMMARY = SlideMarker("Title 1", "Earnings Summary", "earnings summary")
# Same marker shape financial_charts uses to self-discover FS slides.
MARKER_FINANCIAL_SUMMARY = SlideMarker(
    "Rectangle 17", "[Placeholder for Metric #1 Chart]", "financial summary"
)
MARKER_OWNERSHIP = SlideMarker("Title 6", "Ownership", "insider ownership")
MARKER_RISKS = SlideMarker("Title 2", "Considerations and Mitigants", "acquirer considerations/mitigants")
MARKER_COMPS = SlideMarker("Title 6", "Comparable Companies Analysis", "comparable companies")
MARKER_PRECEDENTS = SlideMarker("Title 6", "Precedent Transactions Analysis", "precedent transactions")
MARKER_KIH = SlideMarker("Title 1", "Key Investment Highlights", "key investment highlights")
MARKER_MARKET_ENTRY = SlideMarker("Title 1", "Market Entry Targets", "market entry targets")
MARKER_DISCLAIMER = SlideMarker("Title 1", "Disclaimer", "disclaimer")
MARKER_CONTACT = SlideMarker("Title 6", "Contact", "contact")

# Raw (blank) library slide index -> marker, for every slide the assemblers
# clone, delete, keep, or fill. Discovered from the shipped 17-slide library.
LIBRARY_SLIDE_MARKERS: dict[int, SlideMarker] = {
    0: MARKER_COVER,
    6: MARKER_OVERVIEW,
    7: MARKER_EARNINGS_SUMMARY,
    8: MARKER_FINANCIAL_SUMMARY,
    9: MARKER_OWNERSHIP,
    10: MARKER_RISKS,
    11: MARKER_COMPS,
    12: MARKER_PRECEDENTS,
    13: MARKER_KIH,
    14: MARKER_MARKET_ENTRY,
    15: MARKER_DISCLAIMER,
    16: MARKER_CONTACT,
}


def verify_slide_marker(
    slide, marker: SlideMarker, *, template: str, slide_index: int | None = None
) -> None:
    """Verify a slide carries its identifying marker shape; raise on mismatch.

    ``slide_index`` (zero-based, when known) is included in the error message.
    The marker shape is looked up among the slide's top-level shapes (every
    library marker is top-level).
    """
    where = f"slide index {slide_index}" if slide_index is not None else "slide"
    shape = next((s for s in slide.shapes if s.name == marker.shape_name), None)
    if shape is None:
        names = sorted({s.name for s in slide.shapes})
        raise TemplateLayoutError(
            f"{template}: layout verification failed — {where} was expected to be "
            f"the {marker.description} slide (shape {marker.shape_name!r} containing "
            f"{marker.contains!r}) but has no shape named {marker.shape_name!r} "
            f"(shapes: {names}). The slide library has been re-ordered or edited; "
            "refusing to clone/delete/fill by index blind."
        )
    text = shape.text_frame.text if getattr(shape, "has_text_frame", False) else ""
    if marker.contains not in text:
        raise TemplateLayoutError(
            f"{template}: layout verification failed — {where} was expected to be "
            f"the {marker.description} slide, but shape {marker.shape_name!r} reads "
            f"{text!r} instead of containing {marker.contains!r}. The slide library "
            "has been re-ordered or edited; refusing to clone/delete/fill by index blind."
        )


def verify_library_slide(prs, index: int, *, template: str = SLIDE_LIBRARY_TEMPLATE) -> None:
    """Verify the slide at a raw library index is the concept the map expects.

    ``prs`` is an open python-pptx ``Presentation`` of the blank library.
    Raises :class:`TemplateLayoutError` on an out-of-range index, an index this
    map does not cover, or a marker mismatch.
    """
    marker = LIBRARY_SLIDE_MARKERS.get(index)
    if marker is None:
        raise TemplateLayoutError(
            f"{template}: no layout marker is defined for library slide index {index} "
            f"(known indices: {sorted(LIBRARY_SLIDE_MARKERS)})"
        )
    if index >= len(prs.slides):
        raise TemplateLayoutError(
            f"{template}: layout verification failed — expected the "
            f"{marker.description} slide at index {index}, but the library has only "
            f"{len(prs.slides)} slides (expected {SLIDE_LIBRARY_SLIDE_COUNT})."
        )
    verify_slide_marker(prs.slides[index], marker, template=template, slide_index=index)
