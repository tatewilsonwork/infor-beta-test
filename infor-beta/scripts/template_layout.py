"""The shipped templates' layout map: defined names and slide markers.

Nothing here is addressed blind. Every load-bearing cell in the four workbook
templates is reachable by a **defined name** the template itself carries, and
every library slide the assemblers touch is located by a **marker shape**
rather than an index. The point of both is the same: an analyst re-saving a
template or inserting a library slide must not require a code change.

**Named ranges.** Writers resolve an address through ``resolve_name_cell`` /
``resolve_name_range`` instead of hardcoding it, and verify the names they are
about to use with ``verify_names`` first. A defined name is metadata that Excel
moves with its cell when rows or columns are inserted, so ``infor_fx_rate``
follows the FX cell from ``F7`` to ``F9`` and the writer never notices. The
names are all ``infor_``-prefixed and **worksheet-scoped**: the cap table ships
33 Capital IQ defined names (``CIQWBGuid``, ``IQ_LTM``, …) and the comps
template 1,245 legacy artefacts (``_______AOL2``, ``__123Graph_A``, …), so a
bare ``fx_rate`` would be neither obviously ours nor safe from collision.

What verification catches is a template that no longer carries the names — an
analyst who re-created it, or re-saved it through a tool that strips them. That
halts the run with the remedy (re-run ``tools/add_template_named_ranges.py``)
rather than writing to a stale address. Sentinel labels — a second, parallel
table pairing each address with the caption text beside it — were carried
through Phase C as a cross-check that the names had been mapped to the right
cells, and were deleted once that had shipped (v0.5.40 → v0.5.42). Do not
reintroduce them: a sentinel pins an *address*, which is the thing the names
exist to stop mattering, and the two tables disagreeing was itself a failure
mode.

**Sheet names are for the SOURCE templates only.** Every ``*_SOURCE_SHEET``
constant here is the sheet name inside one of the four shipped source
templates — the artefacts ``tools/add_template_named_ranges.py`` stamps and
``tools/build_deal_workbook_template.py`` reads. The **pipeline** does not
produce those: Phase D put every stage on one deal workbook, whose tabs are
named by ``deal_workbook.TAB_*`` (``captable``, ``comps``, ``precedents``, …).
So anything holding a *built artefact* — both deck assemblers, ``financial_charts``,
the tab producers — addresses ``TAB_*``, never a ``*_SOURCE_SHEET``.

The two sets are deliberately not interchangeable, and the suffix is the whole
reminder: ``CAP_TABLE_SOURCE_SHEET`` is ``'Cap with Links'`` while
``TAB_CAPTABLE`` is ``'captable'``, and ``COMPS_SOURCE_SHEET`` / ``TAB_COMPS``
differ only in case — which openpyxl treats as a different sheet. Both deck
assemblers addressed the deal workbook by ``CAP_TABLE_SHEET`` (as it was then
spelled) from Phase D until v0.5.45, so every build that supplied a cap table
failed its deck stage. The one legitimate reader of a source sheet name at
runtime is ``ownership_workbook.read_basic_shares_from_cap_table``, which
accepts a pre-Phase-D standalone cap table as well as the deal workbook and
tries both spellings on purpose.

**Slide markers.** The pitch flow's self-discovery of Financial Summary slides
by their ``Rectangle 17`` marker shape is generalised here into
``find_slides_by_marker`` / ``find_slide_by_marker``, and every hardcoded slide
index now goes through it — the earnings assembler's kept-library set (which
needed three manual migrations), its filled-slide positions, the pitch layout
arithmetic, and ``financial_charts``' overview slide. Library markers identify
a *blank* library slide; the ``MARKER_BUILT_*`` markers identify a slide in an
*assembled* deck, where the titles have been filled and only the deferred
placeholders survive.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


class TemplateLayoutError(RuntimeError):
    """A shipped template's layout no longer matches what the code expects.

    Raised by the verify helpers when a workbook has lost the ``infor_`` defined
    names its writers resolve through, or when a library slide's marker shape is
    missing — i.e. the template was re-created without being re-stamped, or the
    slide library was re-ordered. The message names the template, the sheet or
    slide, what was expected, and the remedy.
    """


# ─── Defined-name resolution ─────────────────────────────────────────────────


def normalize_ref(ref: str) -> str:
    """A defined name's destination as a plain A1 reference.

    Strips the absolute-reference ``$`` markers Excel stores (``$B$15:$F$40`` ->
    ``B15:F40``) and upper-cases, so a resolved name compares directly against
    the shipped addresses recorded in ``TEMPLATE_NAMED_RANGES``.
    """
    return ref.replace("$", "").strip().upper()


def _lookup_defined_name(ws, name: str):
    """The ``DefinedName`` for ``name``, worksheet scope first, or ``None``.

    Our names are authored worksheet-scoped; the workbook-scope fallback is
    there so a template an analyst re-created with a workbook-scoped name of
    the same spelling still resolves rather than failing closed for a reason
    that is not a layout shift.
    """
    local = getattr(ws, "defined_names", None)
    if local is not None and name in local:
        return local[name]
    workbook = getattr(ws, "parent", None)
    if workbook is not None and name in workbook.defined_names:
        return workbook.defined_names[name]
    return None


def defined_name_ref(ws, name: str) -> str | None:
    """The A1 range ``name`` resolves to on ``ws``, or ``None`` if it does not.

    ``None`` covers all three "this sheet has no such usable name" cases: the
    name is absent, it resolves to several areas, or it points at a different
    sheet. Callers that need a hard failure use ``resolve_name_range``.
    """
    defined = _lookup_defined_name(ws, name)
    if defined is None:
        return None
    try:
        destinations = list(defined.destinations)
    except Exception:  # a constant / formula name has no destinations
        return None
    if len(destinations) != 1:
        return None
    sheet, ref = destinations[0]
    if sheet != ws.title:
        return None
    return normalize_ref(ref)


_RESTAMP_REMEDY = (
    "The shipped templates carry it; a re-saved template that dropped it must be "
    "re-stamped with tools/add_template_named_ranges.py before it can be used."
)


def resolve_name_range(ws, name: str, *, template: str | None = None) -> str:
    """The A1 range ``name`` resolves to on ``ws``; raise if it does not.

    This is what the writers call instead of hardcoding an address. The error
    names the template and the missing name, because the remedy is to re-run
    ``tools/add_template_named_ranges.py`` against the re-saved template.
    """
    ref = defined_name_ref(ws, name)
    if ref is None:
        raise TemplateLayoutError(
            f"{template or 'workbook'}, sheet {ws.title!r}: defined name {name!r} is "
            f"missing (or does not resolve to a single range on this sheet). "
            f"{_RESTAMP_REMEDY}"
        )
    return ref


def resolve_name_cell(ws, name: str, *, template: str | None = None) -> str:
    """The single cell ``name`` resolves to on ``ws`` (e.g. ``"F7"``); raise otherwise."""
    ref = resolve_name_range(ws, name, template=template)
    if ":" in ref:
        raise TemplateLayoutError(
            f"{template or 'workbook'}, sheet {ws.title!r}: defined name {name!r} "
            f"resolves to the range {ref!r}, but a single cell was expected."
        )
    return ref


def verify_names(ws, names, *, template: str) -> None:
    """Verify a sheet carries every defined name the caller is about to resolve.

    The pre-flight the writers run before touching a cell: resolving names one
    at a time would write the first few cells and then raise, leaving a
    half-filled tab, and would name only the first casualty. Every missing name
    is reported at once so a template that was re-created rather than re-stamped
    lists everything that has to be restored.
    """
    missing = [name for name in names if defined_name_ref(ws, name) is None]
    if missing:
        raise TemplateLayoutError(
            f"{template}, sheet {ws.title!r}: layout verification failed — defined "
            f"name(s) {', '.join(sorted(missing))} are missing (or do not resolve to "
            f"a single range on this sheet). {_RESTAMP_REMEDY} Refusing to read/write "
            f"these addresses."
        )


def verify_workbook_names(
    workbook_path: Path | str,
    *,
    sheet: str,
    names,
    template: str | None = None,
) -> None:
    """Open a workbook and verify one sheet carries the given defined names.

    Convenience wrapper for callers holding only a path — the deck assemblers'
    picture-range pre-flights. ``template`` defaults to the workbook's filename.
    Raises :class:`TemplateLayoutError` when the sheet is missing or a name is
    absent; file-level errors (missing/corrupt file) propagate as their native
    exceptions.
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
        verify_names(wb[sheet], names, template=name)
    finally:
        wb.close()


def resolve_workbook_range(
    workbook_path: Path | str, *, sheet: str, name: str, fallback: str
) -> str:
    """The A1 range ``name`` resolves to on a saved workbook, else ``fallback``.

    For the callers holding a path to a **built artefact** rather than a
    template: the deck assemblers' Excel picture ranges. Every workbook the
    pipeline produces descends from ``INFOR Deal Workbook Template.xlsx``, so
    this resolves; ``fallback`` is the shipped address, and is only reachable
    for a caller that skipped ``verify_workbook_names`` — which requires the
    name and so would already have raised.
    """
    from openpyxl import load_workbook

    wb = load_workbook(Path(workbook_path), read_only=False, data_only=False)
    try:
        if sheet not in wb.sheetnames:
            return fallback
        return defined_name_ref(wb[sheet], name) or fallback
    finally:
        wb.close()


# ─── Cap table (INFOR Cap Table Template.xlsx, sheet 'Cap with Links') ───────

CAP_TABLE_TEMPLATE = "INFOR Cap Table Template.xlsx"
CAP_TABLE_SOURCE_SHEET = "Cap with Links"

NAME_CAP_TICKER = "infor_cap_ticker"
NAME_CAP_OUTPUT_CCY = "infor_cap_output_ccy"
NAME_FX_RATE = "infor_fx_rate"
NAME_SHARE_PRICE = "infor_share_price"
NAME_BASIC_SHARES = "infor_basic_shares"
NAME_LTM_REVENUE_VALUATION = "infor_ltm_revenue_valuation"
NAME_LTM_EBITDA_VALUATION = "infor_ltm_ebitda_valuation"
NAME_CAP_PICTURE_RANGE = "infor_cap_picture_range"
NAME_CAP_SHARE_INPUTS = "infor_cap_share_inputs"

#: ``{name: address as shipped}``. The addresses are what the prep tool stamps;
#: nothing reads them at runtime — that is the whole point of the names.
CAP_TABLE_NAMED_RANGES: dict[str, str] = {
    NAME_CAP_TICKER: "F3",
    NAME_CAP_OUTPUT_CCY: "F5",
    NAME_FX_RATE: "F7",
    NAME_SHARE_PRICE: "F16",
    NAME_BASIC_SHARES: "F17",
    NAME_LTM_REVENUE_VALUATION: "D47",
    NAME_LTM_EBITDA_VALUATION: "D48",
    NAME_CAP_PICTURE_RANGE: "B15:F40",
    NAME_CAP_SHARE_INPUTS: "F168:F185",
}

#: Everything the captable skill writes: the header inputs (Steps 3/3b), the LTM
#: valuation cells (Step 6b), and the Section VII basic-share input block.
CAP_TABLE_WRITE_NAMES = (
    NAME_CAP_TICKER,
    NAME_CAP_OUTPUT_CCY,
    NAME_FX_RATE,
    NAME_SHARE_PRICE,
    NAME_BASIC_SHARES,
    NAME_LTM_REVENUE_VALUATION,
    NAME_LTM_EBITDA_VALUATION,
    NAME_CAP_SHARE_INPUTS,
)

#: Slide picture range — both deck assemblers paste this block.
CAP_TABLE_PICTURE_NAMES = (NAME_CAP_PICTURE_RANGE,)
CAP_TABLE_PICTURE_RANGE = "B15:F40"  # shipped address; `resolve_workbook_range` fallback

#: Section VII basic-share input rows, read by
#: `ownership_workbook.read_basic_shares_from_cap_table`.
CAP_TABLE_SECTION_VII_NAMES = (NAME_CAP_SHARE_INPUTS,)

#: Output-currency cell, read by the pitch assembler for the `[x]$MM` footnote.
CAP_TABLE_OUTPUT_CCY_NAMES = (NAME_CAP_OUTPUT_CCY,)
CAP_TABLE_OUTPUT_CCY_CELL = "F5"  # shipped address; fallback for a nameless workbook


def verify_cap_table_before_write(ws) -> None:
    """Verify every cap-table name the captable skill resolves before writing.

    Called by the captable skill (SKILL.md Step 3) on the freshly copied
    template before any cell write. Raises :class:`TemplateLayoutError` if the
    workbook has lost the names.
    """
    verify_names(ws, CAP_TABLE_WRITE_NAMES, template=CAP_TABLE_TEMPLATE)


# ─── Ownership (INFOR Ownership Template.xlsx) ────────────────────────────────

OWNERSHIP_TEMPLATE = "INFOR Ownership Template.xlsx"
OWNERSHIP_SOURCE_SHEET = "Ownership"
OWNERSHIP_BBG_SOURCE_SHEET = "Bloomberg Output"

NAME_OWN_INSIDER_BLOCK = "infor_own_insider_block"
NAME_OWN_TOTAL_SHARES = "infor_own_total_shares"
NAME_OWN_BBG_LINK_BLOCK = "infor_own_bbg_link_block"
NAME_OWN_BBG_HOLDER_BLOCK = "infor_own_bbg_holder_block"
NAME_OWN_INSIDERS_PICTURE = "infor_own_insiders_picture_range"
NAME_OWN_INSTITUTIONS_PICTURE = "infor_own_institutions_picture_range"

OWNERSHIP_NAMED_RANGES: dict[str, str] = {
    NAME_OWN_INSIDER_BLOCK: "B39:J65",
    NAME_OWN_TOTAL_SHARES: "F35",
    NAME_OWN_BBG_LINK_BLOCK: "H68:J185",
    NAME_OWN_INSIDERS_PICTURE: "B4:G17",
    NAME_OWN_INSTITUTIONS_PICTURE: "B19:G35",
}
OWNERSHIP_BBG_NAMED_RANGES: dict[str, str] = {NAME_OWN_BBG_HOLDER_BLOCK: "C14:AC131"}

#: The Select-Insiders data block plus the `%` denominator it divides by.
OWNERSHIP_INSIDER_WRITE_NAMES = (NAME_OWN_INSIDER_BLOCK, NAME_OWN_TOTAL_SHARES)

#: The Bloomberg side: the holder block the export's rows are copied into, and
#: the Ownership tab's pre-wired link rows that XLOOKUP against it. The two are
#: pre-wired one-to-one, so the builder also checks their spans match.
OWNERSHIP_BBG_HOLDER_NAMES = (NAME_OWN_BBG_HOLDER_BLOCK,)
OWNERSHIP_BBG_LINK_NAMES = (NAME_OWN_BBG_LINK_BLOCK,)

#: Slide picture ranges (pitch deck ownership slide).
OWNERSHIP_INSIDERS_PICTURE_NAMES = (NAME_OWN_INSIDERS_PICTURE,)
OWNERSHIP_INSIDERS_PICTURE_RANGE = "B4:G17"  # shipped address; fallback
OWNERSHIP_INSTITUTIONS_PICTURE_NAMES = (NAME_OWN_INSTITUTIONS_PICTURE,)
OWNERSHIP_INSTITUTIONS_PICTURE_RANGE = "B19:G35"  # shipped address; fallback


# ─── Comps (INFOR Comps Template.xlsx, sheet 'Comps') ─────────────────────────

COMPS_TEMPLATE = "INFOR Comps Template.xlsx"
COMPS_SOURCE_SHEET = "Comps"

NAME_COMPS_OUTPUT_CCY = "infor_comps_output_ccy"
# One label + one block name per vertical. The builder derives the ticker
# (column B) and description (column AA) rows from the block's own extent, so
# a template that grows a vertical from six rows to eight needs no code change.
NAME_COMPS_GROUP_LABELS = (
    "infor_comps_group1_label",
    "infor_comps_group2_label",
    "infor_comps_group3_label",
)
NAME_COMPS_GROUP_BLOCKS = (
    "infor_comps_group1_block",
    "infor_comps_group2_block",
    "infor_comps_group3_block",
)

#: `infor_comps_output_ccy` is stamped but resolved by nothing today — the
#: workbook aggregator relinked it to the cap table's output currency, and Phase
#: D deleted the aggregator. It stays in the template so the cell keeps a handle.
COMPS_NAMED_RANGES: dict[str, str] = {
    NAME_COMPS_OUTPUT_CCY: "F3",
    NAME_COMPS_GROUP_LABELS[0]: "D9",
    NAME_COMPS_GROUP_LABELS[1]: "D19",
    NAME_COMPS_GROUP_LABELS[2]: "D29",
    NAME_COMPS_GROUP_BLOCKS[0]: "B10:AA15",
    NAME_COMPS_GROUP_BLOCKS[1]: "B20:AA25",
    NAME_COMPS_GROUP_BLOCKS[2]: "B30:AA35",
}

#: Everything the comps builder writes: each vertical's label cell and data
#: block.
COMPS_WRITE_NAMES = (*NAME_COMPS_GROUP_LABELS, *NAME_COMPS_GROUP_BLOCKS)


# ─── Precedents (INFOR Precedents Template.xlsx, sheet 'Precedents') ──────────

PRECEDENTS_TEMPLATE = "INFOR Precedents Template.xlsx"
PRECEDENTS_SOURCE_SHEET = "Precedents"

# The plan calls this one `precedents_input_ccy`; the template labels the cell
# "Output:" and it carries the cap table's *output* currency, so the name
# follows the artefact rather than the plan's shorthand.
NAME_PREC_OUTPUT_CCY = "infor_prec_output_ccy"
NAME_PREC_GROUP_LABELS = (
    "infor_prec_group1_label",
    "infor_prec_group2_label",
)
NAME_PREC_GROUP_BLOCKS = (
    "infor_prec_group1_block",
    "infor_prec_group2_block",
)

PRECEDENTS_NAMED_RANGES: dict[str, str] = {
    NAME_PREC_OUTPUT_CCY: "C2",
    NAME_PREC_GROUP_LABELS[0]: "E7",
    NAME_PREC_GROUP_LABELS[1]: "E16",
    NAME_PREC_GROUP_BLOCKS[0]: "B8:AI13",
    NAME_PREC_GROUP_BLOCKS[1]: "B17:AI22",
}

#: Everything the precedents builder writes: the output-currency cell plus each
#: peer group's label and transaction block.
PRECEDENTS_WRITE_NAMES = (
    NAME_PREC_OUTPUT_CCY,
    *NAME_PREC_GROUP_LABELS,
    *NAME_PREC_GROUP_BLOCKS,
)


# ─── The defined-name registry ────────────────────────────────────────────────
# `tools/add_template_named_ranges.py` stamps exactly this, and
# `tools/build_deal_workbook_template.py` verifies the assembled deal workbook
# against it — so a name can never be declared in one place and stamped in
# another.

#: template filename -> sheet -> ``{name: A1 target}``. The order here is the
#: order the prep tool stamps them in.
TEMPLATE_NAMED_RANGES: dict[str, dict[str, dict[str, str]]] = {
    CAP_TABLE_TEMPLATE: {CAP_TABLE_SOURCE_SHEET: CAP_TABLE_NAMED_RANGES},
    OWNERSHIP_TEMPLATE: {
        OWNERSHIP_SOURCE_SHEET: OWNERSHIP_NAMED_RANGES,
        OWNERSHIP_BBG_SOURCE_SHEET: OWNERSHIP_BBG_NAMED_RANGES,
    },
    COMPS_TEMPLATE: {COMPS_SOURCE_SHEET: COMPS_NAMED_RANGES},
    PRECEDENTS_TEMPLATE: {PRECEDENTS_SOURCE_SHEET: PRECEDENTS_NAMED_RANGES},
}

# ─── Slide library (INFOR Slide Library.pptx) ─────────────────────────────────

SLIDE_LIBRARY_TEMPLATE = "INFOR Slide Library.pptx"
SLIDE_LIBRARY_SLIDE_COUNT = 17


@dataclass(frozen=True)
class SlideMarker:
    """A slide's identifying shape + expected text fragment.

    Generalises the pitch flow's FS-slide self-discovery (the ``Rectangle 17``
    metric-chart placeholder in ``financial_charts``): a slide is *located* and
    verified by a named shape whose text contains a fragment unique to that
    slide concept.

    Two families, because a marker has to survive whatever the caller is
    looking at. ``MARKER_*`` identify slides in the **blank library**, where
    every title is still its placeholder. ``MARKER_BUILT_*`` identify slides in
    an **assembled deck**, where the titles have been filled and only the
    deferred placeholders remain — so they key off those instead.
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
#
# Nothing addresses a library slide by these indices any more — the assemblers
# locate every slide with `find_slide_by_marker`. The map survives as the
# library's inventory: `verify_library_slide` uses it, `deck_contract` reads it
# to name a matched baseline slide, and `test_template_layout` asserts each
# marker still lands on exactly the index recorded here, which is what turns a
# re-ordered library into a caught regression rather than a silent one.
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

# Markers for an ASSEMBLED deck. The library markers above key off titles that
# assembly overwrites ("Introduction to [Client Name]" becomes "Introduction to
# Acme Corp"), so a post-fill lookup needs a shape the fill leaves alone. The
# deferred chart placeholders are exactly that, and they are also what the
# post-assembly stages are looking for: `financial_charts` finds the overview
# slide by the pie placeholder it is about to replace, the same way it already
# found the Financial Summary slides by their Metric #1 placeholder.
MARKER_BUILT_OVERVIEW = SlideMarker(
    "Rectangle 4", "[Pie Chart Placeholder]", "public-company overview"
)
MARKER_BUILT_FINANCIAL_SUMMARY = MARKER_FINANCIAL_SUMMARY


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


def slide_matches_marker(slide, marker: SlideMarker) -> bool:
    """True when ``slide`` carries ``marker``'s shape with its text fragment."""
    shape = next((s for s in slide.shapes if s.name == marker.shape_name), None)
    if shape is None or not getattr(shape, "has_text_frame", False):
        return False
    return marker.contains in shape.text_frame.text


def find_slides_by_marker(prs, marker: SlideMarker) -> list[int]:
    """Zero-based indices of every slide carrying ``marker``, in deck order.

    The general form of ``financial_charts._find_financial_summary_slides``.
    Returns ``[]`` rather than raising — callers that require a hit use
    ``find_slide_by_marker``, and the sections that legitimately repeat
    (Financial Summary, market entry) want the whole list.
    """
    return [i for i, slide in enumerate(prs.slides) if slide_matches_marker(slide, marker)]


def find_slide_by_marker(
    prs, marker: SlideMarker, *, template: str = SLIDE_LIBRARY_TEMPLATE
) -> int:
    """The index of the one slide carrying ``marker``; raise otherwise.

    This is what replaced the hardcoded slide indices. Insisting on **exactly
    one** match is the substance of the check: zero means the slide was removed
    or renamed, and more than one means the marker no longer identifies a
    single slide, which would make the caller's choice arbitrary. Both are
    layout errors, and both used to be silent.
    """
    hits = find_slides_by_marker(prs, marker)
    if len(hits) == 1:
        return hits[0]
    if not hits:
        raise TemplateLayoutError(
            f"{template}: no slide carries the {marker.description} marker "
            f"(shape {marker.shape_name!r} containing {marker.contains!r}) in a "
            f"{len(prs.slides)}-slide deck. The slide was removed, renamed, or its "
            f"marker shape was edited."
        )
    raise TemplateLayoutError(
        f"{template}: the {marker.description} marker (shape {marker.shape_name!r} "
        f"containing {marker.contains!r}) matches {len(hits)} slides {hits}, but it "
        f"must identify exactly one. The marker is no longer unique."
    )


def find_optional_slide_by_marker(
    prs, marker: SlideMarker, *, template: str = SLIDE_LIBRARY_TEMPLATE
) -> int | None:
    """``find_slide_by_marker``, but ``None`` when the slide is absent.

    For slides the deck spec can legitimately drop — today only Key Investment
    Highlights. Two or more matches still raise: an ambiguous marker is a
    layout error whether or not the slide is optional.
    """
    hits = find_slides_by_marker(prs, marker)
    if not hits:
        return None
    return find_slide_by_marker(prs, marker, template=template)


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
