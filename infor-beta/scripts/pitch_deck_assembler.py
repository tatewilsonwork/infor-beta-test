"""Assembler for the INFOR slide-library pitch deck.

The blank library is 16 slides (incl. the insider-ownership slide, which
follows the Financial Summary slide, and the precedent-transactions slide,
which follows the comparable-companies slide). The deck's slide mix is
configurable, so slide indices are computed from the SlidePlan rather than
hardcoded:

- the market-entry section grows across multiple slides — two targets per
  slide — by cloning the library's market-entry slide (8 targets → 4
  market-entry slides → 19-slide deck);
- the Financial Summary section grows the same way — four metrics per slide —
  when the SlidePlan carries more than one ``financial-summary`` entry (the
  deck spec's "2 slides / 8 metrics" option);
- the Key Investment Highlights slide is deleted when the SlidePlan omits its
  entry (the deck spec's "omit" option).

Every slide *after* the Financial Summary section shifts with the section's
size; disclaimer/contact remain the last two slides.
"""

from __future__ import annotations

import math
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Inches

from excel_to_powerpoint import insert_excel_into_placeholder
from naming import safe_filename
from pptx_helpers import (
    clone_slide_after,
    delete_slide,
    fill_footnote_token,
    find_shape,
    find_table_shape,
    fit_overview_textbox,
    iter_all_shapes,
    palatino_text_width_in,
    set_cell_text,
    set_text,
    write_bullets_or_plain,
)
from schemas import PitchDeckContent, SlidePlan
from template_layout import (
    CAP_TABLE_OUTPUT_CCY_ANCHOR,
    CAP_TABLE_PICTURE_ANCHORS,
    CAP_TABLE_PICTURE_RANGE,
    CAP_TABLE_SHEET,
    MARKER_COVER,
    MARKER_COMPS,
    MARKER_FINANCIAL_SUMMARY,
    MARKER_KIH,
    MARKER_MARKET_ENTRY,
    MARKER_OVERVIEW,
    MARKER_OWNERSHIP,
    MARKER_PRECEDENTS,
    MARKER_RISKS,
    OVERVIEW_SLIDE_INDEX,
    OWNERSHIP_INSIDERS_PICTURE_ANCHORS,
    OWNERSHIP_INSIDERS_PICTURE_RANGE,
    OWNERSHIP_INSTITUTIONS_PICTURE_ANCHORS,
    OWNERSHIP_INSTITUTIONS_PICTURE_RANGE,
    OWNERSHIP_SHEET,
    verify_library_slide,
    verify_slide_marker,
    verify_workbook_anchors,
)

# Zero-based index of the earnings-summary entry inserted into the shared
# 17-slide library. The pitch deck does not use it, so it is dropped on open,
# restoring the 16-slide pitch ordering the layout math below assumes.
_EARNINGS_LIBRARY_SLIDE_INDEX = 7

# Clone/delete targets in the RAW 17-slide library (before the earnings slide
# is dropped). Extra market-entry / Financial Summary slides are cloned BEFORE
# the delete so python-pptx allocates fresh, non-colliding slide part names.
# Each raw index is verified against its template_layout marker before it is
# cloned or deleted, so a re-ordered library raises TemplateLayoutError.
_LIBRARY_FINANCIAL_SUMMARY_INDEX = 8
_LIBRARY_MARKET_ENTRY_INDEX = 14

# Fixed deck indices after the earnings slide is dropped. Only slides BEFORE
# the Financial Summary section have fixed indices — everything after it is
# computed per-deck by `_pitch_layout` (the FS section can hold 1-2 slides and
# the Key Investment Highlights slide can be omitted). The overview index is
# shared with financial_charts via template_layout.OVERVIEW_SLIDE_INDEX.
_FINANCIAL_SUMMARY_FIRST_INDEX = 7 # slide 8 — first Financial Summary slide

# Financial Summary slide title shape + the four metric-label tiles it carries
# (each cloned FS slide has the same shape names).
_FS_TITLE_SHAPE = "Title 6"
_FS_METRIC_TILES = ["Rectangle 13", "Rectangle 12", "Rectangle 15", "Rectangle 14"]
_FS_TILES_PER_SLIDE = len(_FS_METRIC_TILES)

# Slide 7 cap-table placeholder; the picture covers the capitalization summary
# plus the Financial/Valuation metric rows (same range as the earnings overview;
# the range + its sentinel anchors live in template_layout).
_CAP_TABLE_PLACEHOLDER = "Rectangle 3"
_CAP_TABLE_SHEET = CAP_TABLE_SHEET
_CAP_TABLE_RANGE = CAP_TABLE_PICTURE_RANGE

# Insider-ownership slide. The left "Insiders" placeholder ('Rectangle 1') is
# replaced by a picture of the ownership workbook's Select-Insiders block. The
# right "Institutions" placeholder ('Rectangle 3') is replaced by the
# Select-Institutions block (top-12 + subtotal + Other Shareholders + Total)
# when the ownership stage ingested a Bloomberg export (Bloomberg Output C14
# populated); otherwise it stays a Bloomberg placeholder. The slide follows the
# Financial Summary section, so its deck index comes from `_pitch_layout`.
_OWNERSHIP_PLACEHOLDER = "Rectangle 1"
_OWNERSHIP_SHEET = OWNERSHIP_SHEET
_OWNERSHIP_RANGE = OWNERSHIP_INSIDERS_PICTURE_RANGE
_INSTITUTIONS_PLACEHOLDER = "Rectangle 3"
_INSTITUTIONS_RANGE = OWNERSHIP_INSTITUTIONS_PICTURE_RANGE


def _bullet_tuple(bullet) -> tuple[str, int]:
    return (bullet.text, bullet.level)


def _all_text(prs: Presentation) -> str:
    parts: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                parts.append(shape.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        parts.append(cell.text)
    return "\n".join(parts)


def _shape_text(shape) -> str:
    return shape.text if getattr(shape, "has_text_frame", False) else ""


def _replace_first_line(shape, first_line: str, remaining_lines: list[str] | None = None) -> None:
    lines = [first_line]
    if remaining_lines:
        lines.extend(remaining_lines)
    set_text(shape, lines)


def _rounded_rectangles(slide):
    return [shape for shape in slide.shapes if shape.name == "Rounded Rectangle 19"]


def _set_table_height(table_frame, total_height: int, min_heights: list[int] | None = None) -> None:
    """Resize a table to an exact total height by scaling its row heights.

    Mirrors setting a table's height in PowerPoint (drag the bottom handle / set
    Height in the Size pane): the graphic-frame extent and every row height are
    scaled to `total_height` proportionally, so the rows still sum to exactly the
    target (rounding remainder folded into the last row). Call AFTER the cells are
    filled — row heights are only meaningful once content is in place.

    `min_heights` (EMU, one per row) are per-row *content* minimums: a stored row
    height is only a MINIMUM, so PowerPoint re-grows any row whose text needs more
    than we declared and the rendered table ends up taller than the clamp (the
    live-run "5.91-inch table" report — one wrapped label re-grew its 0.28" row).
    Rows whose proportional share would fall below their minimum are pinned at the
    minimum and the shortfall is taken from the rows with headroom, so the
    *rendered* total still lands on `total_height`. If even the minimums exceed
    the target the minimums win (the table then overflows by the true content
    excess — content must shrink, not the declared heights).
    """
    table = table_frame.table
    rows = list(table.rows)
    if not rows or sum(r.height for r in rows) <= 0:
        return
    target = int(total_height)
    mins = [int(m) for m in min_heights] if min_heights is not None else [0] * len(rows)

    # Waterfall: pin rows whose proportional share of the remaining budget would
    # dip below their content minimum; re-share the rest until stable.
    free = set(range(len(rows)))
    budget = target
    alloc: dict[int, float] = {}
    while free:
        base = sum(rows[i].height for i in free)
        if base <= 0 or budget <= 0:
            for i in free:
                alloc[i] = mins[i]
            break
        alloc = {i: rows[i].height * budget / base for i in free} | {
            i: float(mins[i]) for i in range(len(rows)) if i not in free
        }
        pinned = [i for i in free if alloc[i] < mins[i]]
        if not pinned:
            break
        for i in pinned:
            free.discard(i)
            budget -= mins[i]

    heights = [max(int(round(alloc[i])), mins[i]) for i in range(len(rows))]
    # Fold the rounding remainder into the last unpinned row so rows sum exactly.
    spill = target - sum(heights)
    for i in reversed(range(len(rows))):
        if heights[i] + spill >= mins[i]:
            heights[i] += spill
            break
    for row, h in zip(rows, heights):
        row.height = max(0, h)
    table_frame.height = sum(max(0, h) for h in heights)


def _write_flexible_bullets(shape, bullets) -> None:
    """Write bullets, falling back to plain paragraphs if a library shape has no bullet glyph template."""
    write_bullets_or_plain(shape, [_bullet_tuple(bullet) for bullet in bullets])


def _fill_investment_highlights(slide, content: PitchDeckContent) -> None:
    """Fill the four numbered highlight quadrants; leave placeholders if no content."""
    if not content.investment_highlights:
        return
    quadrants: list[tuple[int, object, object]] = []
    for group in slide.shapes:
        if group.shape_type != MSO_SHAPE_TYPE.GROUP or not group.name.startswith("Group"):
            continue
        number = header_shape = body_shape = None
        for sub in iter_all_shapes(group.shapes):
            if sub.name.startswith("Oval") and getattr(sub, "has_text_frame", False) and sub.text.strip().isdigit():
                number = int(sub.text.strip())
            elif sub.name.startswith("Arrow: Pentagon"):
                header_shape = sub
            elif sub.name.startswith("Rectangle") and getattr(sub, "has_text_frame", False) and "[x]" in sub.text:
                body_shape = sub
        if number is not None and header_shape is not None and body_shape is not None:
            quadrants.append((number, header_shape, body_shape))
    quadrants.sort(key=lambda q: q[0])
    for idx, (_, header_shape, body_shape) in enumerate(quadrants):
        if idx < len(content.investment_highlights):
            highlight = content.investment_highlights[idx]
            set_text(header_shape, [highlight.header])
            set_text(body_shape, list(highlight.bullets))
        else:
            set_text(header_shape, [""])
            set_text(body_shape, [""])
    if content.investment_highlights_tagline:
        for shape in slide.shapes:
            if shape.name == "Text Placeholder 2" and getattr(shape, "has_text_frame", False):
                set_text(shape, [content.investment_highlights_tagline])


# Market-entry cell sizing: the label column is white at 11 pt (stepping down
# per-label when a long label would wrap — see _me_label_size_pt); the target
# value columns are 9 pt. The value cells carry the long Overview / Strategic
# Rationale copy, and PowerPoint grows a table row to fit its text (a stored row
# height is only a MINIMUM). At 10 pt the real per-target copy wrapped tall
# enough that PowerPoint re-expanded the whole table to ~6.3" on open — past the
# 5.71" clamp `_set_table_height` writes. 9 pt keeps that copy inside the
# clamped rows so the rendered table stays at 5.71"; pair it with concise
# Overview / Strategic Rationale cells (see pitch-content) for headroom.
_ME_VALUE_SIZE = 9
_ME_LABEL_COLOR = "FFFFFF"  # scheme bg1 (white) in the library

# Target total height for a filled market-entry (acquisition-target) table.
# Real content can grow rows past the library's ~5.7" so the table runs off the
# slide; after the cells are filled we clamp the table back to this height
# (mirrors dragging the table's resize handle in PowerPoint). 5.71" keeps the
# table's bottom above the slide edge (table top is 1.2" on a 7.5" slide).
_ME_TABLE_HEIGHT = Inches(5.71)

# Table-cell layout constants for the market-entry row-minimum estimates: the
# library cells carry 0.1" side / 0.05" top+bottom insets, and a Palatino line
# is ~1.2× the font size tall (PowerPoint renders a single 11 pt label row at
# 0.283" = 11/72 × 1.2 + 0.1 — the floor below which no row can be declared,
# since a stored row height is only a minimum).
_ME_CELL_SIDE_INSETS_IN = 0.2
_ME_CELL_TB_INSETS_IN = 0.1
_ME_LINE_HEIGHT_EM = 1.2
_ME_MIN_ROW_IN = 11 * _ME_LINE_HEIGHT_EM / 72 + _ME_CELL_TB_INSETS_IN  # ≈ 0.283
# Word-wrapping wastes some of each line (breaks fall on word boundaries), so
# widen the estimated text width before dividing by the cell width.
_ME_WRAP_WASTE = 1.08
# A label that would wrap in the label column steps down until it fits on one
# line (a wrapped label is what re-grew the 0.28" rows to 5.91" total in the
# live run — 'Geographic Footprint' is wider than the 1.457" usable column at
# 11 pt). Pair with concise row labels from pitch-content (≤ ~18 chars).
_ME_LABEL_SIZE_STEPS = (11, 10, 9)


def _me_label_size_pt(label: str, usable_width_in: float) -> int:
    """Largest step size at which `label` fits the label column on one line."""
    for pt in _ME_LABEL_SIZE_STEPS:
        if palatino_text_width_in(label, pt) <= usable_width_in:
            return pt
    return _ME_LABEL_SIZE_STEPS[-1]


def _me_cell_min_height_in(
    text: str, usable_width_in: float, font_pt: float, wrap_waste: float = _ME_WRAP_WASTE
) -> float:
    """Estimated rendered height of one filled cell (its row's content minimum).

    `wrap_waste` pads prose for word-boundary wrapping; labels pass 1.0 — the
    per-label font step-down already guarantees they fit on one line, and the
    character table's kerning-less sums are the conservative side of that check.
    """
    lines = 0
    for segment in (text.splitlines() or [""]):
        width = palatino_text_width_in(segment, font_pt) * wrap_waste
        lines += max(1, math.ceil(width / usable_width_in))
    return lines * font_pt * _ME_LINE_HEIGHT_EM / 72 + _ME_CELL_TB_INSETS_IN

# Slide 10 Considerations/Mitigants table sizing. The library ships the header
# row at 12 pt and the body cells at 10 pt; the old code hardcoded 9 pt / 8 pt,
# which rendered noticeably smaller than the template.
_RISK_HEADER_SIZE = 12
_RISK_BODY_SIZE = 10
# The table must render at the library's shipped height (5.17", PowerPoint's
# Size pane shows 5.18") — a stored row height is only a render-time MINIMUM,
# so long mitigant copy re-grows its row and the whole table with it (a live
# run rendered 5.36"). When the estimated content heights don't fit at 10 pt,
# the body steps down until they do (the header row stays 12 pt), then the
# declared rows are clamped back to the library height with the estimates as
# per-row floors (mirrors the market-entry treatment).
_RISK_BODY_SIZE_STEPS = (_RISK_BODY_SIZE, 9, 8)


def _fill_risk_table(slide, content: PitchDeckContent) -> None:
    """Fill the Considerations/Mitigants table and clamp it to the library height.

    Estimates each row's rendered height from its widest cell (risk vs. joined
    mitigants), steps the body font down ``_RISK_BODY_SIZE_STEPS`` until the
    estimates fit the library's shipped table height, writes the cells at that
    size, then scales the declared row heights back to exactly the library
    height with the estimates as per-row content floors — so the rendered
    table always lands on the library's 5.18".
    """
    table_frame = find_table_shape(slide)
    table = table_frame.table
    library_height = table_frame.height  # the library ships this table at 5.18"

    header = ("Considerations", "Mitigants")
    max_rows = min(len(content.risk_mitigants), len(table.rows) - 1)
    body = [
        (row.risk, "\n".join(row.mitigants)) for row in content.risk_mitigants[:max_rows]
    ]
    body += [("", "")] * (len(table.rows) - 1 - len(body))
    usable = [Emu(col.width).inches - _ME_CELL_SIDE_INSETS_IN for col in table.columns]

    def row_minimums(body_pt: int) -> list[int]:
        sized_rows = [(header, _RISK_HEADER_SIZE)] + [(cells, body_pt) for cells in body]
        return [
            Inches(max(_me_cell_min_height_in(text, usable[c], pt) for c, text in enumerate(cells)))
            for cells, pt in sized_rows
        ]

    body_size = _RISK_BODY_SIZE_STEPS[0]
    minimums = row_minimums(body_size)
    for pt in _RISK_BODY_SIZE_STEPS[1:]:
        if sum(minimums) <= library_height:
            break
        body_size = pt
        minimums = row_minimums(pt)

    set_cell_text(table.cell(0, 0), header[0], size_pt=_RISK_HEADER_SIZE)
    set_cell_text(table.cell(0, 1), header[1], size_pt=_RISK_HEADER_SIZE)
    for idx, (risk, mitigants) in enumerate(body):
        set_cell_text(table.cell(idx + 1, 0), risk, size_pt=body_size)
        set_cell_text(table.cell(idx + 1, 1), mitigants, size_pt=body_size)
    _set_table_height(table_frame, library_height, minimums)


def _output_currency_letter(workbook_path) -> str:
    """Derive the footnote currency letter ('US' / 'C') from the cap table.

    Reads the cap table's output-currency cell (``F5`` on ``Cap with Links``) so
    the ``[x]$MM`` footnote token resolves to ``US$MM`` / ``C$MM`` instead of
    being hardcoded. Falls back to ``C`` (the template default) if the cell is
    missing or unreadable, so a footnote never ships the literal ``[x]``.
    """
    try:
        from openpyxl import load_workbook

        wb = load_workbook(workbook_path, data_only=True)
        ws = wb["Cap with Links"] if "Cap with Links" in wb.sheetnames else wb.active
        code = str(ws["F5"].value or "").strip().upper()
    except Exception:
        code = ""
    if code.startswith("US"):
        return "US"
    if code.startswith("C"):  # CAD / C$
        return "C"
    return "C"


def _ownership_has_bloomberg(workbook_path) -> bool:
    """True when the ownership workbook carries Bloomberg institutional data.

    The ownership skill only populates the ``Bloomberg Output`` tab (first
    holder in ``C14``) when the analyst attached a Bloomberg export, so an
    empty C14 means the Select-Institutions block is uncomputed and the
    slide's right side must stay a placeholder.
    """
    try:
        from openpyxl import load_workbook

        wb = load_workbook(workbook_path, read_only=True)
        try:
            if "Bloomberg Output" not in wb.sheetnames:
                return False
            value = wb["Bloomberg Output"]["C14"].value
        finally:
            wb.close()
        return value is not None and str(value).strip() != ""
    except Exception:
        return False


def _fill_market_entry_targets(
    slide,
    *,
    row_labels: list[str],
    targets: list,
    market: str | None,
    currency_letter: str | None,
    slide_number: int,
    total_slides: int,
) -> None:
    """Fill one market-entry slide with up to two targets.

    `targets` holds the 1-2 targets for THIS slide. The table is the fixed
    12-row structure (Overview / HQ / Year Founded → 7 consistent industry
    metrics → Scale KPIs / Strategic Rationale): the label column (col 0) is
    written white at 11 pt (stepping down per-label so no label wraps in the
    column) and the target value columns at `_ME_VALUE_SIZE` (9 pt —
    deliberately below the library's 10 pt so the 5.71" table clamp holds; see
    the constant's comment). Each populated column's logo box is relabelled
    '[<target name> Logo]' (generic '[Company Name Logo]' when the target has no
    name); the unused box is blanked on an odd final slide so a single-target
    slide shows no stray logo box.
    """
    title = "Potential " + (f"{market} " if market else "") + "Market Entry Targets"
    if total_slides > 1:
        title += f" ({slide_number} of {total_slides})"
    set_text(find_shape(slide, "Title 1"), [title])

    if currency_letter is not None:
        footnote = next(
            (s for s in slide.shapes
             if s.name == "Text Placeholder 3" and getattr(s, "has_text_frame", False)),
            None,
        )
        if footnote is not None:
            fill_footnote_token(footnote, currency_letter)

    if not targets:
        return

    table_frame = find_table_shape(slide)
    table = table_frame.table
    n_cols = len(table.columns)  # label column + target columns (3 in the library)
    usable = [
        max(0.5, Emu(col.width).inches - _ME_CELL_SIDE_INSETS_IN) for col in table.columns
    ]
    # Per-row rendered content minimums (see _set_table_height): every row floors
    # at a single 11 pt label line; filled cells raise it by their estimated wrap.
    row_mins_in = [_ME_MIN_ROW_IN] * len(table.rows)
    # Table row 0 is the blank logo/header row; data labels start at row 1. With
    # the fixed 12-row structure every data row is populated — no blank rows.
    for i, label in enumerate(row_labels):
        row = i + 1
        if row >= len(table.rows):
            break
        label_pt = _me_label_size_pt(label, usable[0])
        set_cell_text(table.cell(row, 0), label, size_pt=label_pt, color_hex=_ME_LABEL_COLOR)
        row_mins_in[row] = max(
            row_mins_in[row], _me_cell_min_height_in(label, usable[0], label_pt, wrap_waste=1.0)
        )
        for col in range(1, n_cols):
            target = targets[col - 1] if (col - 1) < len(targets) else None
            value = target.cells[i] if target is not None else ""
            set_cell_text(table.cell(row, col), value, size_pt=_ME_VALUE_SIZE)
            row_mins_in[row] = max(
                row_mins_in[row], _me_cell_min_height_in(value, usable[col], _ME_VALUE_SIZE)
            )
    for row in range(len(row_labels) + 1, len(table.rows)):
        for col in range(n_cols):
            set_cell_text(table.cell(row, col), "", size_pt=_ME_VALUE_SIZE)

    # Align logo placeholders left→right with the target columns (sort by .left).
    # The library ships each box as the default '[Placeholder for Logo]'; for a
    # populated column, relabel it '[<target name> Logo]' (generic '[Company Name
    # Logo]' when the target carries no name) so the box names whose logo belongs
    # there. Blank any box whose column has no target (odd final slide).
    logos = sorted(
        (s for s in slide.shapes
         if getattr(s, "has_text_frame", False) and "[Placeholder for Logo]" in s.text),
        key=lambda s: s.left,
    )
    for col_idx, logo in enumerate(logos):
        if col_idx < len(targets):
            name = getattr(targets[col_idx], "name", None)
            set_text(logo, [f"[{name} Logo]" if name else "[Company Name Logo]"])
        else:
            set_text(logo, [""])

    # Clamp the now-filled table to a fixed height so long content can't run it
    # off the slide. Done last, after every cell is populated; the per-row
    # content minimums keep the declared heights achievable so PowerPoint's
    # render-time row growth can't push the total past the clamp.
    _set_table_height(
        table_frame, _ME_TABLE_HEIGHT, min_heights=[Inches(m) for m in row_mins_in]
    )


class _PitchLayout:
    """Computed zero-based deck indices for a configurable pitch slide mix.

    The prefix (cover .. overview, indices 0-6) is fixed; everything after the
    Financial Summary section shifts with ``n_financial_summary`` and the
    optional Key Investment Highlights slide. Disclaimer/contact follow the
    market-entry section.
    """

    def __init__(self, *, n_financial_summary: int, include_investment_highlights: bool, n_market_entry: int):
        if n_financial_summary < 1:
            raise ValueError("the deck needs at least one Financial Summary slide")
        self.n_financial_summary = n_financial_summary
        self.include_investment_highlights = include_investment_highlights
        self.n_market_entry = n_market_entry
        nfs = n_financial_summary
        self.financial_summary = list(range(_FINANCIAL_SUMMARY_FIRST_INDEX, _FINANCIAL_SUMMARY_FIRST_INDEX + nfs))
        self.ownership = 7 + nfs
        self.risks = 8 + nfs
        self.comps = 9 + nfs
        self.precedents = 10 + nfs
        self.investment_highlights = 11 + nfs if include_investment_highlights else None
        self.market_entry_first = 11 + nfs + (1 if include_investment_highlights else 0)
        self.total = self.market_entry_first + n_market_entry + 2  # + disclaimer/contact


def _verify_layout_slides(prs, layout: "_PitchLayout", template_name: str) -> None:
    """Verify every slide the layout math targets is the concept it expects.

    Runs once the slide mix is final (after clones + deletes), before any fill
    or insertion — the computed indices are exactly where a re-ordered library
    would silently misplace content. Raises TemplateLayoutError on mismatch.
    """
    checks: list[tuple[int, object]] = [
        (0, MARKER_COVER),
        (OVERVIEW_SLIDE_INDEX, MARKER_OVERVIEW),
    ]
    checks += [(idx, MARKER_FINANCIAL_SUMMARY) for idx in layout.financial_summary]
    checks += [
        (layout.ownership, MARKER_OWNERSHIP),
        (layout.risks, MARKER_RISKS),
        (layout.comps, MARKER_COMPS),
        (layout.precedents, MARKER_PRECEDENTS),
    ]
    if layout.investment_highlights is not None:
        checks.append((layout.investment_highlights, MARKER_KIH))
    checks += [
        (layout.market_entry_first + j, MARKER_MARKET_ENTRY)
        for j in range(layout.n_market_entry)
    ]
    for idx, marker in checks:
        verify_slide_marker(prs.slides[idx], marker, template=template_name, slide_index=idx)


def assemble_pitch_deck(
    *,
    slide_plan_path: Path | str,
    content_path: Path | str,
    template_path: Path | str,
    output_dir: Path | str,
    captable_workbook_path: Path | str | None = None,
    ownership_workbook_path: Path | str | None = None,
    financial_metric_labels: list[str] | None = None,
) -> Path:
    """Fill the INFOR slide-library pitch deck.

    The blank library is 16 slides; the SlidePlan drives the slide mix:

    - the market-entry section expands across multiple slides (two targets per
      slide) based on ``content.market_entry_targets``;
    - the Financial Summary section expands to one slide per four metrics when
      the SlidePlan carries repeated ``financial-summary`` entries (the extra
      slides are cloned from the library's FS slide and retitled ``(k of n)``);
    - the Key Investment Highlights slide is deleted when the SlidePlan omits
      its ``key-investment-highlights`` entry.

    The cap table is pasted into slide 7 when ``captable_workbook_path`` is
    supplied, and the insider-ownership table into the ownership slide when
    ``ownership_workbook_path`` is supplied (both via the ``excel_to_powerpoint``
    insertion helper). When the ownership workbook also carries Bloomberg
    institutional data, the Select-Institutions block is pasted into the slide's
    right placeholder too; other chart/table insertions remain deferred
    placeholders.

    ``financial_metric_labels`` are the Financial Summary tile labels, selected
    by and handed off from the ``financial-summary`` stage (no longer on
    ``PitchDeckContent``). When supplied it must hold exactly four names per
    Financial Summary slide in the plan (4 for the default single slide, 8 for
    two); when None the FS tiles keep their template placeholder text.
    """
    slide_plan = SlidePlan.model_validate_json(Path(slide_plan_path).read_text(encoding="utf-8"))
    content = PitchDeckContent.model_validate_json(Path(content_path).read_text(encoding="utf-8"))
    if slide_plan.deliverable_type != "pitch":
        raise ValueError("pitch deck assembler only supports pitch SlidePlan objects")
    if len(slide_plan.slides) < 15:
        raise ValueError("pitch deck expects at least the base INFOR Slide Library entries (incl. precedent transactions)")

    template = Path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"pitch library template not found: {template}")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"Pitch Deck - {safe_filename(content.client_name, default='Client')}.pptx"

    # Layout pre-flight on the companion workbooks whose cells and picture
    # ranges are read blind below (F5 for the footnote currency letter, then
    # 'Cap with Links'!B15:F40 / 'Ownership'!B4:G17 for the slide pictures) —
    # a shifted template raises here instead of pasting the wrong rows.
    if captable_workbook_path is not None:
        verify_workbook_anchors(
            captable_workbook_path,
            sheet=_CAP_TABLE_SHEET,
            anchors=(CAP_TABLE_OUTPUT_CCY_ANCHOR, *CAP_TABLE_PICTURE_ANCHORS),
        )
    if ownership_workbook_path is not None:
        verify_workbook_anchors(
            ownership_workbook_path,
            sheet=_OWNERSHIP_SHEET,
            anchors=OWNERSHIP_INSIDERS_PICTURE_ANCHORS,
        )

    # Footnote currency letter for the slide-7 + market-entry '[x]$MM' tokens,
    # derived from the cap table's output currency (None when no workbook).
    currency_letter = (
        _output_currency_letter(captable_workbook_path)
        if captable_workbook_path is not None
        else None
    )

    # The SlidePlan drives the deck's slide mix: the Financial Summary slide
    # count and the Key Investment Highlights toggle come from the wireframe's
    # entries; the market-entry slide count comes from the true content-bundle
    # target count (the wireframe count is only a default).
    plan_entry_ids = [entry.library_entry_id for entry in slide_plan.slides]
    n_financial_summary = max(1, plan_entry_ids.count("financial-summary"))
    include_kih = "key-investment-highlights" in plan_entry_ids
    n_market_entry = (
        max(1, math.ceil(len(content.market_entry_targets) / 2))
        if content.market_entry_targets
        else 1
    )
    layout = _PitchLayout(
        n_financial_summary=n_financial_summary,
        include_investment_highlights=include_kih,
        n_market_entry=n_market_entry,
    )
    if financial_metric_labels and len(financial_metric_labels) != _FS_TILES_PER_SLIDE * n_financial_summary:
        raise ValueError(
            f"financial_metric_labels holds {len(financial_metric_labels)} names but the "
            f"SlidePlan carries {n_financial_summary} Financial Summary slide(s) — expected "
            f"{_FS_TILES_PER_SLIDE * n_financial_summary} (four tiles per slide)"
        )

    prs = Presentation(template)

    # Verify every raw-library slide that is about to be cloned or deleted by
    # index — a re-ordered or re-saved library raises TemplateLayoutError here
    # instead of cloning/deleting the wrong slide.
    verify_library_slide(prs, _LIBRARY_MARKET_ENTRY_INDEX, template=template.name)
    verify_library_slide(prs, _LIBRARY_FINANCIAL_SUMMARY_INDEX, template=template.name)
    verify_library_slide(prs, _EARNINGS_LIBRARY_SLIDE_INDEX, template=template.name)

    # Grow the market-entry section (two targets per slide) and the Financial
    # Summary section (four metrics per slide) by cloning their library slides.
    # Clone BEFORE dropping the earnings slide so python-pptx allocates fresh,
    # non-colliding slide part names; market-entry first so its raw index is
    # still valid when the FS clones shift everything after index 8.
    for _ in range(n_market_entry - 1):
        clone_slide_after(prs, _LIBRARY_MARKET_ENTRY_INDEX)
    for _ in range(n_financial_summary - 1):
        clone_slide_after(prs, _LIBRARY_FINANCIAL_SUMMARY_INDEX)

    # The shared library carries the earnings-update slide at index 7. Drop it
    # so the remaining slides keep the canonical pitch ordering.
    delete_slide(prs, _EARNINGS_LIBRARY_SLIDE_INDEX)

    # Drop the Key Investment Highlights slide when the plan omits it. Done
    # after the earnings delete, at its computed post-delete index — verified
    # first, so a shifted library can't delete the wrong slide.
    if not include_kih:
        kih_index = 11 + n_financial_summary
        verify_slide_marker(
            prs.slides[kih_index], MARKER_KIH, template=template.name, slide_index=kih_index
        )
        delete_slide(prs, kih_index)

    # With the slide mix final, verify every slide the layout math targets
    # before anything is filled or inserted.
    _verify_layout_slides(prs, layout, template.name)

    # Slide 1 — cover: client name/date only.
    slide1 = prs.slides[0]
    title = find_shape(slide1, "Title 1")
    _replace_first_line(title, content.client_name, ["Internal Discussion Materials", ""])
    for shape in slide1.shapes:
        if shape.name == "Subtitle 2" and "[Date]" in _shape_text(shape):
            set_text(shape, [content.presentation_date])

    # Slide 2 — flexible executive summary bullets.
    slide2 = prs.slides[1]
    _write_flexible_bullets(
        find_shape(slide2, "Content Placeholder 7"),
        content.executive_summary_bullets,
    )

    # Slides 3–5 are static credentials. Do not touch.

    # Slide 6 — section divider labels.
    slide6 = prs.slides[5]
    rects = _rounded_rectangles(slide6)
    for idx, rect in enumerate(rects):
        if idx < len(content.section_labels):
            set_text(rect, [content.section_labels[idx]])
        else:
            set_text(rect, [""])

    # Slide 7 — public company overview. The cap table is pasted into the
    # 'Rectangle 3' placeholder after save (when a workbook is supplied); the
    # revenue pie stays a deferred placeholder. The bullets box is sized to the
    # band above the 'LTM Revenue Breakdown' header and gets an explicit
    # autofit fontScale when over-long, so the copy cannot render into the pie
    # section (PowerPoint ignores a scale-less autofit on open).
    slide7 = prs.slides[OVERVIEW_SLIDE_INDEX]
    set_text(find_shape(slide7, "Title 6"), [f"Introduction to {content.client_name}"])
    overview_shape = find_shape(slide7, "TextBox 9")
    _write_flexible_bullets(overview_shape, content.company_overview_bullets)
    fit_overview_textbox(slide7, overview_shape)
    if currency_letter is not None:
        fill_footnote_token(find_shape(slide7, "Text Placeholder 1"), currency_letter)

    # Financial Summary slide(s) — metric labels only (from the financial-summary
    # stage); charts remain placeholders. Left as template placeholders when no
    # labels. With two FS slides, tiles fill four labels per slide in order and
    # each slide is retitled '(k of n)'.
    for k, fs_index in enumerate(layout.financial_summary):
        fs_slide = prs.slides[fs_index]
        if n_financial_summary > 1:
            set_text(
                find_shape(fs_slide, _FS_TITLE_SHAPE),
                [f"Financial Summary ({k + 1} of {n_financial_summary})"],
            )
        if financial_metric_labels:
            slide_labels = financial_metric_labels[
                _FS_TILES_PER_SLIDE * k : _FS_TILES_PER_SLIDE * (k + 1)
            ]
            for shape_name, label in zip(_FS_METRIC_TILES, slide_labels, strict=True):
                set_text(find_shape(fs_slide, shape_name), [label])

    # Acquirer considerations/mitigants — concise risks + tagline; the table is
    # clamped back to the library's 5.18" after fill (see _fill_risk_table).
    slide_risks = prs.slides[layout.risks]
    set_text(find_shape(slide_risks, "Text Placeholder 6"), [content.risks_tagline])
    _fill_risk_table(slide_risks, content)

    # Comps takeaway; chart placeholder remains unless insertion later replaces it.
    slide_comps = prs.slides[layout.comps]
    set_text(find_shape(slide_comps, "Text Placeholder 5"), [content.comps_takeaway])

    # Precedent-transactions takeaway; chart placeholder remains (no
    # Excel→PowerPoint while Capital IQ can't be refreshed), mirroring comps.
    slide_prec = prs.slides[layout.precedents]
    set_text(find_shape(slide_prec, "Text Placeholder 5"), [content.precedents_takeaway])

    # Key investment highlights — only when the plan carries the slide;
    # placeholders remain unless content supplies them.
    if layout.investment_highlights is not None:
        slide_kih = prs.slides[layout.investment_highlights]
        _fill_investment_highlights(slide_kih, content)
        if currency_letter is not None:
            fill_footnote_token(find_shape(slide_kih, "Text Placeholder 13"), currency_letter)

    # Market-entry targets, two per slide. The section was grown above; fill
    # each slide with its pair and title it '(N of M)'.
    for j in range(n_market_entry):
        pair = content.market_entry_targets[2 * j : 2 * j + 2]
        _fill_market_entry_targets(
            prs.slides[layout.market_entry_first + j],
            row_labels=content.market_entry_row_labels,
            targets=pair,
            market=content.market_entry_market,
            currency_letter=currency_letter,
            slide_number=j + 1,
            total_slides=n_market_entry,
        )

    # Disclaimer + contact are static library entries — left untouched.

    prs.save(output_path)

    # Paste the generated cap table into slide 7's placeholder (mirrors the
    # earnings overview insertion). Done after save so the picture write re-opens
    # and re-saves the finished deck.
    if captable_workbook_path is not None:
        insert_excel_into_placeholder(
            deck_path=output_path,
            workbook_path=captable_workbook_path,
            output_path=output_path,
            slide_index=OVERVIEW_SLIDE_INDEX,
            placeholder_name=_CAP_TABLE_PLACEHOLDER,
            sheet_name=_CAP_TABLE_SHEET,
            source_range=_CAP_TABLE_RANGE,
        )

    # Paste the insider-ownership table into the ownership slide's left
    # "Insiders" placeholder (mirrors the cap-table insertion). The ownership
    # slide follows the Financial Summary section at its computed layout index.
    # When the ownership stage also ingested a Bloomberg export (Bloomberg
    # Output C14 populated), the right "Institutions" placeholder gets the
    # Select-Institutions block too; otherwise it stays a Bloomberg placeholder.
    institutions_inserted = False
    if ownership_workbook_path is not None:
        insert_excel_into_placeholder(
            deck_path=output_path,
            workbook_path=ownership_workbook_path,
            output_path=output_path,
            slide_index=layout.ownership,
            placeholder_name=_OWNERSHIP_PLACEHOLDER,
            sheet_name=_OWNERSHIP_SHEET,
            source_range=_OWNERSHIP_RANGE,
        )
        if _ownership_has_bloomberg(ownership_workbook_path):
            # The Select-Institutions picture range is read blind — verify its
            # sentinel anchors before pasting.
            verify_workbook_anchors(
                ownership_workbook_path,
                sheet=_OWNERSHIP_SHEET,
                anchors=OWNERSHIP_INSTITUTIONS_PICTURE_ANCHORS,
            )
            insert_excel_into_placeholder(
                deck_path=output_path,
                workbook_path=ownership_workbook_path,
                output_path=output_path,
                slide_index=layout.ownership,
                placeholder_name=_INSTITUTIONS_PLACEHOLDER,
                sheet_name=_OWNERSHIP_SHEET,
                source_range=_INSTITUTIONS_RANGE,
            )
            institutions_inserted = True

    _verify_pitch_output(
        output_path,
        cap_table_inserted=captable_workbook_path is not None,
        ownership_inserted=ownership_workbook_path is not None,
        institutions_inserted=institutions_inserted,
        expected_slides=layout.total,
    )
    return output_path


def _verify_pitch_output(
    path: Path,
    *,
    cap_table_inserted: bool = False,
    ownership_inserted: bool = False,
    institutions_inserted: bool = False,
    expected_slides: int | None = None,
) -> None:
    prs = Presentation(path)
    if expected_slides is not None and len(prs.slides) != expected_slides:
        raise ValueError(
            f"assembled pitch deck has {len(prs.slides)} slides; the slide-mix "
            f"layout expected {expected_slides}"
        )
    text = _all_text(prs)
    forbidden = ["[CLIENT NAME]", "[Date]"]
    leftovers = [token for token in forbidden if token in text]
    if leftovers:
        raise ValueError(f"assembled pitch deck still contains required-field placeholders: {leftovers}")
    required_placeholders = [
        "[Pie Chart Placeholder]",
        "[Placeholder for Metric #1 Chart]",
        "[Placeholder for Comps Chart]",
        # The precedent-transactions slide stays a chart placeholder, like comps
        # (no Excel→PowerPoint while Capital IQ can't be refreshed here).
        "[Placeholder for Precedents Chart]",
    ]
    missing = [token for token in required_placeholders if token not in text]
    if missing:
        raise ValueError(f"deferred placeholders were unexpectedly removed: {missing}")
    has_cap_placeholder = "[Cap Table Placeholder]" in text
    if cap_table_inserted and has_cap_placeholder:
        raise ValueError("slide 7 cap-table placeholder was not replaced by the Excel insertion stage")
    if not cap_table_inserted and not has_cap_placeholder:
        raise ValueError("slide 7 cap-table placeholder must remain when no workbook is supplied")
    has_insider_placeholder = "[Placeholder for Insider Ownership]" in text
    if ownership_inserted and has_insider_placeholder:
        raise ValueError("ownership slide insider placeholder was not replaced by the Excel insertion stage")
    if not ownership_inserted and not has_insider_placeholder:
        raise ValueError("ownership slide insider placeholder must remain when no workbook is supplied")
    # The ownership slide's institutional side is filled only when the ownership
    # stage ingested a Bloomberg export; otherwise it must stay a placeholder.
    has_institutions_placeholder = "[Placeholder for Institutional Ownership]" in text
    if institutions_inserted and has_institutions_placeholder:
        raise ValueError("ownership slide institutions placeholder was not replaced by the Excel insertion stage")
    if not institutions_inserted and not has_institutions_placeholder:
        raise ValueError("ownership slide institutions placeholder must remain without Bloomberg data")
