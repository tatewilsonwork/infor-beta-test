"""Assembler for the INFOR slide-library pitch deck.

The blank library is 14 slides; the market-entry section grows across multiple
slides — two targets per slide — by cloning the library's market-entry slide, so
an assembled deck has ``14 + (market_entry_slides - 1)`` slides (e.g. 8 targets →
4 market-entry slides → 17-slide deck, disclaimer/contact at 16/17).
"""

from __future__ import annotations

import math
import re
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from excel_to_powerpoint import insert_cap_table_into_placeholder
from pptx_helpers import (
    clone_slide_after,
    delete_slide,
    find_shape,
    set_cell_text,
    set_text,
    write_bulleted_shape,
)
from schemas import PitchDeckContent, SlidePlan

# Zero-based index of the earnings-summary entry inserted into the shared
# 15-slide library. The pitch deck does not use it, so it is dropped on open,
# restoring the original 14-slide ordering this assembler's indices assume.
_EARNINGS_LIBRARY_SLIDE_INDEX = 7

# Market-entry slide index in the raw 15-slide library (before the earnings
# slide is dropped). Market-entry slides are cloned here BEFORE the delete so
# python-pptx allocates fresh, non-colliding slide part names.
_LIBRARY_MARKET_ENTRY_INDEX = 12

# Deck indices after the earnings slide is dropped (final 14-slide ordering).
_OVERVIEW_SLIDE_INDEX = 6          # slide 7 — public-company overview
_MARKET_ENTRY_SLIDE_INDEX = 11     # slide 12 — first market-entry slide

# Slide 7 cap-table placeholder; the picture covers the capitalization summary
# plus the Financial/Valuation metric rows (same range as the earnings overview).
_CAP_TABLE_PLACEHOLDER = "Rectangle 3"
_CAP_TABLE_RANGE = "B15:F40"


def _safe_name(value: str) -> str:
    safe = re.sub(r"[/\\:*?\"<>|]+", "-", value).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe or "Client"


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


def _table_shape(slide):
    for shape in slide.shapes:
        if getattr(shape, "has_table", False):
            return shape
    raise KeyError("table shape not found")


def _write_flexible_bullets(shape, bullets) -> None:
    """Write bullets, falling back to plain paragraphs if a library shape has no bullet glyph template."""
    items = [_bullet_tuple(bullet) for bullet in bullets]
    try:
        write_bulleted_shape(shape, items)
    except RuntimeError:
        set_text(shape, [bullet.text for bullet in bullets])


def _iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_all_shapes(shape.shapes)


def _fill_investment_highlights(slide, content: PitchDeckContent) -> None:
    """Fill the four numbered highlight quadrants; leave placeholders if no content."""
    if not content.investment_highlights:
        return
    quadrants: list[tuple[int, object, object]] = []
    for group in slide.shapes:
        if group.shape_type != MSO_SHAPE_TYPE.GROUP or not group.name.startswith("Group"):
            continue
        number = header_shape = body_shape = None
        for sub in _iter_all_shapes(group.shapes):
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


# Market-entry cell sizing: the library ships the label column white at 11 pt
# and the target value columns at 10 pt. The old code hardcoded a single 8 pt,
# which rendered the labels black and everything too small.
_ME_LABEL_SIZE = 11
_ME_VALUE_SIZE = 10
_ME_LABEL_COLOR = "FFFFFF"  # scheme bg1 (white) in the library

# Slide 9 Considerations/Mitigants table sizing. The library ships the header row
# at 12 pt and the body cells at 10 pt; the old code hardcoded 9 pt / 8 pt, which
# rendered noticeably smaller than the template.
_RISK_HEADER_SIZE = 12
_RISK_BODY_SIZE = 10


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


def _fill_footnote_currency(shape, letter: str) -> None:
    """Swap the ``[x]$MM`` currency-letter token in a library footnote, keeping
    the rest of the standardized source/note lines (and their formatting)."""
    lines = [p.text.replace("[x]", letter) for p in shape.text_frame.paragraphs]
    set_text(shape, lines)


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
    written white at 11 pt and the target value columns at 10 pt, matching the
    library style. Each populated column's logo box is relabelled
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
            _fill_footnote_currency(footnote, currency_letter)

    if not targets:
        return

    table = _table_shape(slide).table
    n_cols = len(table.columns)  # label column + target columns (3 in the library)
    # Table row 0 is the blank logo/header row; data labels start at row 1. With
    # the fixed 12-row structure every data row is populated — no blank rows.
    for i, label in enumerate(row_labels):
        row = i + 1
        if row >= len(table.rows):
            break
        set_cell_text(table.cell(row, 0), label, size_pt=_ME_LABEL_SIZE, color_hex=_ME_LABEL_COLOR)
        for col in range(1, n_cols):
            target = targets[col - 1] if (col - 1) < len(targets) else None
            value = target.cells[i] if target is not None else ""
            set_cell_text(table.cell(row, col), value, size_pt=_ME_VALUE_SIZE)
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


def assemble_pitch_deck(
    *,
    slide_plan_path: Path | str,
    content_path: Path | str,
    template_path: Path | str,
    output_dir: Path | str,
    captable_workbook_path: Path | str | None = None,
    comps_workbook_path: Path | str | None = None,
) -> Path:
    """Fill the INFOR slide-library pitch deck.

    The blank library is 14 slides; the market-entry section expands across
    multiple slides (two targets per slide) based on
    ``content.market_entry_targets``. The cap table is pasted into slide 7 when
    ``captable_workbook_path`` is supplied (via `excel-to-powerpoint`); other
    chart/table insertions remain deferred placeholders.
    """
    slide_plan = SlidePlan.model_validate_json(Path(slide_plan_path).read_text(encoding="utf-8"))
    content = PitchDeckContent.model_validate_json(Path(content_path).read_text(encoding="utf-8"))
    if slide_plan.deliverable_type != "pitch":
        raise ValueError("pitch deck assembler only supports pitch SlidePlan objects")
    if len(slide_plan.slides) < 14:
        raise ValueError("pitch deck expects at least the 14 base INFOR Slide Library entries")

    template = Path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"pitch library template not found: {template}")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"Pitch Deck - {_safe_name(content.client_name)}.pptx"

    # Footnote currency letter for the slide-7 + market-entry '[x]$MM' tokens,
    # derived from the cap table's output currency (None when no workbook).
    currency_letter = (
        _output_currency_letter(captable_workbook_path)
        if captable_workbook_path is not None
        else None
    )

    prs = Presentation(template)

    # Grow the market-entry section (two targets per slide) by cloning the
    # library's market-entry slide. Clone BEFORE dropping the earnings slide so
    # python-pptx allocates fresh, non-colliding slide part names.
    n_market_entry = (
        max(1, math.ceil(len(content.market_entry_targets) / 2))
        if content.market_entry_targets
        else 1
    )
    for _ in range(n_market_entry - 1):
        clone_slide_after(prs, _LIBRARY_MARKET_ENTRY_INDEX)

    # The shared library carries the earnings-update slide at index 7. Drop it
    # so the remaining slides keep the canonical pitch ordering.
    delete_slide(prs, _EARNINGS_LIBRARY_SLIDE_INDEX)

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
    # revenue pie stays a deferred placeholder.
    slide7 = prs.slides[_OVERVIEW_SLIDE_INDEX]
    set_text(find_shape(slide7, "Title 6"), [f"Introduction to {content.client_name}"])
    _write_flexible_bullets(
        find_shape(slide7, "TextBox 9"),
        content.company_overview_bullets,
    )
    if currency_letter is not None:
        _fill_footnote_currency(find_shape(slide7, "Text Placeholder 1"), currency_letter)

    # Slide 8 — financial metric labels only; charts remain placeholders.
    slide8 = prs.slides[7]
    metric_shapes = ["Rectangle 13", "Rectangle 12", "Rectangle 15", "Rectangle 14"]
    for shape_name, label in zip(metric_shapes, content.financial_metric_labels, strict=True):
        set_text(find_shape(slide8, shape_name), [label])

    # Slide 9 — concise acquirer risks and mitigants + tagline.
    slide9 = prs.slides[8]
    set_text(find_shape(slide9, "Text Placeholder 6"), [content.risks_tagline])
    table = _table_shape(slide9).table
    set_cell_text(table.cell(0, 0), "Considerations", size_pt=_RISK_HEADER_SIZE)
    set_cell_text(table.cell(0, 1), "Mitigants", size_pt=_RISK_HEADER_SIZE)
    max_rows = min(len(content.risk_mitigants), len(table.rows) - 1)
    for idx in range(max_rows):
        row = content.risk_mitigants[idx]
        set_cell_text(table.cell(idx + 1, 0), row.risk, size_pt=_RISK_BODY_SIZE)
        set_cell_text(table.cell(idx + 1, 1), "\n".join(row.mitigants), size_pt=_RISK_BODY_SIZE)
    for idx in range(max_rows + 1, len(table.rows)):
        set_cell_text(table.cell(idx, 0), "", size_pt=_RISK_BODY_SIZE)
        set_cell_text(table.cell(idx, 1), "", size_pt=_RISK_BODY_SIZE)

    # Slide 10 — comps takeaway; chart placeholder remains unless insertion later replaces it.
    slide10 = prs.slides[9]
    set_text(find_shape(slide10, "Text Placeholder 5"), [content.comps_takeaway])

    # Slide 11 — key investment highlights; placeholders remain unless content supplies them.
    slide11 = prs.slides[10]
    _fill_investment_highlights(slide11, content)
    if currency_letter is not None:
        _fill_footnote_currency(find_shape(slide11, "Text Placeholder 13"), currency_letter)

    # Slides 12+ — potential market-entry targets, two per slide. The section was
    # grown above; fill each slide with its pair and title it '(N of M)'.
    for j in range(n_market_entry):
        pair = content.market_entry_targets[2 * j : 2 * j + 2]
        _fill_market_entry_targets(
            prs.slides[_MARKET_ENTRY_SLIDE_INDEX + j],
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
        insert_cap_table_into_placeholder(
            deck_path=output_path,
            workbook_path=captable_workbook_path,
            output_path=output_path,
            slide_index=_OVERVIEW_SLIDE_INDEX,
            placeholder_name=_CAP_TABLE_PLACEHOLDER,
            source_range=_CAP_TABLE_RANGE,
        )

    _verify_pitch_output(output_path, cap_table_inserted=captable_workbook_path is not None)
    return output_path


def _verify_pitch_output(path: Path, *, cap_table_inserted: bool = False) -> None:
    prs = Presentation(path)
    text = _all_text(prs)
    forbidden = ["[CLIENT NAME]", "[Date]"]
    leftovers = [token for token in forbidden if token in text]
    if leftovers:
        raise ValueError(f"assembled pitch deck still contains required-field placeholders: {leftovers}")
    required_placeholders = [
        "[Pie Chart Placeholder]",
        "[Placeholder for Metric #1 Chart]",
        "[Placeholder for Comps Chart]",
    ]
    missing = [token for token in required_placeholders if token not in text]
    if missing:
        raise ValueError(f"deferred placeholders were unexpectedly removed: {missing}")
    has_cap_placeholder = "[Cap Table Placeholder]" in text
    if cap_table_inserted and has_cap_placeholder:
        raise ValueError("slide 7 cap-table placeholder was not replaced by the Excel insertion stage")
    if not cap_table_inserted and not has_cap_placeholder:
        raise ValueError("slide 7 cap-table placeholder must remain when no workbook is supplied")
