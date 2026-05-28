"""Assembler for the 14-slide INFOR slide-library POC deck."""

from __future__ import annotations

import re
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from pptx_helpers import find_shape, set_cell_text, set_text, write_bulleted_shape
from schemas import PitchDeckContent, SlidePlan


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


def _fill_market_entry_targets(slide, content: PitchDeckContent) -> None:
    """Fill the market-entry comparison table; logos stay deferred image placeholders."""
    if content.market_entry_market:
        set_text(
            find_shape(slide, "Title 1"),
            [f"Potential {content.market_entry_market} Market Entry Targets"],
        )
    if not content.market_entry_targets:
        return
    table = _table_shape(slide).table
    labels = content.market_entry_row_labels
    # Table row 0 is the blank logo/header row; data labels start at row 1.
    for i, label in enumerate(labels):
        row = i + 1
        if row >= len(table.rows):
            break
        set_cell_text(table.cell(row, 0), label, size_pt=8)
        for col, target in enumerate(content.market_entry_targets, start=1):
            set_cell_text(table.cell(row, col), target.cells[i], size_pt=8)
    for row in range(len(labels) + 1, len(table.rows)):
        for col in range(len(table.columns)):
            set_cell_text(table.cell(row, col), "", size_pt=8)


def assemble_pitch_deck(
    *,
    slide_plan_path: Path | str,
    content_path: Path | str,
    template_path: Path | str,
    output_dir: Path | str,
    captable_workbook_path: Path | str | None = None,
    comps_workbook_path: Path | str | None = None,
) -> Path:
    """Fill the canonical 14-slide blank INFOR slide-library deck.

    Complex Excel-to-PowerPoint chart/table insertion is intentionally delegated
    to `excel-to-powerpoint-infor`; this assembler preserves placeholders when
    no inserted artefact is available.
    """
    slide_plan = SlidePlan.model_validate_json(Path(slide_plan_path).read_text(encoding="utf-8"))
    content = PitchDeckContent.model_validate_json(Path(content_path).read_text(encoding="utf-8"))
    if slide_plan.deliverable_type != "pitch":
        raise ValueError("pitch deck assembler only supports pitch SlidePlan objects")
    if len(slide_plan.slides) != 14:
        raise ValueError("pitch deck POC expects the 14-slide INFOR Slide Library template")

    template = Path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"pitch library template not found: {template}")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"Pitch Deck - {_safe_name(content.client_name)}.pptx"

    prs = Presentation(template)

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

    # Slide 7 — public company overview; cap table/revenue chart placeholders stay unless insertion later replaces them.
    slide7 = prs.slides[6]
    set_text(find_shape(slide7, "Title 6"), [f"Introduction to {content.client_name}"])
    _write_flexible_bullets(
        find_shape(slide7, "TextBox 9"),
        content.company_overview_bullets,
    )

    # Slide 8 — financial metric labels only; charts remain placeholders.
    slide8 = prs.slides[7]
    metric_shapes = ["Rectangle 13", "Rectangle 12", "Rectangle 15", "Rectangle 14"]
    for shape_name, label in zip(metric_shapes, content.financial_metric_labels, strict=True):
        set_text(find_shape(slide8, shape_name), [label])

    # Slide 9 — concise acquirer risks and mitigants + tagline.
    slide9 = prs.slides[8]
    set_text(find_shape(slide9, "Text Placeholder 6"), [content.risks_tagline])
    table = _table_shape(slide9).table
    set_cell_text(table.cell(0, 0), "Considerations", size_pt=9)
    set_cell_text(table.cell(0, 1), "Mitigants", size_pt=9)
    max_rows = min(len(content.risk_mitigants), len(table.rows) - 1)
    for idx in range(max_rows):
        row = content.risk_mitigants[idx]
        set_cell_text(table.cell(idx + 1, 0), row.risk, size_pt=8)
        set_cell_text(table.cell(idx + 1, 1), "\n".join(row.mitigants), size_pt=8)
    for idx in range(max_rows + 1, len(table.rows)):
        set_cell_text(table.cell(idx, 0), "", size_pt=8)
        set_cell_text(table.cell(idx, 1), "", size_pt=8)

    # Slide 10 — comps takeaway; chart placeholder remains unless insertion later replaces it.
    slide10 = prs.slides[9]
    set_text(find_shape(slide10, "Text Placeholder 5"), [content.comps_takeaway])

    # Slide 11 — key investment highlights; placeholders remain unless content supplies them.
    _fill_investment_highlights(prs.slides[10], content)

    # Slide 12 — potential market-entry targets; logos remain deferred image placeholders.
    _fill_market_entry_targets(prs.slides[11], content)

    # Slides 13–14 (disclaimer, contact) are static. Do not touch.

    prs.save(output_path)
    _verify_pitch_output(output_path)
    return output_path


def _verify_pitch_output(path: Path) -> None:
    prs = Presentation(path)
    text = _all_text(prs)
    forbidden = ["[CLIENT NAME]", "[Date]"]
    leftovers = [token for token in forbidden if token in text]
    if leftovers:
        raise ValueError(f"assembled pitch deck still contains required-field placeholders: {leftovers}")
    required_placeholders = [
        "[Cap Table Placeholder]",
        "[Pie Chart Placeholder]",
        "[Placeholder for Metric #1 Chart]",
        "[Placeholder for Comps Chart]",
    ]
    missing = [token for token in required_placeholders if token not in text]
    if missing:
        raise ValueError(f"deferred placeholders were unexpectedly removed: {missing}")
