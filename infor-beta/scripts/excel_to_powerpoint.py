"""Reusable Excel-to-PowerPoint insertion helpers.

The first implemented path replaces the earnings-update deck's slide 2
Macabacus placeholder with an editable PowerPoint table extracted from the
`Cap with Links` summary section of the cap-table workbook. Other insertion
paths can still record deferred intent while their concrete adapters are built.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from pptx_helpers import PALATINO, set_cell_text


def _format_cell_value(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.startswith("=") or text.startswith("#"):
        return "n/a"
    return text


def extract_cap_table_rows(workbook_path: Path | str) -> list[tuple[str, str]]:
    """Extract the visible summary rows used for PowerPoint cap-table insertion."""
    workbook = Path(workbook_path)
    if not workbook.exists():
        raise FileNotFoundError(f"workbook not found: {workbook}")

    wb_values = load_workbook(workbook, data_only=True, read_only=True)
    wb_formulas = load_workbook(workbook, data_only=False, read_only=True)
    if "Cap with Links" not in wb_values.sheetnames:
        raise KeyError("workbook does not contain required 'Cap with Links' sheet")

    ws_values = wb_values["Cap with Links"]
    ws_formulas = wb_formulas["Cap with Links"]
    rows: list[tuple[str, str]] = []
    title = _format_cell_value(ws_values["B13"].value)
    if not title or title == "n/a":
        title = "Capitalization Table"
    rows.append((title or "Capitalization Table", ""))

    for row_idx in range(15, 32):
        label = _format_cell_value(ws_values.cell(row=row_idx, column=2).value)
        if not label or label == "n/a":
            label = _format_cell_value(ws_formulas.cell(row=row_idx, column=2).value)
        value = _format_cell_value(ws_values.cell(row=row_idx, column=6).value)
        if label or value:
            rows.append((label, value))
    return rows


def insert_cap_table_into_placeholder(
    *,
    deck_path: Path | str,
    workbook_path: Path | str,
    output_path: Path | str | None = None,
    slide_index: int = 1,
    placeholder_name: str = "Rectangle 4",
) -> Path:
    """Replace an earnings-update cap-table placeholder with a PPT table from Excel."""
    deck = Path(deck_path)
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    rows_data = extract_cap_table_rows(workbook_path)
    if len(rows_data) < 2:
        raise ValueError("cap table workbook did not expose any summary rows for insertion")

    prs = Presentation(deck)
    slide = prs.slides[slide_index]
    placeholder = next((shape for shape in slide.shapes if shape.name == placeholder_name), None)
    if placeholder is None:
        raise KeyError(f"placeholder {placeholder_name!r} not found on slide {slide_index + 1}")

    left, top, width, height = placeholder.left, placeholder.top, placeholder.width, placeholder.height
    placeholder._element.getparent().remove(placeholder._element)

    table_shape = slide.shapes.add_table(len(rows_data), 2, left, top, width, height)
    table = table_shape.table
    table.columns[0].width = int(width * 0.68)
    table.columns[1].width = int(width * 0.32)

    for idx, (label, value) in enumerate(rows_data):
        set_cell_text(table.cell(idx, 0), label, size_pt=6.7)
        set_cell_text(table.cell(idx, 1), value, size_pt=6.7)
        for col_idx, cell in enumerate((table.cell(idx, 0), table.cell(idx, 1))):
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                for run in paragraph.runs:
                    run.font.name = PALATINO
                    run.font.size = Pt(6.7)
                    if idx == 0:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
        if idx == 0:
            for cell in (table.cell(idx, 0), table.cell(idx, 1)):
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0, 32, 96)

    out = Path(output_path) if output_path is not None else deck
    prs.save(out)
    return out


def record_insertion_intent(
    *,
    workbook_path: Path | str,
    deck_path: Path | str,
    placeholder_id: str,
    output_dir: Path | str,
) -> Path:
    """Write a small marker file documenting a deferred Excel→PPT insertion.

    This gives the conductor a typed side effect/run-log artefact in the POC
    while keeping actual chart/table transfer out of scope until the foundation
    is proven.
    """
    workbook = Path(workbook_path)
    deck = Path(deck_path)
    if not workbook.exists():
        raise FileNotFoundError(f"workbook not found: {workbook}")
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    marker = out / f"excel-to-powerpoint-{placeholder_id}.txt"
    marker.write_text(
        f"workbook_path={workbook}\n"
        f"deck_path={deck}\n"
        f"placeholder_id={placeholder_id}\n"
        "status=deferred_poc_placeholder\n",
        encoding="utf-8",
    )
    return marker
