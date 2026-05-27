"""Reusable Excel-to-PowerPoint insertion helpers.

The cap-table insertion path renders the Excel source range as a picture via
Excel COM automation, then pastes it into the deck placeholder at the
placeholder's exact width and height. Tune the workbook's column widths and
row heights so the natural aspect ratio of the source range matches the
target placeholder; the script stretches the picture to fit either way.

Mechanism: `Range.CopyPicture(Format=xlPicture)` puts a metafile on the
Office clipboard; we then paste it into a temporary `ChartObject` sized to
the range, export the chart as PNG via `Chart.Export`, and feed those bytes
to python-pptx. The chart-export round-trip bypasses the system clipboard,
so Excel can stay invisible — no flashing window for the analyst.

Requires Microsoft Excel on Windows. `pywin32` is a runtime dependency.
"""

from __future__ import annotations

import io
import os
import tempfile
from pathlib import Path

from pptx import Presentation


_XL_SCREEN = 1
_XL_PICTURE = -4147


def insert_cap_table_into_placeholder(
    *,
    deck_path: Path | str,
    workbook_path: Path | str,
    output_path: Path | str | None = None,
    slide_index: int = 1,
    placeholder_name: str = "Rectangle 4",
    sheet_name: str = "Cap with Links",
    source_range: str = "B13:F31",
) -> Path:
    """Replace a deck placeholder with a picture of an Excel range."""
    deck = Path(deck_path).resolve()
    workbook = Path(workbook_path).resolve()
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    if not workbook.exists():
        raise FileNotFoundError(f"workbook not found: {workbook}")

    prs = Presentation(deck)
    slide = prs.slides[slide_index]
    placeholder = next((shape for shape in slide.shapes if shape.name == placeholder_name), None)
    if placeholder is None:
        raise KeyError(f"placeholder {placeholder_name!r} not found on slide {slide_index + 1}")
    left, top, width, height = placeholder.left, placeholder.top, placeholder.width, placeholder.height

    png_buffer = _excel_range_to_png(workbook, sheet_name, source_range)

    placeholder._element.getparent().remove(placeholder._element)
    slide.shapes.add_picture(png_buffer, left, top, width=width, height=height)

    out = Path(output_path) if output_path is not None else deck
    prs.save(out)
    return out


def _excel_range_to_png(workbook: Path, sheet_name: str, source_range: str) -> io.BytesIO:
    """Open Excel via COM, copy the range, export as PNG via a temporary chart."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for picture-based cap-table insertion "
            "(Windows + Microsoft Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    tmp_png_path: str | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(workbook), ReadOnly=True, UpdateLinks=0)
        try:
            ws = wb.Worksheets(sheet_name)
            rng = ws.Range(source_range)
            rng.CopyPicture(Appearance=_XL_SCREEN, Format=_XL_PICTURE)

            chart_obj = ws.ChartObjects().Add(Left=0, Top=0, Width=rng.Width, Height=rng.Height)
            try:
                chart = chart_obj.Chart
                chart.ChartArea.Border.LineStyle = 0
                chart.Paste()
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                    tmp_png_path = f.name
                chart.Export(Filename=tmp_png_path, FilterName="PNG")
            finally:
                chart_obj.Delete()

            with open(tmp_png_path, "rb") as f:
                buf = io.BytesIO(f.read())
            buf.seek(0)
            return buf
        finally:
            wb.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
        if tmp_png_path is not None and os.path.exists(tmp_png_path):
            os.unlink(tmp_png_path)


def record_insertion_intent(
    *,
    workbook_path: Path | str,
    deck_path: Path | str,
    placeholder_id: str,
    output_dir: Path | str,
) -> Path:
    """Write a small marker file documenting a deferred Excel→PPT insertion.

    Used by deferred POC adapters that don't yet ship a real insertion path;
    gives the conductor a typed side-effect artefact for the run log.
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
