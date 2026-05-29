"""Reusable Excel-to-PowerPoint insertion helpers.

The cap-table insertion path renders the Excel source range as a picture and
pastes it into the deck placeholder at the placeholder's exact width and
height. Two render backends are wired in:

  - **Excel COM** (Windows + Excel) — we force a full recalc first (the
    workbook is saved by openpyxl with stripped caches and is manual-calc, so
    its formulas otherwise load blank), then `Range.CopyPicture(xlScreen,
    xlPicture)` puts a metafile on the Office clipboard; we paste it into a
    temporary `ChartObject` sized to the range and export the chart as PNG via
    `Chart.Export`. A recalc invalidates an invisible instance's render buffer
    (yielding a blank picture), so the instance runs visible but parked far
    off-screen.

  - **LibreOffice headless** (Cowork / Linux / macOS) — set the print area
    to the source range via openpyxl, convert the workbook to PDF with
    `soffice --headless --convert-to pdf`, render PDF page 1 to PIL via
    `pypdfium2`, and feed the bytes to python-pptx.

Tune the workbook's column widths and row heights so the natural aspect
ratio of the source range matches the target placeholder; the picture is
stretched to fit either way.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from pptx import Presentation


_XL_SCREEN = 1
_XL_NORMAL = -4143
_XL_PICTURE = -4147


def insert_cap_table_into_placeholder(
    *,
    deck_path: Path | str,
    workbook_path: Path | str,
    output_path: Path | str | None = None,
    slide_index: int = 1,
    placeholder_name: str = "Rectangle 3",
    sheet_name: str = "Cap with Links",
    source_range: str = "B15:F40",
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

    png_buffer = _render_range_to_png(workbook, sheet_name, source_range)

    placeholder._element.getparent().remove(placeholder._element)
    slide.shapes.add_picture(png_buffer, left, top, width=width, height=height)

    out = Path(output_path) if output_path is not None else deck
    prs.save(out)
    return out


def _render_range_to_png(workbook: Path, sheet_name: str, source_range: str) -> io.BytesIO:
    """Render an Excel range as PNG, preferring Excel COM on Windows.

    On Windows the Excel COM path matches what the analyst sees in Excel
    (fonts, formats, conditional formatting, cached CapIQ values). On other
    platforms — notably Claude Cowork's Linux sandbox — we fall back to a
    LibreOffice headless PDF round-trip. Fidelity is close but not pixel
    perfect; tune the workbook to render cleanly under both backends.
    """
    if sys.platform == "win32":
        try:
            return _excel_com_range_to_png(workbook, sheet_name, source_range)
        except RuntimeError:
            # Excel COM unavailable on this Windows machine (no Excel install
            # or pywin32 missing). Fall through to LibreOffice.
            pass
    return _libreoffice_range_to_png(workbook, sheet_name, source_range)


def _excel_com_range_to_png(workbook: Path, sheet_name: str, source_range: str) -> io.BytesIO:
    """Open Excel via COM, copy the range, export as PNG via a temporary chart."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for COM-based cap-table insertion "
            "(Windows + Microsoft Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    tmp_png_path: str | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        # CopyPicture(xlScreen) captures what the instance renders, and the
        # recalc below invalidates the render buffer of an invisible instance
        # (producing a blank picture). So run visible — but parked far
        # off-screen so the window doesn't pop in front of the analyst.
        excel.Visible = True
        excel.DisplayAlerts = False
        try:
            excel.WindowState = _XL_NORMAL  # minimized windows don't render
            excel.Top, excel.Left = 4000, 6000
        except Exception:
            pass
        wb = excel.Workbooks.Open(str(workbook), ReadOnly=True, UpdateLinks=0)
        try:
            # openpyxl drops the cached value of every formula cell when the
            # cap-table skill saves the workbook, and the template is manual-calc,
            # so the EV cascade (market cap, net debt, Enterprise Value, multiples
            # — all in-workbook math, no CapIQ functions in the picture range)
            # loads blank. Force a recalc before snapshotting or the image shows
            # empty cells down through Enterprise Value.
            try:
                excel.CalculateFull()
            except Exception:
                # A flaky recalc is no worse than the prior no-recalc behaviour;
                # never let it abort the deck assembly.
                pass

            ws = wb.Worksheets(sheet_name)
            rng = ws.Range(source_range)
            # CopyPicture uses the shared Office clipboard; retry a couple of
            # times in case another Excel instance momentarily holds it.
            last_exc: Exception | None = None
            for attempt in range(5):
                try:
                    rng.CopyPicture(Appearance=_XL_SCREEN, Format=_XL_PICTURE)
                    chart_obj = ws.ChartObjects().Add(
                        Left=0, Top=0, Width=rng.Width, Height=rng.Height
                    )
                    try:
                        chart = chart_obj.Chart
                        chart.ChartArea.Border.LineStyle = 0
                        chart.Paste()
                        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                            tmp_png_path = f.name
                        chart.Export(Filename=tmp_png_path, FilterName="PNG")
                    finally:
                        chart_obj.Delete()
                    break
                except Exception as exc:
                    last_exc = exc
                    time.sleep(0.5 * (attempt + 1))
            else:
                # Exhausted retries. Re-raise the underlying Excel/clipboard
                # error itself (a com_error, not a RuntimeError) so the caller
                # does not mistake it for "pywin32 unavailable" and fall through
                # to the LibreOffice renderer.
                raise last_exc  # type: ignore[misc]

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


# LibreOffice recalculates OOXML formulas on load only when told to. openpyxl
# writes the cap table with no cached values and the template is manual-calc, so
# without this the headless export prints every formula cell (market cap, net
# debt, Enterprise Value, the whole Financial/Valuation block) as 0/blank. The
# profile below sets Calc's "recalc OOXML/ODF on load" mode to 0 = Always, so the
# conversion recomputes them — and 0 = Always (not 2 = Prompt) is essential in
# headless mode, where a prompt would silently skip the recalc.
_LO_RECALC_XCU = """\
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
  <prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop>
 </item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
  <prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop>
 </item>
</oor:items>
"""


def _write_lo_recalc_profile(base_dir: Path) -> str:
    """Create a throwaway LibreOffice user profile that forces a formula recalc
    on load, and return its ``file://`` URI for ``-env:UserInstallation``.

    Self-contained under ``base_dir`` (LibreOffice reads
    ``<UserInstallation>/user/registrymodifications.xcu``), so the user's global
    LibreOffice profile is never touched.
    """
    profile = base_dir / "lo_profile"
    (profile / "user").mkdir(parents=True, exist_ok=True)
    (profile / "user" / "registrymodifications.xcu").write_text(
        _LO_RECALC_XCU, encoding="utf-8"
    )
    return profile.as_uri()


def _soffice_convert(soffice: str, src: Path, out_fmt: str, out_dir: Path) -> None:
    """Convert ``src`` to ``out_fmt`` with headless LibreOffice, recalculating
    in-workbook formulas on load via a throwaway recalc profile.

    ``out_fmt`` is a ``--convert-to`` target (e.g. ``"pdf"`` or
    ``"xlsx:Calc MS Excel 2007 XML"``). Raises RuntimeError on failure. CapIQ
    ``_xll.*`` cells are unknown to LibreOffice and resolve to ``#NAME?`` (the
    template's IFERROR wrappers degrade those to ``n/a``); the recalc does not
    crash on them and every in-workbook arithmetic cell still computes.
    """
    with tempfile.TemporaryDirectory() as prof_base:
        profile_uri = _write_lo_recalc_profile(Path(prof_base))
        try:
            subprocess.run(
                [
                    soffice,
                    f"-env:UserInstallation={profile_uri}",
                    "--headless",
                    "--nologo",
                    "--nodefault",
                    "--nofirststartwizard",
                    "--convert-to",
                    out_fmt,
                    "--outdir",
                    str(out_dir),
                    str(src),
                ],
                check=True,
                capture_output=True,
                timeout=180,
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(
                f"LibreOffice {out_fmt!r} conversion failed: "
                f"{exc.stderr.decode(errors='replace')}"
            ) from exc


def _libreoffice_range_to_png(workbook: Path, sheet_name: str, source_range: str) -> io.BytesIO:
    """Render an Excel range as PNG via LibreOffice headless PDF export.

    Strategy: open the workbook with openpyxl, hide every sheet except the
    target, set the target sheet's print area to the source range, force
    fit-to-1-page, and save to a temp copy. Convert that copy to PDF with
    headless LibreOffice — **forcing a recalculation on load** so the
    openpyxl-authored (cache-less), manual-calc cap-table formulas actually
    compute (see `_soffice_convert` / `_write_lo_recalc_profile`) — then render
    PDF page 1 via pypdfium2 at 200 DPI and return as PNG bytes.

    CapIQ `_xll.*` cells (forward consensus estimates) are unknown to LibreOffice
    and resolve to `#NAME?`, which the template's IFERROR wrappers degrade to
    `n/a`; the in-workbook arithmetic and the hardcoded LTM column still compute.

    Requires `soffice` (or `libreoffice`) on PATH and the `pypdfium2`
    package. Raises RuntimeError with a clear message if either is
    missing — the conductor surfaces this to the analyst.
    """
    from openpyxl import load_workbook

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice/libreoffice) not found on PATH; required "
            "for the non-Windows cap-table renderer. Install LibreOffice or "
            "run the conductor on a Windows machine with Excel."
        )
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError(
            "pypdfium2 is required for the non-Windows cap-table renderer; "
            "run `pip install pypdfium2`."
        ) from exc

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_xlsx = Path(tmp_dir) / "captable_print.xlsx"
        wb = load_workbook(workbook)
        if sheet_name not in wb.sheetnames:
            raise KeyError(
                f"sheet {sheet_name!r} not found in workbook (available: {wb.sheetnames})"
            )
        for name in wb.sheetnames:
            wb[name].sheet_state = "visible" if name == sheet_name else "hidden"
        ws = wb[sheet_name]
        ws.print_area = source_range
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25
        wb.save(tmp_xlsx)

        # Convert to PDF with recalc-on-load forced (the saved xlsx has no cached
        # formula values), so the EV cascade and metric blocks print computed.
        _soffice_convert(soffice, tmp_xlsx, "pdf", Path(tmp_dir))

        pdf_path = Path(tmp_dir) / "captable_print.pdf"
        if not pdf_path.exists():
            raise RuntimeError("LibreOffice produced no PDF output for the cap-table workbook")

        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            page = pdf[0]
            pil_image = page.render(scale=200 / 72).to_pil().convert("RGB")
            pil_image = _trim_white_margins(pil_image)
            buf = io.BytesIO()
            pil_image.save(buf, format="PNG")
            buf.seek(0)
            return buf
        finally:
            pdf.close()


def _trim_white_margins(image, threshold: int = 250):
    """Crop solid-white borders left by LibreOffice's print-area PDF export."""
    from PIL import ImageChops, Image

    bg = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, bg)
    # Treat near-white as background by quantising
    diff = diff.point(lambda p: 0 if p < (255 - threshold) else 255)
    bbox = diff.getbbox()
    return image.crop(bbox) if bbox else image
