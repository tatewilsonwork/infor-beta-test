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

import contextlib
import gc
import io
import os
import shutil
import subprocess
import sys
import tempfile
import time
from collections.abc import Iterator
from pathlib import Path
from typing import Any

from pptx import Presentation


_XL_SCREEN = 1
_XL_NORMAL = -4143
_XL_PICTURE = -4147


class _ClipboardPasteError(Exception):
    """Chart.Paste silently pasted nothing (CopyPicture/clipboard race).

    Deliberately NOT a RuntimeError: `_render_range_to_png` treats RuntimeError
    from the COM path as "Excel unavailable" and falls through to LibreOffice,
    which would mislabel this transient clipboard race.
    """


# ---------------------------------------------------------------------------
# The one Excel-COM instance owner (v0.5.38)
# ---------------------------------------------------------------------------
# Every `DispatchEx("Excel.Application")` in this plugin goes through
# `excel_com_app` — the same consolidation `find_soffice` did for LibreOffice,
# and for the same reason: four hand-rolled copies is how they diverged. The
# divergence found at v0.5.37 was the Quit guard — `slide_render.py` wrapped its
# `Quit()` in try/except so a raising Quit could not skip `CoUninitialize()`,
# while all four Excel sites called `excel.Quit()` bare in a `finally`, where a
# raising Quit escapes and the apartment is never torn down.
#
# What the 2026-07-27 measurements actually showed, so the next reader does not
# re-derive it (all counts are EXCEL.EXE deltas around a real render):
#
#   * `.Quit()` was already called at all four sites. "Add cleanup" was never
#     the fix.
#   * Holding `wb` / `ws` / `rng` bound across `Quit()` did NOT orphan the
#     process: `CoUninitialize()` tears the apartment down and releases the
#     proxies that belong to it, so the server exits anyway — measured delta 0
#     both with and without an explicit release, including with the raising
#     frame's traceback deliberately retained the way a pytest report retains
#     it. The explicit release below is therefore **defence in depth, not the
#     leak fix**: it keeps the release ordered and inside the apartment that
#     created the proxies, rather than relying on apartment teardown to mop up.
#   * The orphans on this box came from somewhere else entirely — see
#     `_dispatch_excel` below.
_ORPHAN_SIGNATURE_NOTE = (
    "A fully-started EXCEL.EXE with no window and no client may remain "
    "(~0.27 GB, ~60 invisible windows, no dialog — the observed orphan "
    "signature); it cannot be closed gracefully because no interface pointer "
    "was ever handed back. Do not force-kill it: that trips Office "
    "crash-resiliency and disables the analyst's CapIQ add-ins."
)


def _dispatch_excel(win32com_client: Any, purpose: str) -> Any:
    """Create a private Excel instance, or raise RuntimeError.

    **The measured orphan source.** When `CoCreateInstanceEx` fails with
    `-2146959355 "Server execution failed"` the launch itself already succeeded:
    Excel is up with ~60 threads and a full add-in set, but the interface handoff
    timed out, so no pointer comes back and there is nothing to `Quit()`. Two
    such orphans were produced in a row on 2026-07-27, and they are
    byte-for-byte the signature of the pre-existing orphan and of Phase B's 13
    (3.18 GB / 13 ≈ 0.245 GB each).

    This is not fixable from here — a graceful `Quit()` needs a pointer we never
    got, and the alternatives are both barred: force-killing corrupts Office
    add-in state, and attaching through the ROT would grab the analyst's own
    Excel and close their workbooks. So the failure is made **loud** instead of
    silent. Previously it was normalized to RuntimeError and the caller quietly
    degraded to LibreOffice, the test still passed, and the orphan was invisible
    — which is exactly how 13 accumulate unnoticed.
    """
    try:
        return win32com_client.DispatchEx("Excel.Application")
    except Exception as exc:
        print(
            f"[excel-com] Excel COM startup failed for {purpose}: {exc}. "
            f"{_ORPHAN_SIGNATURE_NOTE}",
            file=sys.stderr,
        )
        # Normalize so each caller's documented `except RuntimeError` fall-through
        # to its non-COM backend engages. (Failures PAST startup stay raw by
        # design — a mid-operation Excel error must not read as "no Excel".)
        raise RuntimeError(f"Excel COM unavailable: {exc}") from exc


@contextlib.contextmanager
def excel_com_app(
    *,
    purpose: str,
    visible: bool,
    park_offscreen: bool = False,
    hide_comment_indicators: bool = False,
) -> Iterator[Any]:
    """Own one private Excel instance for the duration of the block.

    Handles the apartment (`CoInitialize` / `CoUninitialize`), the `DispatchEx`,
    the standard app-level settings, and — the part that kept drifting — a
    **guarded** `Quit()` that cannot skip the apartment teardown.

    `visible`: renders need it (`CopyPicture(xlScreen)` and `Chart.Export`
    capture what the instance renders, and a recalc invalidates an invisible
    instance's render buffer, yielding a blank picture); the workbook merge does
    not. `park_offscreen` moves that visible window far off-screen so it never
    pops in front of the analyst. `hide_comment_indicators` suppresses the red
    cell-comment corner triangles that `CopyPicture` would otherwise bake into
    the picture.

    Callers MUST release their own COM children (workbook, worksheet, range,
    chart, …) in reverse creation order before leaving the block — this context
    manager can only see the app. See the notes above for why that is ordering
    hygiene rather than the leak fix.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            f"pywin32 is required for {purpose} (Windows + Microsoft Excel only)"
        ) from exc

    pythoncom.CoInitialize()
    app = None
    try:
        app = _dispatch_excel(win32com.client, purpose)
        app.Visible = visible
        app.DisplayAlerts = False
        if visible and park_offscreen:
            try:
                app.WindowState = _XL_NORMAL  # minimized windows don't render
                app.Top, app.Left = 4000, 6000
            except Exception:
                pass
        if hide_comment_indicators:
            try:
                app.DisplayCommentIndicator = 0  # xlNoIndicator
            except Exception:
                pass
        yield app
    finally:
        if app is not None:
            # Guarded, unlike the four bare `excel.Quit()` calls this replaces: a
            # Quit that raises (a modal add-in dialog will do it) must not escape
            # the finally and skip CoUninitialize, which would leave the thread
            # COM-initialized for the rest of the process.
            try:
                app.Quit()
            except Exception as exc:
                print(
                    f"[excel-com] Excel Quit() failed after {purpose}: {exc}. "
                    "Releasing the instance and tearing down the COM apartment "
                    "anyway.",
                    file=sys.stderr,
                )
            app = None
            gc.collect()
        # Last, and unconditionally: this is what actually releases the
        # apartment's proxies and lets the server exit.
        pythoncom.CoUninitialize()


def insert_excel_into_placeholder(
    *,
    deck_path: Path | str,
    workbook_path: Path | str,
    placeholder_name: str,
    source_range: str,
    sheet_name: str,
    output_path: Path | str | None = None,
    slide_index: int = 0,
) -> Path:
    """Replace a deck placeholder with a picture of an Excel range.

    Generic over (slide, placeholder, sheet, range): the cap table pastes
    ``Cap with Links!B15:F40`` into a slide-7 ``Rectangle 3``; the ownership
    slide pastes ``Ownership!B4:G17`` into its ``Rectangle 1`` insider
    placeholder. The picture is stretched to the placeholder's exact box.
    """
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
        except _ClipboardPasteError as exc:
            # The clipboard retry loop EXHAUSTED. `_ClipboardPasteError` is
            # deliberately not a RuntimeError (a transient clipboard race must
            # not be mislabeled "Excel unavailable"), but once the retries are
            # spent this render is not going to happen on the COM path — and
            # LibreOffice renders the same range fine. Degrading here is the
            # difference between a deck with a cap-table picture and an aborted
            # stage; letting it escape is why
            # `test_pitch_deck_inserts_cap_table_into_slide7` could go red on a
            # transient race.
            print(
                f"[excel-com] Excel clipboard retries exhausted ({exc}); "
                "falling back to the LibreOffice range renderer.",
                file=sys.stderr,
            )
    return _libreoffice_range_to_png(workbook, sheet_name, source_range)


def _excel_com_range_to_png(workbook: Path, sheet_name: str, source_range: str) -> io.BytesIO:
    """Open Excel via COM, copy the range, export as PNG via a temporary chart."""
    tmp_png_path: str | None = None
    try:
        with excel_com_app(
            purpose="COM-based cap-table insertion",
            visible=True,
            park_offscreen=True,
            hide_comment_indicators=True,
        ) as excel:
            wb = ws = rng = chart_obj = chart = None
            try:
                wb = excel.Workbooks.Open(str(workbook), ReadOnly=True, UpdateLinks=0)
                # openpyxl drops the cached value of every formula cell when the
                # cap-table skill saves the workbook, and the template is
                # manual-calc, so the EV cascade (market cap, net debt, Enterprise
                # Value, multiples — all in-workbook math, no CapIQ functions in
                # the picture range) loads blank. Force a recalc before
                # snapshotting or the image shows empty cells down through
                # Enterprise Value.
                try:
                    excel.CalculateFull()
                except Exception:
                    # A flaky recalc is no worse than the prior no-recalc
                    # behaviour; never let it abort the deck assembly.
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
                            # Chart.Paste silently no-ops when the shared Office
                            # clipboard hasn't been populated by CopyPicture yet
                            # (a race — empirically ~1-in-3 on a fast machine), and
                            # the export is then a blank white picture. A successful
                            # paste always lands the metafile as a chart Shape, so
                            # an empty Shapes collection means the paste failed:
                            # raise to engage this retry loop (fresh CopyPicture).
                            if chart.Shapes.Count == 0:
                                raise _ClipboardPasteError(
                                    f"Chart.Paste pasted nothing for {source_range} "
                                    "(Office clipboard not ready)"
                                )
                            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                                tmp_png_path = f.name
                            chart.Export(Filename=tmp_png_path, FilterName="PNG")
                        finally:
                            chart = None
                            try:
                                chart_obj.Delete()
                            finally:
                                chart_obj = None
                        break
                    except Exception as exc:
                        last_exc = exc
                        time.sleep(0.5 * (attempt + 1))
                else:
                    # Exhausted retries. Re-raise the underlying Excel/clipboard
                    # error itself (a com_error or _ClipboardPasteError, not a
                    # RuntimeError) so the caller does not mistake it for
                    # "pywin32 unavailable". `_render_range_to_png` catches both
                    # and degrades to LibreOffice.
                    raise last_exc  # type: ignore[misc]

                with open(tmp_png_path, "rb") as f:
                    buf = io.BytesIO(f.read())
                buf.seek(0)
                return buf
            finally:
                # Release the COM children in reverse creation order, inside the
                # apartment that created them, before `excel_com_app` quits the
                # instance and tears the apartment down.
                chart = chart_obj = rng = ws = None
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                    wb = None
    finally:
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


# Standard install locations for platforms whose LibreOffice installer does not
# put `soffice` on PATH. The Windows MSI never does — so a Windows dev box with
# LibreOffice correctly installed still fails `shutil.which("soffice")`, which
# would leave the LibreOffice-by-default slide renderer unusable on the very
# machine it exists to keep honest.
_SOFFICE_FALLBACK_PATHS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


def find_soffice() -> str | None:
    """Return a usable ``soffice`` command, or None when LibreOffice is absent.

    Prefers PATH (how Cowork / Linux prod resolves it), then falls back to the
    standard per-platform install locations.

    **The single LibreOffice locator — never call ``shutil.which("soffice")``
    directly.** A bare PATH lookup is what shipped in v0.5.35: the renderer was
    flipped to LibreOffice-by-default on every platform while five other call
    sites (this module's range renderer, three in ``financial_charts``, one in
    the aggregator, since deleted) still resolved through PATH only, so on a Windows
    dev box (MSI install, no PATH entry) they failed or silently degraded —
    inverting the dev/prod parity the flip existed to create. The drift lock in
    ``test_excel_to_powerpoint.py`` fails if a bare lookup reappears.
    """
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    return next((p for p in _SOFFICE_FALLBACK_PATHS if Path(p).is_file()), None)


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
        except subprocess.TimeoutExpired as exc:
            # A wedged soffice must degrade like a missing one: every caller's
            # graceful-degradation net catches RuntimeError only, so a raw
            # TimeoutExpired would abort the whole stage even though the durable
            # artefacts (e.g. the native workbook charts) are already saved.
            raise RuntimeError(
                f"LibreOffice {out_fmt!r} conversion timed out after "
                f"{exc.timeout:.0f}s"
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

    Requires LibreOffice (resolved by `find_soffice`) and the `pypdfium2`
    package. Raises RuntimeError with a clear message if either is
    missing — the conductor surfaces this to the analyst.
    """
    from openpyxl import load_workbook

    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice/libreoffice) not found on PATH or in the "
            "standard install locations; required for the non-Windows "
            "cap-table renderer. Install LibreOffice or run the conductor on "
            "a Windows machine with Excel."
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
