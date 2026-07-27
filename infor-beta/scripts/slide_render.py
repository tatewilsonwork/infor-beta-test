"""Render deck slides to PNG so the assembler stage can visually QA overflow.

Text overflow (bullets spilling past a divider, a metric label wrapping onto a
third line) is invisible to python-pptx — the XML is valid, the text just
doesn't fit. The deck-assembler workflow renders the overflow-prone slides to
PNG, the agent inspects them, and shrinks/autofits text until clean.

**LibreOffice headless is the default backend on every platform** — convert the
deck to PDF with `soffice --headless --convert-to pdf`, render the requested
pages via `pypdfium2`. Production (Cowork / Linux) has no other option, and
before v0.5.35 Windows dev rendered through PowerPoint COM instead, which meant
a production rendering bug could not be reproduced locally. One renderer
everywhere closes that gap.

The **PowerPoint COM** backend (`Slide.Export(path, "PNG")`, Windows +
PowerPoint only) is still reachable, but only on an explicit opt-in:

    render_deck_to_png(deck, out, backend="powerpoint")
    # or, for a whole session: INFOR_SLIDE_RENDER_BACKEND=powerpoint

There is deliberately **no automatic fallback** from LibreOffice to COM. A
silent fallback is what produced the dev/prod divergence in the first place: a
missing LibreOffice should be a loud failure telling you to install it, not a
quiet switch back to a renderer production does not have. Phase D deletes the
COM path outright.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

BACKEND_LIBREOFFICE = "libreoffice"
BACKEND_POWERPOINT = "powerpoint"
BACKEND_ENV_VAR = "INFOR_SLIDE_RENDER_BACKEND"


def render_deck_to_png(
    deck_path: Path | str,
    output_dir: Path | str,
    *,
    slide_indices: list[int] | None = None,
    dpi: int = 150,
    backend: str | None = None,
) -> list[Path]:
    """Render slides to PNG and return the image paths in slide order.

    `slide_indices` is zero-based; None renders every slide.

    `backend` selects the renderer: ``"libreoffice"`` (default, and the only
    backend production has) or ``"powerpoint"`` (Windows + PowerPoint COM,
    explicit opt-in only — Phase D removes it). Omitting it honours
    ``INFOR_SLIDE_RENDER_BACKEND`` before defaulting to LibreOffice.
    """
    deck = Path(deck_path).resolve()
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    resolved = (backend or os.environ.get(BACKEND_ENV_VAR) or BACKEND_LIBREOFFICE).strip().lower()
    if resolved == BACKEND_POWERPOINT:
        return _powerpoint_com_render(deck, out_dir, slide_indices)
    if resolved != BACKEND_LIBREOFFICE:
        raise ValueError(
            f"unknown slide-render backend {resolved!r}; "
            f"expected {BACKEND_LIBREOFFICE!r} or {BACKEND_POWERPOINT!r}"
        )
    return _libreoffice_render(deck, out_dir, slide_indices, dpi)


def _powerpoint_com_render(deck: Path, out_dir: Path, slide_indices: list[int] | None) -> list[Path]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for PowerPoint COM slide rendering") from exc

    pythoncom.CoInitialize()
    powerpoint = None
    paths: list[Path] = []
    try:
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        # PowerPoint refuses to open with the window fully hidden on some builds;
        # keep it minimized rather than Visible=False.
        presentation = powerpoint.Presentations.Open(
            str(deck), ReadOnly=True, WithWindow=False
        )
        try:
            count = presentation.Slides.Count
            indices = slide_indices if slide_indices is not None else list(range(count))
            for idx in indices:
                if idx < 0 or idx >= count:
                    continue
                png = out_dir / f"slide_{idx + 1}.png"
                presentation.Slides(idx + 1).Export(str(png), "PNG")
                paths.append(png)
        finally:
            presentation.Close()
    except Exception as exc:  # COM errors surface as generic pywintypes errors
        raise RuntimeError(f"PowerPoint COM render failed: {exc}") from exc
    finally:
        if powerpoint is not None:
            # PowerPoint's COM server is a SINGLETON — DispatchEx attaches to the
            # analyst's already-running instance when one exists, so an
            # unconditional Quit() would close their open presentations (and a
            # force-kill of the "leftover" process corrupts Office add-in state —
            # the CapIQ LoadBehavior=2 incident). Only quit an instance that has
            # nothing else open, i.e. one this render effectively owns.
            try:
                if powerpoint.Presentations.Count == 0:
                    powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return paths


# A conversion that starts is normally deterministic, but soffice occasionally
# exits non-zero (or writes no PDF) on a transient — a stale lock, a font-cache
# rebuild racing the run. Observed once across a 610-test suite, as a
# `render-unavailable` finding on a deck that renders fine in isolation. One
# retry, because the Phase B converge loop renders several times per assembly
# and a per-render coin flip would surface as an unconverged deck. A genuinely
# absent or wedged LibreOffice still fails after the retry, unchanged.
_CONVERT_ATTEMPTS = 2
_CONVERT_RETRY_PAUSE_S = 2.0


def _convert_to_pdf(soffice: str, deck: Path, tmp_dir: Path) -> Path:
    """Convert `deck` to a PDF in `tmp_dir` and return its path.

    Raises RuntimeError on failure — callers' degradation nets catch RuntimeError
    only, so a wedged soffice must degrade like a missing one rather than abort
    the stage.
    """
    problem = ""
    for attempt in range(1, _CONVERT_ATTEMPTS + 1):
        try:
            subprocess.run(
                [soffice, "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
                 "--convert-to", "pdf", "--outdir", str(tmp_dir), str(deck)],
                check=True, capture_output=True, timeout=300,
            )
            pdf_path = next(tmp_dir.glob("*.pdf"), None)
            if pdf_path is not None:
                return pdf_path
            problem = "LibreOffice produced no PDF output for the deck"
        except subprocess.CalledProcessError as exc:
            problem = f"LibreOffice PDF conversion failed: {exc.stderr.decode(errors='replace')}"
        except subprocess.TimeoutExpired as exc:
            problem = f"LibreOffice PDF conversion timed out after {exc.timeout:.0f}s"
        if attempt < _CONVERT_ATTEMPTS:
            print(
                f"slide_render: {problem} — retrying ({attempt}/{_CONVERT_ATTEMPTS})",
                file=sys.stderr,
            )
            time.sleep(_CONVERT_RETRY_PAUSE_S)
    raise RuntimeError(f"{problem} (after {_CONVERT_ATTEMPTS} attempts)")


def _libreoffice_render(
    deck: Path, out_dir: Path, slide_indices: list[int] | None, dpi: int
) -> list[Path]:
    from excel_to_powerpoint import find_soffice  # shared LibreOffice locator

    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice/libreoffice) not found on PATH; it is the "
            "default slide renderer on every platform. Install LibreOffice — on "
            "Windows dev too, so local renders match production. To render "
            "through PowerPoint instead, opt in explicitly with "
            f"backend={BACKEND_POWERPOINT!r} (Windows only; removed in Phase D)."
        )
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError("pypdfium2 is required for the LibreOffice slide renderer") from exc

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = _convert_to_pdf(soffice, deck, Path(tmp_dir))

        pdf = pdfium.PdfDocument(str(pdf_path))
        paths: list[Path] = []
        try:
            count = len(pdf)
            indices = slide_indices if slide_indices is not None else list(range(count))
            for idx in indices:
                if idx < 0 or idx >= count:
                    continue
                png = out_dir / f"slide_{idx + 1}.png"
                pdf[idx].render(scale=dpi / 72).to_pil().convert("RGB").save(png, format="PNG")
                paths.append(png)
        finally:
            pdf.close()
    return paths
