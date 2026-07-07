"""Shared PDF text extraction with a garble-detecting OCR fallback.

Several skills (`ltm-metrics`, `captable`, `ownership`) read numbers out of
analyst-attached filing PDFs. Most PDFs carry a clean text layer, but some —
scanned statements, or PDFs whose embedded fonts use CID encodings with no
`ToUnicode` map — extract as **garbage**: empty pages, U+FFFD replacement
characters, private-use-area glyph ids, or CJK-looking mojibake. Each skill had
independently re-discovered this and bolted on its own tesseract fallback.

This module centralises that pattern so it behaves identically everywhere:

    text -> detect garble -> (if garbled) render pages + OCR -> clean text

Pipeline:

  1. `extract_text_pdfium(path)` pulls the embedded text layer via ``pypdfium2``
     (already a dependency on the non-Windows / Cowork runtime).
  2. `is_garbled(text)` decides whether that text is usable (a pure string
     heuristic — see its docstring; it is the unit-tested core).
  3. If garbled (or `prefer_ocr=True`), `ocr_pdf(path)` renders each page to an
     image and runs the ``tesseract`` CLI over it.

`extract_pdf_text(path)` ties the three together and returns a `PdfText` carrying
the text, the method used (``"text"`` / ``"ocr"``), and whether the text layer was
garbled. Every external dependency degrades gracefully:

  - ``pypdfium2`` missing  -> `PdfiumUnavailableError` (text step yields "").
  - ``tesseract`` missing  -> `OcrUnavailableError`; `extract_pdf_text` then
    returns the (garbled) text layer flagged ``garbled=True`` rather than raising,
    so a CI box without tesseract — or the analyst's machine — never crashes the
    stage. Callers should check `PdfText.garbled` and surface it.

Nothing here imports tesseract/Pillow at module load — those are pulled in lazily
inside the OCR path, so importing this module (and unit-testing `is_garbled`) has
no heavy dependencies.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

# Fraction of "implausible" characters (see `_is_plausible_char`) among the
# non-whitespace characters above which a text layer is considered garbled.
# A clean English filing scores ~0.0; a CID-scrambled one scores near 1.0, so a
# generous 0.20 cut cleanly separates them without tripping on the odd stray glyph.
_GARBLE_THRESHOLD = 0.20

# Below this many non-whitespace characters there isn't enough signal to trust the
# ratio, so the page is treated as garbled/empty and pushed to OCR. The filings we
# read are many KB of text, so a near-empty extraction means the text layer failed.
_MIN_TEXT_CHARS = 16

# Render scale for the OCR fallback (1.0 = 72 dpi). ~300 dpi reads small statement
# type reliably without exploding memory.
_OCR_RENDER_SCALE = 300 / 72


class PdfExtractError(RuntimeError):
    """Base class for PDF-extraction failures in this module."""


class PdfiumUnavailableError(PdfExtractError):
    """``pypdfium2`` is not importable (text extraction / rendering unavailable)."""


class OcrUnavailableError(PdfExtractError):
    """The ``tesseract`` OCR binary is not on PATH (OCR fallback unavailable)."""


@dataclass(frozen=True)
class PdfText:
    """Result of :func:`extract_pdf_text`.

    - ``text``    — the extracted text (OCR'd text when ``method == "ocr"``).
    - ``method``  — ``"text"`` (embedded layer) or ``"ocr"`` (rendered + OCR'd).
    - ``garbled`` — whether the *embedded text layer* was judged garbled. When
      ``method == "ocr"`` this is True (that is why OCR ran). When OCR was needed
      but unavailable, ``method`` stays ``"text"`` and ``garbled`` stays True so the
      caller knows the returned text is unreliable.
    """

    text: str
    method: str
    garbled: bool


# ---------------------------------------------------------------------------
# Garble detection — the unit-tested core (pure function of a string)
# ---------------------------------------------------------------------------
_PLAUSIBLE_PUNCT = "—–’‘“”…•·€£¥™®©°§¶"


def _is_plausible_char(ch: str) -> bool:
    """True if `ch` is a character we'd expect in an English financial filing.

    Plausible: ASCII whitespace, printable ASCII, accented Latin (Latin-1
    Supplement + Latin Extended-A), and a small set of common typographic
    punctuation/symbols. Everything else — private-use glyph ids, CJK/other
    scripts produced by CID mis-mapping, and the U+FFFD replacement char — is
    implausible and counts toward the garble ratio.
    """
    if ch in "\t\n\r\f\v ":
        return True
    o = ord(ch)
    if 0x20 <= o <= 0x7E:          # printable ASCII
        return True
    if 0xA0 <= o <= 0x017F:        # Latin-1 Supplement + Latin Extended-A
        return True
    if ch in _PLAUSIBLE_PUNCT:
        return True
    return False


def garble_ratio(text: str) -> float:
    """Fraction of non-whitespace characters that are implausible (0.0–1.0).

    Empty / whitespace-only text returns 1.0 (maximally garbled). This is the
    raw signal behind :func:`is_garbled`.
    """
    if not text:
        return 1.0
    non_ws = [ch for ch in text if not ch.isspace()]
    if not non_ws:
        return 1.0
    implausible = sum(1 for ch in non_ws if not _is_plausible_char(ch))
    return implausible / len(non_ws)


def is_garbled(text: str | None, *, threshold: float = _GARBLE_THRESHOLD) -> bool:
    """Heuristic: is `text` an unusable (garbled / empty) PDF text layer?

    True when the text is missing, has too few characters to be a real statements
    page (`_MIN_TEXT_CHARS`), or its :func:`garble_ratio` exceeds `threshold`.
    False for ordinary English text — including number-heavy financial tables,
    which are still mostly digits/punctuation/Latin letters and so score ~0.0.
    """
    if text is None:
        return True
    non_ws_len = sum(1 for ch in text if not ch.isspace())
    if non_ws_len < _MIN_TEXT_CHARS:
        return True
    return garble_ratio(text) > threshold


# ---------------------------------------------------------------------------
# Text-layer extraction (pypdfium2)
# ---------------------------------------------------------------------------
def _load_pdfium():
    try:
        import pypdfium2 as pdfium  # noqa: WPS433 (lazy import is intentional)
    except ImportError as exc:  # pragma: no cover - exercised via monkeypatch
        raise PdfiumUnavailableError(
            "pypdfium2 is required to read the PDF text layer; it is installed on "
            "the Cowork/Linux runtime (`pypdfium2 ; sys_platform != 'win32'`)."
        ) from exc
    return pdfium


def extract_text_pdfium(path: Path | str, *, max_pages: int | None = None) -> str:
    """Extract the embedded text layer from a PDF via pypdfium2.

    Returns the concatenated page text (pages separated by form feeds). Raises
    `PdfiumUnavailableError` if pypdfium2 is not importable, `FileNotFoundError`
    if the path is missing.
    """
    pdf_path = Path(path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    pdfium = _load_pdfium()
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        n = len(pdf)
        if max_pages is not None:
            n = min(n, max_pages)
        chunks: list[str] = []
        for i in range(n):
            page = pdf[i]
            textpage = page.get_textpage()
            try:
                chunks.append(textpage.get_text_range())
            finally:
                textpage.close()
                page.close()
        return "\f".join(chunks)
    finally:
        pdf.close()


# ---------------------------------------------------------------------------
# OCR fallback (tesseract CLI over rendered pages)
# ---------------------------------------------------------------------------
def _tesseract_path() -> str | None:
    """Absolute path to the ``tesseract`` binary, or None if not on PATH."""
    return shutil.which("tesseract")


def ocr_available() -> bool:
    """True when the tesseract OCR binary is available."""
    return _tesseract_path() is not None


def _render_pages_to_pngs(path: Path, out_dir: Path, *, max_pages: int | None) -> list[Path]:
    """Render each PDF page to a PNG in `out_dir`; return the PNG paths in order."""
    pdfium = _load_pdfium()
    pdf = pdfium.PdfDocument(str(path))
    pngs: list[Path] = []
    try:
        n = len(pdf)
        if max_pages is not None:
            n = min(n, max_pages)
        for i in range(n):
            page = pdf[i]
            try:
                pil = page.render(scale=_OCR_RENDER_SCALE).to_pil().convert("RGB")
            finally:
                page.close()
            png = out_dir / f"page_{i:04d}.png"
            pil.save(png, format="PNG")
            pngs.append(png)
        return pngs
    finally:
        pdf.close()


def _tesseract_image(tesseract: str, png: Path) -> str:
    """Run tesseract over a single image, returning the recognised text."""
    proc = subprocess.run(
        [tesseract, str(png), "stdout", "--psm", "6"],
        check=True,
        capture_output=True,
        timeout=300,
    )
    return proc.stdout.decode("utf-8", errors="replace")


def ocr_pdf(path: Path | str, *, max_pages: int | None = None) -> str:
    """Render the PDF's pages and OCR them with tesseract; return the joined text.

    Raises `OcrUnavailableError` if the tesseract binary is not on PATH (so callers
    can degrade gracefully), `PdfiumUnavailableError` if pages can't be rendered,
    `FileNotFoundError` if the path is missing.
    """
    pdf_path = Path(path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    tesseract = _tesseract_path()
    if tesseract is None:
        raise OcrUnavailableError(
            "tesseract OCR binary not found on PATH; cannot OCR a garbled PDF. "
            "Install tesseract (e.g. `apt-get install tesseract-ocr`) to enable the "
            "fallback, or extract the figures from a clean copy of the filing."
        )
    with tempfile.TemporaryDirectory() as tmp:
        pngs = _render_pages_to_pngs(pdf_path, Path(tmp), max_pages=max_pages)
        return "\f".join(_tesseract_image(tesseract, png) for png in pngs)


# ---------------------------------------------------------------------------
# Public orchestrator
# ---------------------------------------------------------------------------
def extract_pdf_text(
    path: Path | str,
    *,
    prefer_ocr: bool = False,
    garble_threshold: float = _GARBLE_THRESHOLD,
    max_pages: int | None = None,
) -> PdfText:
    """Extract clean text from a PDF, falling back to OCR when the text is garbled.

    1. Pull the embedded text layer (unless `prefer_ocr`).
    2. If that text is garbled (`is_garbled`) — or `prefer_ocr` is set — render the
       pages and OCR them.
    3. If OCR is needed but tesseract is unavailable, return the (garbled) text
       layer with ``garbled=True`` instead of raising, so the stage degrades
       gracefully. Callers should inspect `PdfText.garbled`.

    `FileNotFoundError` is raised for a missing path; pypdfium2 being unavailable
    yields an empty text layer (which then routes to OCR).
    """
    pdf_path = Path(path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    text = ""
    if not prefer_ocr:
        try:
            text = extract_text_pdfium(pdf_path, max_pages=max_pages)
        except PdfiumUnavailableError:
            text = ""

    garbled = is_garbled(text, threshold=garble_threshold)
    if prefer_ocr or garbled:
        try:
            ocr_text = ocr_pdf(pdf_path, max_pages=max_pages)
            return PdfText(text=ocr_text, method="ocr", garbled=True)
        except OcrUnavailableError:
            # Degrade gracefully: hand back whatever text we have, flagged garbled.
            return PdfText(text=text, method="text", garbled=True)

    return PdfText(text=text, method="text", garbled=False)
