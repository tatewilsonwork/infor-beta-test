"""Phase A exit criterion: the two render backends agree on page geometry.

v0.5.35 made LibreOffice the default renderer on every platform so that a
production rendering bug can be reproduced on a Windows dev box. This test is
the executable form of that claim: render a frozen fixture deck through both
backends and assert they agree on slide count, raster aspect, and per-slide ink
geometry.

**Scope of the claim.** Byte-comparable PNGs are not achievable across two
rasterizers and are not asserted. What is asserted is that neither backend
crops, letterboxes, rescales, or drops content — i.e. that a geometric finding
measured on one backend's output means the same thing on the other's.

**What is deliberately NOT asserted: live text layout.** The two engines do not
lay out text identically even with the real Palatino Linotype installed on both
sides. Measured on the pitch fixture's overview slide (both on Windows, same
font file): LibreOffice wraps earlier, producing one extra line and a text block
~6% taller than PowerPoint's. The direction matters and is load-bearing for
Phase B — LibreOffice is the *conservative* renderer, so "fits under
LibreOffice" implies "fits in PowerPoint", but not the reverse. A parity test
that asserted text-block equality would be asserting something false.
"""

from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

from excel_to_powerpoint import find_soffice
from slide_render import BACKEND_LIBREOFFICE, BACKEND_POWERPOINT, render_deck_to_png

_FIXTURES = Path(__file__).parent / "fixtures"
_DECK = _FIXTURES / "pitch-deck.pptx"

# Cover (mostly shapes) + the dense overview slide (text, a pie, a cap-table
# picture) — the two extremes of what the deck contains.
_SLIDES = [0, 6]

_GRID = (40, 30)

# Tolerances, set from measurement rather than guessed. Observed worst case on
# this fixture: aspect 0.0000, ink bbox 0.0024, grid mean 0.0117, coverage
# 0.0043. Each tolerance is roughly 3x the observed worst case, which leaves
# room for antialiasing/hinting drift across LibreOffice versions without
# leaving room for a real layout divergence.
_ASPECT_TOL = 0.005
_BBOX_TOL = 0.010
_GRID_MEAN_TOL = 0.035
_COVERAGE_TOL = 0.015


def _both_backends_available() -> bool:
    if find_soffice() is None:
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


pytestmark = pytest.mark.skipif(
    not _both_backends_available(),
    reason="parity needs BOTH LibreOffice and PowerPoint COM (Windows dev only)",
)


def _ink_bbox_norm(png: Path) -> tuple[float, float, float, float] | None:
    """Bounding box of non-white content, as fractions of the page."""
    im = Image.open(png).convert("L")
    w, h = im.size
    bbox = im.point(lambda v: 255 if v < 250 else 0).getbbox()
    if bbox is None:
        return None
    left, top, right, bottom = bbox
    return (left / w, top / h, right / w, bottom / h)


def _grid_density(png: Path) -> list[float]:
    """Mean intensity per cell of a coarse grid over the page, 0.0–1.0."""
    im = Image.open(png).convert("L").resize(_GRID, Image.BOX)
    return [b / 255.0 for b in im.tobytes()]


@pytest.fixture(scope="module")
def renders(tmp_path_factory) -> dict[str, list[Path]]:
    if not _DECK.exists():
        pytest.skip(f"fixture deck not present: {_DECK}")
    root = tmp_path_factory.mktemp("parity")
    return {
        backend: render_deck_to_png(
            _DECK, root / backend, slide_indices=_SLIDES, backend=backend
        )
        for backend in (BACKEND_LIBREOFFICE, BACKEND_POWERPOINT)
    }


def test_both_backends_render_every_requested_slide(renders):
    for backend, paths in renders.items():
        assert len(paths) == len(_SLIDES), f"{backend} rendered {len(paths)} of {len(_SLIDES)}"
        for p in paths:
            assert p.exists() and p.stat().st_size > 0


def test_backends_agree_on_page_aspect(renders):
    """Same page shape from both — no crop, letterbox, or rescale."""
    deck_aspect = Emu(Presentation(_DECK).slide_width).inches / Emu(
        Presentation(_DECK).slide_height
    ).inches

    for i in range(len(_SLIDES)):
        aspects = {}
        for backend, paths in renders.items():
            w, h = Image.open(paths[i]).size
            aspects[backend] = w / h
        lo, pp = aspects[BACKEND_LIBREOFFICE], aspects[BACKEND_POWERPOINT]
        assert abs(lo - pp) < _ASPECT_TOL, f"slide {_SLIDES[i] + 1}: aspect {lo} vs {pp}"
        assert abs(lo - deck_aspect) < _ASPECT_TOL, (
            f"slide {_SLIDES[i] + 1}: rendered aspect {lo} != deck aspect {deck_aspect}"
        )


def test_backends_agree_on_ink_bounding_box(renders):
    """Content occupies the same region of the page in both renders."""
    for i, idx in enumerate(_SLIDES):
        boxes = {b: _ink_bbox_norm(paths[i]) for b, paths in renders.items()}
        assert all(v is not None for v in boxes.values()), f"slide {idx + 1}: a render was blank"

        lo, pp = boxes[BACKEND_LIBREOFFICE], boxes[BACKEND_POWERPOINT]
        deltas = [abs(a - b) for a, b in zip(lo, pp)]
        assert max(deltas) < _BBOX_TOL, (
            f"slide {idx + 1}: ink bbox differs by {max(deltas):.4f} of the page "
            f"(LibreOffice {lo}, PowerPoint {pp})"
        )


def test_backends_agree_on_ink_distribution(renders):
    """Ink lands in the same places, at the same density, across the page.

    Coarser than a pixel diff and finer than a bounding box: this is what
    catches a backend dropping a chart, mis-placing a picture, or reflowing a
    whole block, without tripping on antialiasing.
    """
    for i, idx in enumerate(_SLIDES):
        grids = {b: _grid_density(paths[i]) for b, paths in renders.items()}
        lo, pp = grids[BACKEND_LIBREOFFICE], grids[BACKEND_POWERPOINT]

        deltas = [abs(a - b) for a, b in zip(lo, pp)]
        mean_delta = sum(deltas) / len(deltas)
        assert mean_delta < _GRID_MEAN_TOL, (
            f"slide {idx + 1}: mean ink-density delta {mean_delta:.4f} exceeds "
            f"{_GRID_MEAN_TOL} — the backends disagree about page layout"
        )

        cover = {b: sum(1 - v for v in g) / len(g) for b, g in grids.items()}
        assert abs(cover[BACKEND_LIBREOFFICE] - cover[BACKEND_POWERPOINT]) < _COVERAGE_TOL, (
            f"slide {idx + 1}: total ink coverage differs — {cover}"
        )
