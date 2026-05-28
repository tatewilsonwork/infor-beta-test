"""Unit tests for the slide-to-PNG overflow QA renderer."""

import shutil
import sys
from pathlib import Path

import pytest

from slide_render import render_deck_to_png

_LIBRARY = Path("infor-beta/templates/INFOR Slide Library.pptx")


def _backend_available() -> bool:
    if sys.platform == "win32":
        try:
            import win32com.client  # noqa: F401

            return True
        except ImportError:
            pass
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


pytestmark = pytest.mark.skipif(
    not _backend_available(),
    reason="no slide-render backend (PowerPoint COM or LibreOffice) available",
)


def test_render_missing_deck_raises(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        render_deck_to_png(tmp_path / "nope.pptx", tmp_path / "out")


def test_render_selected_slides_returns_png_paths(tmp_path: Path):
    if not _LIBRARY.exists():
        pytest.skip("slide library template not present")
    if sys.platform != "win32":
        pytest.importorskip("pypdfium2", reason="LibreOffice backend needs pypdfium2")

    out_dir = tmp_path / "png"
    paths = render_deck_to_png(_LIBRARY, out_dir, slide_indices=[0, 6])

    assert len(paths) == 2
    for p in paths:
        assert p.exists()
        assert p.suffix == ".png"
        assert p.stat().st_size > 0
