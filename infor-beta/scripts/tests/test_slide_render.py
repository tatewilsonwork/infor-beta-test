"""Unit tests for the slide-to-PNG overflow QA renderer.

Backend selection is tested on every platform (it is pure dispatch); the actual
render tests skip when the backend binary is absent.
"""

import sys
from pathlib import Path

import pytest

import slide_render
from excel_to_powerpoint import find_soffice
from slide_render import (
    BACKEND_ENV_VAR,
    BACKEND_LIBREOFFICE,
    BACKEND_POWERPOINT,
    render_deck_to_png,
)

_LIBRARY = Path("infor-beta/templates/INFOR Slide Library.pptx")


def _libreoffice_available() -> bool:
    return find_soffice() is not None


@pytest.fixture
def stub_backends(monkeypatch):
    """Replace both renderers with recorders so dispatch can be asserted."""
    calls: list[str] = []
    monkeypatch.setattr(
        slide_render, "_libreoffice_render", lambda *a, **k: calls.append(BACKEND_LIBREOFFICE) or []
    )
    monkeypatch.setattr(
        slide_render, "_powerpoint_com_render", lambda *a, **k: calls.append(BACKEND_POWERPOINT) or []
    )
    monkeypatch.delenv(BACKEND_ENV_VAR, raising=False)
    return calls


def _dummy_deck(tmp_path: Path) -> Path:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"not a real pptx - backend is stubbed")
    return deck


# ─── Backend selection (v0.5.35 render parity) ───────────────────────────────


def test_default_backend_is_libreoffice_even_on_windows(tmp_path: Path, stub_backends):
    """The v0.5.35 flip: Windows dev must render the way Cowork/Linux prod does.

    Before the flip this dispatched to PowerPoint COM on win32, which is why a
    production rendering bug could not be reproduced locally.
    """
    render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out")

    assert stub_backends == [BACKEND_LIBREOFFICE]


def test_powerpoint_backend_requires_explicit_opt_in(tmp_path: Path, stub_backends):
    render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out", backend=BACKEND_POWERPOINT)

    assert stub_backends == [BACKEND_POWERPOINT]


def test_env_var_selects_backend(tmp_path: Path, stub_backends, monkeypatch):
    monkeypatch.setenv(BACKEND_ENV_VAR, "PowerPoint")  # case/whitespace tolerant

    render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out")

    assert stub_backends == [BACKEND_POWERPOINT]


def test_explicit_backend_beats_env_var(tmp_path: Path, stub_backends, monkeypatch):
    monkeypatch.setenv(BACKEND_ENV_VAR, BACKEND_POWERPOINT)

    render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out", backend=BACKEND_LIBREOFFICE)

    assert stub_backends == [BACKEND_LIBREOFFICE]


def test_libreoffice_failure_does_not_fall_back_to_com(tmp_path: Path, monkeypatch):
    """No silent fallback — that divergence is the bug Phase A closed."""
    def boom(*a, **k):
        raise RuntimeError("LibreOffice not found on PATH")

    com_calls: list[str] = []
    monkeypatch.setattr(slide_render, "_libreoffice_render", boom)
    monkeypatch.setattr(
        slide_render, "_powerpoint_com_render", lambda *a, **k: com_calls.append("com") or []
    )
    monkeypatch.delenv(BACKEND_ENV_VAR, raising=False)

    with pytest.raises(RuntimeError, match="LibreOffice"):
        render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out")
    assert com_calls == [], "a failed LibreOffice render must not silently use COM"


def test_unknown_backend_raises(tmp_path: Path, stub_backends):
    with pytest.raises(ValueError, match="unknown slide-render backend"):
        render_deck_to_png(_dummy_deck(tmp_path), tmp_path / "out", backend="ghostscript")


def test_render_missing_deck_raises(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        render_deck_to_png(tmp_path / "nope.pptx", tmp_path / "out")


# ─── Real renders ────────────────────────────────────────────────────────────


@pytest.mark.skipif(not _libreoffice_available(), reason="LibreOffice not installed")
def test_render_selected_slides_returns_png_paths(tmp_path: Path):
    if not _LIBRARY.exists():
        pytest.skip("slide library template not present")
    pytest.importorskip("pypdfium2", reason="LibreOffice backend needs pypdfium2")

    out_dir = tmp_path / "png"
    paths = render_deck_to_png(_LIBRARY, out_dir, slide_indices=[0, 6])

    assert len(paths) == 2
    for p in paths:
        assert p.exists()
        assert p.suffix == ".png"
        assert p.stat().st_size > 0


@pytest.mark.skipif(sys.platform != "win32", reason="PowerPoint COM is Windows-only")
def test_powerpoint_com_render_still_works(tmp_path: Path):
    """The opt-in path stays reachable until Phase D deletes it."""
    if not _LIBRARY.exists():
        pytest.skip("slide library template not present")
    pytest.importorskip("win32com.client", reason="COM render needs pywin32 + PowerPoint")

    out_dir = tmp_path / "png"
    paths = render_deck_to_png(
        _LIBRARY, out_dir, slide_indices=[0, 6], backend=BACKEND_POWERPOINT
    )

    assert len(paths) == 2
    for p in paths:
        assert p.exists() and p.stat().st_size > 0
