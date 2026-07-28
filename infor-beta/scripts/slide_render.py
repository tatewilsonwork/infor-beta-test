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

Phase D deleted the **PowerPoint COM** backend (`Slide.Export(path, "PNG")`),
which Phase A had already demoted to an explicit opt-in. Two reasons it is gone
rather than merely unused: it was a second renderer whose output no analyst ever
receives, and it is where the dev/prod divergence lived — keeping it reachable
kept a route back to measuring the wrong engine. A missing LibreOffice is now a
loud failure telling you to install it, which is the only correct outcome when
LibreOffice is the renderer production has.
"""

from __future__ import annotations

import atexit
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

BACKEND_LIBREOFFICE = "libreoffice"
BACKEND_ENV_VAR = "INFOR_SLIDE_RENDER_BACKEND"
CACHE_ENV_VAR = "INFOR_RENDER_CACHE"


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

    `backend` exists only to reject a stale caller: ``"libreoffice"`` is the one
    renderer, on every platform. ``INFOR_SLIDE_RENDER_BACKEND`` is still honoured
    so a leftover ``=libreoffice`` in an environment keeps working, and anything
    else — notably the removed ``"powerpoint"`` — raises rather than silently
    rendering through a different engine.
    """
    deck = Path(deck_path).resolve()
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    resolved = (backend or os.environ.get(BACKEND_ENV_VAR) or BACKEND_LIBREOFFICE).strip().lower()
    if resolved != BACKEND_LIBREOFFICE:
        raise ValueError(
            f"unknown slide-render backend {resolved!r}; only "
            f"{BACKEND_LIBREOFFICE!r} exists (Phase D deleted the PowerPoint-COM "
            f"backend)"
        )
    return _libreoffice_render(deck, out_dir, slide_indices, dpi)


# ─── Render a private copy, never the caller's file ──────────────────────────
# A renderer has no business holding the source open. LibreOffice drops a
# `.~lock.<name>#` beside it, and PowerPoint COM (deleted in Phase D) took it with
# a share mode that DENIED other readers outright. Once the suite went distributed
# that surfaced *confusingly*: another worker doing an ordinary
# `Presentation(LIBRARY)` at that moment failed, and failed as
# `PackageNotFoundError: Package not found` for a file plainly there, because
# `zipfile.is_zipfile` swallows the OSError and returns False. In a real run the
# same applies to the analyst's deck in the deal directory.
#
# Copying first costs milliseconds against a multi-second conversion.
def _private_copy(deck: Path, tmp_dir: Path) -> Path:
    """A copy of `deck` inside `tmp_dir`, keeping the filename.

    The name is preserved because LibreOffice derives the output PDF's name from
    it, and `_convert_to_pdf` finds that PDF by glob.
    """
    copy = tmp_dir / deck.name
    shutil.copyfile(deck, copy)
    return copy

# A conversion that starts is normally deterministic, but soffice occasionally
# exits non-zero (or writes no PDF) on a transient — a stale lock, a font-cache
# rebuild racing the run. Observed once across a 610-test suite, as a
# `render-unavailable` finding on a deck that renders fine in isolation. One
# retry, because the Phase B converge loop renders several times per assembly
# and a per-render coin flip would surface as an unconverged deck. A genuinely
# absent or wedged LibreOffice still fails after the retry, unchanged.
_CONVERT_ATTEMPTS = 2
_CONVERT_RETRY_PAUSE_S = 2.0

# ─── Private LibreOffice profile ─────────────────────────────────────────────
# `soffice` is single-instance PER USER PROFILE: a second invocation against the
# same profile is forwarded to the running one, and a `--convert-to` request
# arriving that way does not reliably complete. Measured on this dev box: four
# concurrent conversions sharing the default profile, 2 of 4 FAILED; four
# concurrent conversions with one profile each, 4 of 4 succeeded, and eight of
# eight at 8-way.
#
# So each process gets its own throwaway profile. Two things follow, and both
# are wanted independently of speed:
#
#   - Renders no longer collide with the analyst's own LibreOffice. Before this,
#     a conductor render while they had a document open could simply fail.
#   - Renders can run concurrently, which is what lets the test suite use more
#     than one core.
#
# Creating a profile costs ~4s, once per process. Rendering is unaffected: font
# resolution is a system concern, and a fresh profile carries no user font
# substitutions — if anything it is the more reproducible of the two. The Phase
# B contract's pinned-findings tests over the frozen fixtures are what hold that.
_PROFILE_DIR: Path | None = None
_PROFILE_PREFIX = "infor-lo-profile-"
# A LibreOffice profile is ~40 MB and an `atexit` cleanup does not run when the
# process is killed — an interrupted run leaves one behind. Sweep our own stale
# directories when creating a new one, so they cannot accumulate unbounded on a
# dev box. A day is well past any live run.
_STALE_AFTER_S = 24 * 60 * 60


def _sweep_stale(root: Path, prefix: str) -> None:
    """Best-effort removal of our own leftover directories from killed runs."""
    cutoff = time.time() - _STALE_AFTER_S
    try:
        candidates = list(root.glob(f"{prefix}*"))
    except OSError:
        return
    for path in candidates:
        try:
            if path.is_dir() and path.stat().st_mtime < cutoff:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _profile_arg() -> list[str]:
    """`-env:UserInstallation=...` for this process's private profile."""
    global _PROFILE_DIR
    if _PROFILE_DIR is None:
        root = Path(tempfile.gettempdir())
        _sweep_stale(root, _PROFILE_PREFIX)
        _PROFILE_DIR = Path(tempfile.mkdtemp(prefix=f"{_PROFILE_PREFIX}{os.getpid()}-"))
        atexit.register(shutil.rmtree, _PROFILE_DIR, True)
    return ["-env:UserInstallation=file:///" + str(_PROFILE_DIR).replace("\\", "/")]


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
                 *_profile_arg(),
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


# ─── Converted-PDF cache ──────────────────────────────────────────────────────
# Starting a `soffice --headless --convert-to pdf` costs ~2-3s before it looks at
# the file, and the Phase B contract renders constantly: `verify_deck` renders
# the deck and then a probe deck, and the converge loop repeats that per
# iteration. Across a pytest session the SAME deck content is converted over and
# over — every test that verifies a frozen fixture pays the full cost again, and
# a probe deck built from identical inputs is identical every time.
#
# So a converted PDF is cached, keyed by the deck's CONTENT. Sound because the
# rendered pages are a pure function of the presentation: identical content
# renders identically, and nothing on a page depends on the file's name or path.
#
# The key deliberately hashes the zip's members rather than the file's bytes.
# python-pptx stamps each member with the current time on save, so two decks
# python-pptx built from the same content have different bytes — which would
# make every generated deck a miss, and generated decks are most of them.
#
# The cache lives on disk rather than in memory so that SIBLING PROCESSES share
# it: point `INFOR_RENDER_CACHE_DIR` at a directory and every process using it
# converts a given deck at most once between them. The test suite does exactly
# that (see `tests/conftest.py`), which is what stops each of the six xdist
# workers from separately rendering the blank library and its attribution probe
# — a fixed per-worker cost that dominated the first distributed runs. Without
# the variable each process gets its own private directory, cleaned up at exit.
_PDF_CACHE: dict[str, Path] = {}
_CACHE_DIR: Path | None = None
_CACHE_DIR_IS_OURS = False
CACHE_DIR_ENV_VAR = "INFOR_RENDER_CACHE_DIR"
# Bounds the disk the cache can accumulate. A pitch deck's PDF is a few MB; 64
# covers a whole pytest session's distinct decks with room to spare, and a
# conductor run needs a handful.
_CACHE_MAX_ENTRIES = 64


def _cache_enabled() -> bool:
    return os.environ.get(CACHE_ENV_VAR, "1").strip().lower() not in {"0", "false", "no"}


def _deck_digest(deck: Path) -> str | None:
    """Content digest of a .pptx, ignoring zip timestamps; None if unreadable.

    Returns None rather than raising so an unreadable or non-zip file simply
    misses the cache and takes the normal conversion path (where the real error
    surfaces).
    """
    try:
        digest = hashlib.sha256()
        with zipfile.ZipFile(deck) as zf:
            for name in sorted(zf.namelist()):
                digest.update(name.encode("utf-8"))
                digest.update(zf.read(name))
        return digest.hexdigest()
    except Exception:
        return None


def _cache_dir() -> Path:
    """The cache directory — shared via the env var, else private to this process."""
    global _CACHE_DIR, _CACHE_DIR_IS_OURS
    if _CACHE_DIR is None:
        shared = os.environ.get(CACHE_DIR_ENV_VAR, "").strip()
        if shared:
            _CACHE_DIR = Path(shared)
            _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        else:
            root = Path(tempfile.gettempdir())
            _sweep_stale(root, "infor-render-cache-")
            _sweep_stale(root, "infor-suite-render-cache-")
            _CACHE_DIR = Path(tempfile.mkdtemp(prefix="infor-render-cache-"))
            _CACHE_DIR_IS_OURS = True
            # Only a directory we created is ours to delete: a shared one belongs
            # to whoever set the variable, and other processes are still using it.
            atexit.register(shutil.rmtree, _CACHE_DIR, True)
    return _CACHE_DIR


def _evict_if_needed(cache: Path) -> None:
    """Keep the cache under its entry cap, oldest first.

    Best-effort throughout: a sibling process may be reading or removing the same
    entry, and losing a cache entry only costs a re-conversion.
    """
    try:
        entries = sorted(cache.glob("*.pdf"), key=lambda p: p.stat().st_mtime)
    except OSError:
        return
    for stale in entries[: max(0, len(entries) - _CACHE_MAX_ENTRIES + 1)]:
        try:
            stale.unlink()
            _PDF_CACHE.pop(stale.stem, None)
        except OSError:
            pass


def _cache_store(digest: str, pdf_path: Path) -> Path:
    """Publish a freshly converted PDF into the cache and return its cached path.

    Published atomically via `os.replace` so a sibling process can never read a
    half-written PDF. Two processes converting the same deck at the same time
    both publish the same digest, and the content is identical, so last-writer
    -wins is not a conflict.
    """
    cache = _cache_dir()
    _evict_if_needed(cache)
    cached = cache / f"{digest}.pdf"
    staging = cache / f".{digest}.{os.getpid()}.part"
    try:
        shutil.copyfile(pdf_path, staging)
        os.replace(staging, cached)
    except OSError:
        staging.unlink(missing_ok=True)
        return pdf_path  # caching failed; the freshly converted PDF is still good
    _PDF_CACHE[digest] = cached
    return cached


def clear_render_cache() -> None:
    """Drop every cached PDF. For tests that need a conversion to actually run."""
    cache = _CACHE_DIR
    _PDF_CACHE.clear()
    if cache is None:
        return
    for path in cache.glob("*.pdf"):
        try:
            path.unlink()
        except OSError:
            pass


def _pdf_for(soffice: str, deck: Path, tmp_dir: Path) -> Path:
    """The deck's PDF, converted or served from the content cache.

    On a miss, converts a private copy rather than the caller's file (see
    `_private_copy`) — LibreOffice writes a `.~lock` beside whatever it opens,
    and that should not land in a deal directory or the plugin's templates.
    """
    digest = _deck_digest(deck) if _cache_enabled() else None
    if digest is not None:
        # The in-process index first, then the directory — a sibling process may
        # have published this deck since we last looked.
        cached = _PDF_CACHE.get(digest) or (_cache_dir() / f"{digest}.pdf")
        if cached.is_file():
            _PDF_CACHE[digest] = cached
            return cached
    staging = tmp_dir / "src"
    staging.mkdir(parents=True, exist_ok=True)
    pdf_path = _convert_to_pdf(soffice, _private_copy(deck, staging), tmp_dir)
    return _cache_store(digest, pdf_path) if digest is not None else pdf_path


def _libreoffice_render(
    deck: Path, out_dir: Path, slide_indices: list[int] | None, dpi: int
) -> list[Path]:
    from excel_to_powerpoint import find_soffice  # shared LibreOffice locator

    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice (soffice/libreoffice) not found; it is the ONLY slide "
            "renderer on every platform since Phase D deleted the PowerPoint-COM "
            "backend. Install LibreOffice — on Windows dev too, so local renders "
            "match production."
        )
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise RuntimeError("pypdfium2 is required for the LibreOffice slide renderer") from exc

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = _pdf_for(soffice, deck, Path(tmp_dir))

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
