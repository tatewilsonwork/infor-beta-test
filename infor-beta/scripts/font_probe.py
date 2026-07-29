"""What typeface the render oracle actually got — probed, never assumed.

`deck_repair` decides every font size and autofit scale in a shipped deck from a
measured LibreOffice render, and that ladder (`TEXT_SCALES`, `TABLE_SIZE_DROPS`)
was calibrated against **real Palatino Linotype** on the Windows dev box. If the
render host substitutes a typeface with different metrics, the loop keeps
measuring perfectly — it just measures a deck the analyst will never see.

Nothing checked. `pptx_helpers.PALATINO` was the only declaration of the typeface
anywhere, and a substitution leaves no trace: the PDF converts, the ink measures,
the deck converges, and every point size in it was chosen against the wrong
advance widths. The first pitch run that needed the answer had to have it
reconstructed from a report, because the resolution appeared nowhere on disk.

The measured answer
-------------------
On Cowork (LibreOffice 26.2.4.2), 2026-07-29::

    $ fc-match "Palatino Linotype"
    P052-Roman.otf: "P052" "Roman"

with 8 palladio entries installed. P052 / URW Palladio L is metric-compatible
with Palatino, so the converge loop is sound as calibrated in production as well
as on dev. That closes what `CLAUDE.md` carried for four phases as its
highest-value open question — and this module is what keeps it closed, by
re-answering the question on every converge instead of pinning the answer in
prose that cannot notice a font package being dropped from an image.

No fontconfig is an ANSWER, not a skipped check
-----------------------------------------------
Windows has no fontconfig at all and ships real Palatino Linotype, so the absence
of `fc-match` there resolves to `METHOD_NATIVE` and names the font file it found
in the system font directories. That distinction is deliberate: three releases
shipped green behind skip guards that had stopped meaning anything (v0.5.36,
v0.5.40, v0.5.41 — the last hiding a production-breaking regression), so
"no fontconfig" must never be able to read as "the probe did not run". Only a
host with neither fontconfig nor an installed copy of the face reports
`METHOD_UNKNOWN`, and that is a warning, not a pass.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from pptx_helpers import PALATINO

METHOD_FONTCONFIG = "fontconfig"
METHOD_NATIVE = "native"
METHOD_UNKNOWN = "unknown"

#: Families whose advance widths match Palatino's, so a render measured in one is
#: a valid measurement of the other. URW Palladio L is the original metric clone;
#: `P052` is the same face under fontconfig's post-2020 name, and TeX Gyre Pagella
#: is derived from it. Anything NOT in here is treated as a substitution that
#: invalidates the ladder — including the near misses production is most likely to
#: fall back to (DejaVu Serif, Liberation Serif, Nimbus Roman, Times New Roman,
#: Carlito), none of which share Palatino's metrics.
METRIC_COMPATIBLE_FAMILIES = frozenset(
    {
        "palatino",
        "palatino linotype",
        "palatino lt std",
        "p052",
        "urw palladio l",
        "palladio",
        "tex gyre pagella",
    }
)

#: The file each requested family ships as, for the no-fontconfig path. Only the
#: faces the decks actually ask for need an entry; a family with none is reported
#: on its name alone.
_NATIVE_FONT_FILES: dict[str, tuple[str, ...]] = {
    "palatino linotype": ("pala.ttf", "palai.ttf", "palab.ttf", "palabi.ttf"),
    "palatino": ("Palatino.ttc", "pala.ttf"),
}

_FC_TIMEOUT_S = 15.0

# `fc-match`'s default output line: `P052-Roman.otf: "P052" "Roman"`. Parsed only
# when `--format` is unavailable (fontconfig older than 2.11).
_FC_DEFAULT_LINE = re.compile(r'^(?P<file>[^:]+):\s*"(?P<family>[^"]*)"')


def _normalize(family: str) -> str:
    return " ".join(family.strip().lower().split())


def is_metric_compatible(family: str) -> bool:
    """Whether a resolved family can stand in for Palatino when measuring ink."""
    return _normalize(family) in METRIC_COMPATIBLE_FAMILIES


@dataclass(frozen=True)
class FontResolution:
    """What the render host resolved a requested typeface to.

    ``metric_compatible`` is tri-state on purpose: ``None`` means the probe could
    not find out, which is a different (and worse) thing than ``False`` — a known
    substitution can be fixed by installing a font package, an unknown one cannot
    even be diagnosed.
    """

    requested: str
    family: str | None
    file: str | None
    method: str
    metric_compatible: bool | None

    @property
    def ok(self) -> bool:
        """True only when the ladder's calibration provably still holds."""
        return self.metric_compatible is True

    def log_line(self) -> str:
        """One line for the stage log — recorded on every converge, pass or fail."""
        where = f" ({self.file})" if self.file else ""
        if self.ok:
            return (
                f'font oracle — "{self.requested}" resolves to "{self.family}"{where} '
                f"via {self.method}; metric-compatible with Palatino, so the measured "
                f"font ladder is calibrated against the metrics the analyst will see"
            )
        if self.metric_compatible is False:
            return (
                f'FONT SUBSTITUTION WARNING — "{self.requested}" resolves to '
                f'"{self.family}"{where} via {self.method}, which is NOT '
                f"metric-compatible with Palatino. deck_repair chooses every font "
                f"size and autofit scale from a render in this face, so the whole "
                f"ladder is measuring a deck that will not be delivered. Install a "
                f"Palatino-metric family on the render host — fonts-urw-base35 "
                f"provides P052 / URW Palladio L — and re-validate the ladder"
            )
        return (
            f'FONT RESOLUTION UNKNOWN — "{self.requested}" could not be traced: no '
            f"fontconfig (fc-match) on this host and no matching font file in its "
            f"font directories, so whatever the renderer substitutes is invisible "
            f"here. deck_repair's measurements may be calibrated against the wrong "
            f"metrics; install fontconfig or the face itself to find out"
        )


def _font_dirs() -> tuple[Path, ...]:
    """Where an installed font file would be, per platform.

    A function rather than a constant so a test can point it at a directory it
    controls and exercise the native branch on any platform.
    """
    if sys.platform.startswith("win"):
        windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
        local = os.environ.get("LOCALAPPDATA")
        dirs = [windir / "Fonts"]
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
        return tuple(dirs)
    if sys.platform == "darwin":
        return (Path("/System/Library/Fonts"), Path("/Library/Fonts"), Path.home() / "Library/Fonts")
    return (
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path.home() / ".fonts",
        Path.home() / ".local/share/fonts",
    )


def _installed_font_file(family: str) -> Path | None:
    """The first on-disk file for `family`, searching the platform's font dirs.

    A direct hit first (Windows puts every face straight in `Fonts\\`), then one
    recursive pass per candidate name for the per-foundry subdirectories Linux and
    macOS use.
    """
    names = _NATIVE_FONT_FILES.get(_normalize(family))
    if not names:
        return None
    for directory in _font_dirs():
        for name in names:
            direct = directory / name
            try:
                if direct.is_file():
                    return direct
                nested = next(directory.rglob(name), None)
            except OSError:
                continue
            if nested is not None:
                return nested
    return None


def _fc_match(fc_match: str, args: list[str]) -> str | None:
    """Run `fc-match`, returning its stdout or None when it could not answer."""
    try:
        completed = subprocess.run(
            [fc_match, *args], capture_output=True, timeout=_FC_TIMEOUT_S, check=True
        )
    except (OSError, subprocess.SubprocessError):
        return None
    return completed.stdout.decode(errors="replace").strip() or None


def _fontconfig_resolution(fc_match: str, family: str) -> FontResolution:
    """Ask fontconfig what it would actually hand the renderer.

    `%{family}` can be a comma-separated list of aliases for one face
    (`URW Palladio L,P052`), so every alias is tested and any metric-compatible
    one is enough — the face is the same file either way.
    """
    resolved: str | None = None
    file: str | None = None

    formatted = _fc_match(fc_match, [f"--format=%{{family}}\t%{{file}}", family])
    if formatted:
        parts = formatted.splitlines()[0].split("\t")
        resolved = parts[0].strip() or None
        file = parts[1].strip() if len(parts) > 1 and parts[1].strip() else None
    else:
        # fontconfig without `--format`: parse the default `file: "Family" "Style"`.
        plain = _fc_match(fc_match, [family])
        match = _FC_DEFAULT_LINE.match(plain.splitlines()[0]) if plain else None
        if match:
            resolved = match.group("family").strip() or None
            file = match.group("file").strip() or None

    if resolved is None:
        return FontResolution(family, None, file, METHOD_UNKNOWN, None)
    aliases = [alias for alias in (a.strip() for a in resolved.split(",")) if alias]
    compatible = any(is_metric_compatible(alias) for alias in aliases)
    return FontResolution(family, resolved, file, METHOD_FONTCONFIG, compatible)


def probe_font_resolution(family: str = PALATINO) -> FontResolution:
    """Resolve `family` the way the render host will, and say whether that is safe.

    Cheap enough (one `fc-match`, ~10 ms) to run on every converge and
    deliberately uncached: a cached verdict is a prose claim with extra steps, and
    the whole point is that the answer is a property of the host and can change
    under us.
    """
    fc_match = shutil.which("fc-match")
    if fc_match is not None:
        return _fontconfig_resolution(fc_match, family)

    installed = _installed_font_file(family)
    if installed is not None:
        return FontResolution(
            family, family, str(installed), METHOD_NATIVE, is_metric_compatible(family)
        )
    return FontResolution(family, None, None, METHOD_UNKNOWN, None)
