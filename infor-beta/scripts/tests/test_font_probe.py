"""The render oracle's typeface, probed rather than assumed.

`deck_repair` chooses every font size in a shipped deck from a measured
LibreOffice render, and the ladder was calibrated against real Palatino Linotype.
Nothing checked what the renderer actually resolved, so a substitution would have
been invisible: the deck converges, and every size in it is measured against the
wrong advance widths.

These tests pin the three answers the probe can give — metric-compatible,
substituted, and native-with-no-fontconfig — and one of them is a rule, not a
measurement: **"no fontconfig" is an ANSWER**. Windows has no fontconfig at all
and ships the real face, so that branch has to report `native` and be a pass. It
must never be able to read as a check that quietly did not run; three releases
shipped green behind guards that had rotted into exactly that (v0.5.36, v0.5.40,
v0.5.41).
"""

from __future__ import annotations

import pytest

import font_probe
from font_probe import (
    METHOD_FONTCONFIG,
    METHOD_NATIVE,
    METHOD_UNKNOWN,
    is_metric_compatible,
    probe_font_resolution,
)
from pptx_helpers import PALATINO


@pytest.fixture
def with_fontconfig(monkeypatch):
    """Stand in for an installed `fc-match` with a scripted answer.

    Returns a setter taking the `--format` output and, optionally, the bare
    `fc-match <family>` output for the no-`--format` fallback.
    """

    def install(formatted: str | None, plain: str | None = None):
        monkeypatch.setattr(
            font_probe.shutil,
            "which",
            lambda name: "/usr/bin/fc-match" if name == "fc-match" else None,
        )

        def fake(_fc_match, args):
            wants_format = any(arg.startswith("--format") for arg in args)
            return formatted if wants_format else plain

        monkeypatch.setattr(font_probe, "_fc_match", fake)

    return install


@pytest.fixture
def without_fontconfig(monkeypatch):
    """A host with no `fc-match`, and font directories the test controls."""

    def install(*font_dirs):
        monkeypatch.setattr(font_probe.shutil, "which", lambda _name: None)
        monkeypatch.setattr(font_probe, "_font_dirs", lambda: tuple(font_dirs))

    return install


# ─── The measured production answer ──────────────────────────────────────────


def test_p052_is_the_metric_compatible_verdict(with_fontconfig):
    """Cowork's real resolution, 2026-07-29: `P052-Roman.otf: "P052" "Roman"`.

    P052 / URW Palladio L is metric-compatible with Palatino, so the converge
    loop is calibrated correctly in production. This is the assertion that keeps
    that closed rather than restating it in prose.
    """
    with_fontconfig("P052\t/usr/share/fonts/opentype/urw-base35/P052-Roman.otf")

    resolution = probe_font_resolution()

    assert resolution.requested == PALATINO
    assert resolution.family == "P052"
    assert resolution.file.endswith("P052-Roman.otf")
    assert resolution.method == METHOD_FONTCONFIG
    assert resolution.metric_compatible is True
    assert resolution.ok
    assert "metric-compatible" in resolution.log_line()


def test_an_alias_list_counts_as_metric_compatible(with_fontconfig):
    """`%{family}` can name one face by several aliases; any match is the face."""
    with_fontconfig("URW Palladio L,P052\t/usr/share/fonts/type1/urw-base35/p052003l.pfb")

    assert probe_font_resolution().ok


def test_the_plain_output_is_parsed_when_format_is_unsupported(with_fontconfig):
    """fontconfig older than 2.11 has no `--format`; the default line still answers."""
    with_fontconfig(None, plain='P052-Roman.otf: "P052" "Roman"')

    resolution = probe_font_resolution()

    assert resolution.family == "P052"
    assert resolution.file == "P052-Roman.otf"
    assert resolution.method == METHOD_FONTCONFIG
    assert resolution.ok


# ─── A substitution warns, loudly ────────────────────────────────────────────


def test_a_dejavu_substitution_warns_and_says_what_to_install(with_fontconfig):
    """The worst case: everything renders, and every measured size is wrong.

    DejaVu Serif is the fallback a Linux image without `fonts-urw-base35` lands
    on. It is not metric-compatible, so the whole ladder would be measuring a deck
    the analyst never receives — which is why the verdict has to be a warning
    naming the fix, not a line in a log nobody greps.
    """
    with_fontconfig("DejaVu Serif\t/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf")

    resolution = probe_font_resolution()

    assert resolution.metric_compatible is False
    assert not resolution.ok
    message = resolution.log_line()
    assert "FONT SUBSTITUTION WARNING" in message
    assert "DejaVu Serif" in message
    assert "fonts-urw-base35" in message, "the warning must name the remedy"


@pytest.mark.parametrize(
    "family",
    ["DejaVu Serif", "Liberation Serif", "Times New Roman", "Nimbus Roman", "Carlito"],
)
def test_the_near_misses_are_not_treated_as_compatible(family):
    """Every family production is likely to fall back to has different metrics."""
    assert not is_metric_compatible(family)


@pytest.mark.parametrize(
    "family", ["P052", "  p052 ", "URW Palladio L", "Palatino", "Palatino Linotype"]
)
def test_the_palatino_metric_family_is_recognised_however_it_is_spelled(family):
    assert is_metric_compatible(family)


def test_a_fontconfig_that_cannot_answer_is_unknown_not_a_pass(with_fontconfig):
    with_fontconfig(None, plain=None)

    resolution = probe_font_resolution()

    assert resolution.method == METHOD_UNKNOWN
    assert resolution.metric_compatible is None
    assert not resolution.ok


# ─── No fontconfig is an answer ──────────────────────────────────────────────


def test_no_fontconfig_with_the_face_installed_reports_native(without_fontconfig, tmp_path):
    """The Windows dev box: no fontconfig at all, real Palatino Linotype installed.

    This branch is a *result*, and the assertions say so both ways — the method is
    `native` and it is emphatically not `unknown`. A "no fontconfig -> skip" here
    would be the stale-skip-guard failure mode one level down: green, and blind.
    """
    (tmp_path / "pala.ttf").write_bytes(b"not really a font")
    without_fontconfig(tmp_path)

    resolution = probe_font_resolution()

    assert resolution.method == METHOD_NATIVE
    assert resolution.method != METHOD_UNKNOWN
    assert resolution.family == PALATINO
    assert resolution.file == str(tmp_path / "pala.ttf")
    assert resolution.ok
    assert "native" in resolution.log_line()


def test_a_nested_font_directory_is_searched(without_fontconfig, tmp_path):
    """Linux and macOS install per-foundry; a direct hit is not the only hit."""
    nested = tmp_path / "truetype" / "msttcorefonts"
    nested.mkdir(parents=True)
    (nested / "pala.ttf").write_bytes(b"not really a font")
    without_fontconfig(tmp_path)

    assert probe_font_resolution().method == METHOD_NATIVE


def test_neither_fontconfig_nor_the_face_is_unknown(without_fontconfig, tmp_path):
    """Nothing to measure with and no way to find out — a warning, never a pass."""
    without_fontconfig(tmp_path)

    resolution = probe_font_resolution()

    assert resolution.method == METHOD_UNKNOWN
    assert resolution.metric_compatible is None
    assert not resolution.ok
    assert "UNKNOWN" in resolution.log_line()


# ─── This host, unmocked ─────────────────────────────────────────────────────


def test_the_probe_answers_on_this_host():
    """Run it for real, so the box the suite runs on states its own resolution.

    Deliberately not a skip and not a pinned family: what it asserts is that the
    probe always reaches a verdict and never returns a half-filled one. The
    verdict itself is printed for a reader, which is the whole point — the
    resolution appearing nowhere on disk is how this defect got found by
    reconstruction from a report.
    """
    resolution = probe_font_resolution()

    print(f"this host: {resolution.log_line()}")
    assert resolution.requested == PALATINO
    assert resolution.method in {METHOD_FONTCONFIG, METHOD_NATIVE, METHOD_UNKNOWN}
    if resolution.method == METHOD_UNKNOWN:
        assert resolution.metric_compatible is None
    else:
        assert resolution.family, "a resolved font must name its family"
        assert isinstance(resolution.metric_compatible, bool)
