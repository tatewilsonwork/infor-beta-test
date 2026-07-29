"""The shipped templates carry INFOR's palette, not Office's.

A theme-indexed colour is a *slot number*, not an RGB: ``<color theme="4"/>`` and
every ``accent1`` a chart or conditional format names resolve against whichever
``clrScheme`` the file's theme part happens to carry. So the theme decides the
palette every tab renders in, and a workbook can be wrong about its whole
appearance while every cell in it is right.

That is exactly what shipped. ``tools/build_deal_workbook_template.py`` assembles
the deal workbook by copying sheets into a workbook ``Excel.Workbooks.Add()``
created — which carries Office's default theme — and Excel resolves a copied
sheet's slots against the *destination*. The aggregator Phase D deleted used to
stamp ``INFORFG.thmx``; nothing replaced it, so from v0.5.41 to v0.5.51 the
shipped template carried Office 2024's scheme (accent1 ``156082``, dk2
``0E2841``, lt2 ``E8E8E8``, Aptos Narrow) and every deal copied it.

These assertions are the pin, and they run anywhere: no Excel, no LibreOffice,
one zip entry per file.

**Assert on the palette, never the theme's name.** All four source templates are
named ``Office Theme`` while carrying the INFOR colour scheme — an analyst built
them by applying INFOR's colours to a default-themed workbook, and the ``.thmx``
is the only file in the repo whose theme is *named* ``INFORFG``. A name assertion
would therefore fail on four correct files and pass on any Office-themed workbook
someone re-themed by hand.

**Why the existing fixtures did not catch this.** ``fixtures/pitch-workbook.xlsx``
descends from the aggregator era, so it was stamped with ``INFORFG`` when it was
generated and has carried the right palette ever since — a fixture frozen before
the regression cannot witness it. (``fixtures/earnings-update-workbook.xlsx`` is
older still: its accent1 is Excel 2007's ``4F81BD``.) Both are inputs to
assembler tests that never look at a colour, which is the general shape of the
gap: nothing downstream of the template read the theme, so nothing failed.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from deal_workbook import DEAL_WORKBOOK_TEMPLATE
from pptx_helpers import INFOR_ACCENTS
from template_layout import (
    CAP_TABLE_TEMPLATE,
    COMPS_TEMPLATE,
    INFOR_THEME,
    OWNERSHIP_TEMPLATE,
    PRECEDENTS_TEMPLATE,
    THEME_ACCENT_SLOTS,
    read_theme,
)

PLUGIN_ROOT = Path(__file__).resolve().parents[2]
TEMPLATES = PLUGIN_ROOT / "templates"
FIXTURES = Path(__file__).resolve().parent / "fixtures"

SOURCE_TEMPLATES = (CAP_TABLE_TEMPLATE, COMPS_TEMPLATE, OWNERSHIP_TEMPLATE, PRECEDENTS_TEMPLATE)

#: What an Office-themed workbook looks like — the palette that shipped from
#: v0.5.41 to v0.5.51, kept here so a re-appearance is named rather than merely
#: "not INFOR's".
OFFICE_2024_ACCENT1 = "156082"


def _accents(path: Path) -> list[str]:
    palette = read_theme(path).palette
    return [palette[slot] for slot in THEME_ACCENT_SLOTS]


def test_the_deal_workbook_template_carries_the_infor_palette():
    """The one workbook every deal copies resolves its colours against INFOR's."""
    template = TEMPLATES / DEAL_WORKBOOK_TEMPLATE
    theme = read_theme(template)

    assert _accents(template) == INFOR_ACCENTS, (
        f"{DEAL_WORKBOOK_TEMPLATE} accent1-6 are {_accents(template)}, expected "
        f"{INFOR_ACCENTS}. Every theme-indexed colour in all five tabs resolves "
        f"against this, so the workbook renders in the wrong palette with nothing "
        f"in any cell wrong. If accent1 is {OFFICE_2024_ACCENT1} this is the Phase D "
        f"regression again: the template was rebuilt without applying "
        f"{INFOR_THEME}. Re-run tools/build_deal_workbook_template.py."
    )
    # dk2/lt2 are the second background/text pair, and INFOR's are plain
    # black/white — Office 2024's navy-ish 0E2841 / grey E8E8E8 tint every table
    # header and banded fill built on them.
    assert theme.palette["dk2"] == "000000"
    assert theme.palette["lt2"] == "FFFFFF"


def test_the_deal_workbook_templates_theme_fonts_are_the_brand_fonts():
    """The scheme-linked fonts the copied styles resolve through, too.

    A cell format carrying ``<scheme val="minor"/>`` names no typeface either;
    an Office-themed deal workbook rendered those as Aptos Narrow.
    """
    built = read_theme(TEMPLATES / DEAL_WORKBOOK_TEMPLATE)
    brand = read_theme(TEMPLATES / INFOR_THEME)
    assert (built.major_font, built.minor_font) == (brand.major_font, brand.minor_font)


@pytest.mark.parametrize("template", SOURCE_TEMPLATES)
def test_each_source_template_carries_the_infor_palette(template: str):
    """The four templates the deal workbook is assembled FROM.

    They are where the theme-indexed colours come from, so a source template that
    lost the palette would put INFOR slot numbers over someone else's colours no
    matter what the build stamps on the destination.
    """
    path = TEMPLATES / template
    assert _accents(path) == INFOR_ACCENTS, f"{template} accent1-6 are {_accents(path)}"
    assert read_theme(path).palette["dk2"] == "000000"
    assert read_theme(path).palette["lt2"] == "FFFFFF"


def test_the_brand_theme_file_is_the_palette_the_build_stamps():
    """`INFORFG.thmx` and `INFOR_ACCENTS` are the same six colours.

    The bridge that keeps this file honest. The build tool verifies against the
    ``.thmx`` — deliberately, so the brand palette is declared in exactly one
    place and the tool cannot disagree with what it stamps — while these tests
    assert against ``pptx_helpers.INFOR_ACCENTS``, the constant the charts colour
    themselves from. This is the assertion that the two are one palette.
    """
    theme = read_theme(TEMPLATES / INFOR_THEME)
    assert _accents(TEMPLATES / INFOR_THEME) == INFOR_ACCENTS
    assert theme.name == "INFORFG"  # the only file in the repo whose theme is named
    assert theme.color_scheme == "INFOR (New)"


def test_the_pitch_fixture_could_not_have_caught_the_regression():
    """Why a green suite shipped an Office-themed template for eleven releases.

    ``pitch-workbook.xlsx`` is a frozen artefact of the aggregator era, when the
    stamping still happened, so it carries INFORFG regardless of what the shipped
    template carries today. Documented as an assertion rather than a comment so
    that regenerating it from a mis-themed template trips here — with a pointer
    to the test above, which is the one that actually pins the shipped binary.
    """
    fixture = FIXTURES / "pitch-workbook.xlsx"
    assert read_theme(fixture).name == "INFORFG", (
        "the pitch fixture was regenerated from a workbook that had lost INFOR's "
        "theme; the fixture is not the pin — "
        "test_the_deal_workbook_template_carries_the_infor_palette is."
    )
