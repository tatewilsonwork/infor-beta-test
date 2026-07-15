"""Unit tests for the comment_citations helper (in-artefact source citations).

The cap-table skill hand-types web-sourced values into F7/F16, whose single
comment slot already holds the CapIQ refresh formula — the citation must be
APPENDED to that comment, never replace it (openpyxl allows one comment per
cell).
"""

import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

from comment_citations import append_source_to_comment, source_line


def test_source_line_from_string_date():
    assert (
        source_line("https://example.com/fx", "2026-07-15")
        == "Source: https://example.com/fx — retrieved 2026-07-15"
    )


def test_source_line_from_date_object():
    assert (
        source_line("https://example.com/fx", datetime.date(2026, 7, 15))
        == "Source: https://example.com/fx — retrieved 2026-07-15"
    )


def test_append_preserves_existing_formula_comment():
    # The cap-table case: F7 already carries the CapIQ refresh formula as its
    # comment. Appending the citation must keep the formula as the first line
    # and the original author.
    ws = Workbook().active
    ws["F7"] = 1.36
    ws["F7"].comment = Comment("=IQ_FX_RATE(...)", "INFOR")

    append_source_to_comment(ws["F7"], "https://example.com/fx", "2026-07-15")

    assert ws["F7"].comment.text == (
        "=IQ_FX_RATE(...)\nSource: https://example.com/fx — retrieved 2026-07-15"
    )
    assert ws["F7"].comment.author == "INFOR"


def test_append_creates_comment_when_cell_has_none():
    ws = Workbook().active
    ws["A1"] = 42.0
    append_source_to_comment(ws["A1"], "https://example.com/quote", "2026-07-15")
    assert ws["A1"].comment is not None
    assert ws["A1"].comment.text == "Source: https://example.com/quote — retrieved 2026-07-15"
    assert ws["A1"].comment.author == "INFOR"


def test_appended_comment_round_trips_through_save(tmp_path: Path):
    # The skill saves the workbook after writing — the two-line comment must
    # survive an openpyxl save/load cycle intact.
    wb = Workbook()
    ws = wb.active
    ws["F16"] = 42.18
    ws["F16"].comment = Comment("=IQ_PRICE_CLOSE(...)", "INFOR")
    append_source_to_comment(ws["F16"], "https://example.com/quote", "2026-07-15")
    path = tmp_path / "cap.xlsx"
    wb.save(path)

    ws2 = load_workbook(path).active
    assert ws2["F16"].comment is not None
    assert ws2["F16"].comment.text == (
        "=IQ_PRICE_CLOSE(...)\nSource: https://example.com/quote — retrieved 2026-07-15"
    )
