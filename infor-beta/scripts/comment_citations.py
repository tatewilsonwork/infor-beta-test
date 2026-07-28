"""In-artefact source citations — the cell-comment VIEW of a provenance record.

openpyxl supports exactly ONE `Comment` per cell, and the cap-table template
already spends that slot: F7 (FX rate) and F16 (share price) carry the CapIQ
refresh formula as a comment, which the analyst pastes back into the cell to
go live in Excel. When a skill hand-types a web-sourced value into such a cell
(captable Step 3b), the source citation therefore cannot be a second comment —
it must be APPENDED to the existing comment text, preserving the formula line.

That append is all this module does now. **What a citation SAYS is decided in
`provenance.py`** (`FigureSource.render`), and this module takes the record, not
a sentence. Before Phase G it was the other way round: a skill composed the
sentence and the comment was the only place it existed, so the fields
(filing / statement / page) were a comma-joined convention and nothing outside
the workbook could read a figure's source at all. Two entry points spelled the
same citation two ways — `source_line` for the URL form, a hand-written string
everywhere else — which is the drift pair H1 collapsed on the intake side.

So the direction of travel is fixed: build a `FigureSource` (or a whole
`FigureProvenance`), record it in the stage's ledger, and let the comment be
rendered from it. Passing a bare string is a `TypeError` naming the fix — that
is deliberate, because a string citation is a citation nothing else can read.
"""

from __future__ import annotations

from openpyxl.comments import Comment

from provenance import FigureProvenance, FigureSource

# Extra comment-box height per appended line so the citation stays visible
# when the analyst opens the comment in Excel (default box fits ~4 lines).
_LINE_HEIGHT_PX = 20

#: The one prefix a citation line carries in an artefact. Kept here rather than in
#: `provenance` because it is a property of the comment rendering, not of the record.
SOURCE_PREFIX = "Source: "


def _require_record(source) -> FigureSource:
    if isinstance(source, FigureSource):
        return source
    raise TypeError(
        f"expected a FigureSource, got {type(source).__name__} ({source!r}). Since "
        f"Phase G a citation is RENDERED FROM a record rather than being one: build "
        f"FigureSource(filing=…, statement=…, page=…) — or "
        f"FigureSource(url=…, retrieved=…) for a web value — and record it in the "
        f"stage's ProvenanceLedger."
    )


def _append_comment_line(cell, line: str, author: str) -> Comment:
    """Append `line` to `cell`'s comment, creating a fresh comment if absent."""
    existing = cell.comment
    if existing is None:
        cell.comment = Comment(line, author)
    else:
        cell.comment = Comment(
            existing.text.rstrip("\n") + "\n" + line,
            existing.author or author,
            height=existing.height + _LINE_HEIGHT_PX,
            width=existing.width,
        )
    return cell.comment


def append_source_to_comment(
    cell,
    source: FigureSource,
    *,
    author: str = "INFOR",
) -> Comment:
    """Append one source's citation line to `cell`'s comment, creating one if absent.

    The line is `FigureSource.render()` behind `Source: `, so both citation forms
    — a filing/statement/page chain and a URL + retrieval date — come out of the
    same renderer:

        Source: FY2025 10-K, Consolidated Statements of Operations, p. 87
        Source: https://example.com/fx — retrieved 2026-07-15

    The existing comment text (e.g. the CapIQ refresh formula on the cap table's
    F7/F16) and its author are preserved — the citation is added as a new final
    line, never replacing anything. Returns the comment now attached to the cell.
    """
    return _append_comment_line(cell, SOURCE_PREFIX + _require_record(source).render(), author)


def cite_cell(
    cell,
    figure: FigureProvenance,
    *,
    author: str = "INFOR",
) -> Comment | None:
    """Render a whole figure's provenance onto its cell — one line per source.

    The record-level entry point: pass the `FigureProvenance` a builder just
    recorded and the cell it wrote, and the comment becomes that record's view.
    A derived figure with no direct sources (an LTM bridge total, whose provenance
    is its components' records) gets its derivation noted instead. Returns the
    comment now attached to the cell, or None when there was nothing to say.
    """
    comment = None
    for source in figure.sources:
        comment = append_source_to_comment(cell, source, author=author)
    if comment is None and figure.derivation:
        comment = _append_comment_line(cell, f"Derived: {figure.derivation}", author)
    return comment
