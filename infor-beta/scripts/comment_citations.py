"""In-artefact source citations on cell comments.

openpyxl supports exactly ONE `Comment` per cell, and the cap-table template
already spends that slot: F7 (FX rate) and F16 (share price) carry the CapIQ
refresh formula as a comment, which the analyst pastes back into the cell to
go live in Excel. When a skill hand-types a web-sourced value into such a cell
(captable Step 3b), the source citation therefore cannot be a second comment —
it must be APPENDED to the existing comment text, preserving the formula line.

`append_source_to_comment` does that safely for both cases: an existing comment
gains a trailing "Source: <url> — retrieved <YYYY-MM-DD>" line (original text
and author preserved), and a bare cell gains a fresh comment holding only the
source line. The workbook aggregator's openpyxl merge path copies comments
verbatim (`copy(cell.comment)`), so the appended citation survives into the
combined workbook alongside the CapIQ formula.
"""

from __future__ import annotations

import datetime as _dt

from openpyxl.comments import Comment

# Extra comment-box height per appended line so the citation stays visible
# when the analyst opens the comment in Excel (default box fits ~4 lines).
_LINE_HEIGHT_PX = 20


def source_line(url: str, retrieved: str | _dt.date) -> str:
    """Render the standard citation line: ``Source: <url> — retrieved <YYYY-MM-DD>``."""
    if isinstance(retrieved, _dt.date):
        retrieved = retrieved.isoformat()
    return f"Source: {url} — retrieved {retrieved}"


def append_source_to_comment(
    cell,
    url: str,
    retrieved: str | _dt.date,
    *,
    author: str = "INFOR",
) -> Comment:
    """Append a source-citation line to `cell`'s comment, creating one if absent.

    The existing comment text (e.g. the CapIQ refresh formula on the cap table's
    F7/F16) and its author are preserved — the citation is added as a new final
    line, never replacing anything. Returns the comment now attached to the cell.
    """
    line = source_line(url, retrieved)
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
