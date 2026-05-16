"""POC boundary for reusable Excel-to-PowerPoint insertion.

The first slide-library POC proves that this is a separate stage from deck
assembly. Rich chart/table extraction is intentionally deferred; the function
records insertion intent and returns the existing deck path unchanged when no
concrete insertion adapter is supplied.
"""

from __future__ import annotations

from pathlib import Path


def record_insertion_intent(
    *,
    workbook_path: Path | str,
    deck_path: Path | str,
    placeholder_id: str,
    output_dir: Path | str,
) -> Path:
    """Write a small marker file documenting a deferred Excel→PPT insertion.

    This gives the conductor a typed side effect/run-log artefact in the POC
    while keeping actual chart/table transfer out of scope until the foundation
    is proven.
    """
    workbook = Path(workbook_path)
    deck = Path(deck_path)
    if not workbook.exists():
        raise FileNotFoundError(f"workbook not found: {workbook}")
    if not deck.exists():
        raise FileNotFoundError(f"deck not found: {deck}")
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    marker = out / f"excel-to-powerpoint-{placeholder_id}.txt"
    marker.write_text(
        f"workbook_path={workbook}\n"
        f"deck_path={deck}\n"
        f"placeholder_id={placeholder_id}\n"
        "status=deferred_poc_placeholder\n",
        encoding="utf-8",
    )
    return marker
