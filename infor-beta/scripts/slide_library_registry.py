"""Simple registry for the INFOR slide-library deck.

The blank library is 15 entries: the original 14 plus an insider-ownership
slide, which follows the Financial Summary slide (before the Considerations /
Mitigants slide).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SlideLibraryEntry:
    library_entry_id: str
    slide_number: int
    title: str
    static: bool = False


_ENTRIES: tuple[SlideLibraryEntry, ...] = (
    SlideLibraryEntry("pitch-cover", 1, "Cover"),
    SlideLibraryEntry("executive-summary", 2, "Executive Summary"),
    SlideLibraryEntry("infor-overview", 3, "INFOR Financial Inc. Overview", static=True),
    SlideLibraryEntry("infor-ma-advisor", 4, "Canada’s Top Independent M&A Advisor", static=True),
    SlideLibraryEntry("infor-key-highlights", 5, "INFOR Key Highlights", static=True),
    SlideLibraryEntry("section-divider", 6, "Section Divider"),
    SlideLibraryEntry("public-company-overview", 7, "Introduction to [Client Name]"),
    SlideLibraryEntry("financial-summary", 8, "Financial Summary"),
    # Insider-ownership slide (Canadian public targets) follows the Financial
    # Summary slide: the left "Insiders" placeholder is replaced by a picture of
    # the ownership workbook; the right "Institutions" side stays a
    # Bloomberg-sourced placeholder.
    SlideLibraryEntry("insider-ownership", 9, "Ownership"),
    SlideLibraryEntry("acquirer-considerations-mitigants", 10, "Potential Perceived Acquiror Considerations and Mitigants"),
    SlideLibraryEntry("comparable-companies", 11, "Comparable Companies Financial Analysis"),
    SlideLibraryEntry("key-investment-highlights", 12, "Key Investment Highlights"),
    # One physical library slide; the wireframe/assembler repeat it two targets
    # per slide when a deal has more than two market-entry targets.
    SlideLibraryEntry("market-entry-targets", 13, "Potential [Market] Market Entry Targets"),
    SlideLibraryEntry("disclaimer", 14, "Disclaimer", static=True),
    SlideLibraryEntry("contact", 15, "Contact", static=True),
)


def load_slide_library_registry() -> list[SlideLibraryEntry]:
    """Return the canonical blank-library order for the deck."""
    return list(_ENTRIES)
