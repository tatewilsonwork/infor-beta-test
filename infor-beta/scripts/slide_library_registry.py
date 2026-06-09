"""Simple registry for the INFOR slide-library deck.

The blank library is 16 entries: the original 14 plus an insider-ownership
slide (which follows the Financial Summary slide, before the Considerations /
Mitigants slide) and a precedent-transactions slide (which follows the
comparable-companies slide, before Key Investment Highlights).
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
    SlideLibraryEntry("comparable-companies", 11, "Comparable Companies Analysis"),
    # Precedent-transactions slide follows the comparable-companies slide; like
    # comps it stays a chart placeholder (no Excel→PowerPoint while Capital IQ
    # can't be refreshed) and carries a one-line takeaway.
    SlideLibraryEntry("precedent-transactions", 12, "Precedent Transactions Analysis"),
    SlideLibraryEntry("key-investment-highlights", 13, "Key Investment Highlights"),
    # One physical library slide; the wireframe/assembler repeat it two targets
    # per slide when a deal has more than two market-entry targets.
    SlideLibraryEntry("market-entry-targets", 14, "Potential [Market] Market Entry Targets"),
    SlideLibraryEntry("disclaimer", 15, "Disclaimer", static=True),
    SlideLibraryEntry("contact", 16, "Contact", static=True),
)


def load_slide_library_registry() -> list[SlideLibraryEntry]:
    """Return the canonical blank-library order for the deck."""
    return list(_ENTRIES)
