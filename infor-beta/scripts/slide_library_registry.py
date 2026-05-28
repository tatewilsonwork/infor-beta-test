"""Simple registry for the 14-slide INFOR slide-library POC."""

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
    SlideLibraryEntry("acquirer-considerations-mitigants", 9, "Potential Perceived Acquiror Considerations and Mitigants"),
    SlideLibraryEntry("comparable-companies", 10, "Comparable Companies Financial Analysis"),
    SlideLibraryEntry("key-investment-highlights", 11, "Key Investment Highlights"),
    SlideLibraryEntry("market-entry-targets", 12, "Potential [Market] Market Entry Targets"),
    SlideLibraryEntry("disclaimer", 13, "Disclaimer", static=True),
    SlideLibraryEntry("contact", 14, "Contact", static=True),
)


def load_slide_library_registry() -> list[SlideLibraryEntry]:
    """Return the canonical blank-library order for the POC."""
    return list(_ENTRIES)


def get_entry(entry_id: str) -> SlideLibraryEntry:
    for entry in _ENTRIES:
        if entry.library_entry_id == entry_id:
            return entry
    raise KeyError(f"unknown slide-library entry id: {entry_id}")
