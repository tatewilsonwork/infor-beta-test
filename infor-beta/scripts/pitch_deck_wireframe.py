"""Build the 14-slide SlidePlan for the INFOR slide-library POC."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from schemas import Company, SlideEntry, SlidePlan
from slide_library_registry import load_slide_library_registry


def _company_name(company: Company | dict[str, Any]) -> str:
    if isinstance(company, dict):
        return str(company.get("legal_name") or company.get("name") or "Company")
    return company.legal_name


def build_pitch_deck_slide_plan(
    *,
    company: Company | dict[str, Any],
    section_labels: list[str] | None = None,
    current_section: str | None = None,
) -> SlidePlan:
    """Return the canonical 14-slide plan for the slide-library POC."""
    name = _company_name(company)
    labels = section_labels or ["Overview", "Financial Summary", "Valuation", "Process"]
    current = current_section or labels[0]
    slides: list[SlideEntry] = []
    for idx, entry in enumerate(load_slide_library_registry()):
        title = entry.title.replace("[Client Name]", name)
        content_block: dict[str, Any] = {"slide_number": entry.slide_number}
        if entry.static:
            content_block["template_static"] = True
        if entry.library_entry_id == "pitch-cover":
            content_block["requires"] = ["client_name", "presentation_date"]
        elif entry.library_entry_id == "executive-summary":
            content_block["requires"] = ["executive_summary_bullets"]
        elif entry.library_entry_id == "section-divider":
            content_block.update({"section_labels": labels, "current_section": current})
        elif entry.library_entry_id == "public-company-overview":
            content_block["requires"] = ["company_overview_bullets", "cap_table"]
            content_block["deferred"] = ["ltm_revenue_breakdown"]
        elif entry.library_entry_id == "financial-summary":
            content_block["requires"] = ["financial_metric_labels"]
            content_block["deferred"] = ["financial_summary_charts"]
        elif entry.library_entry_id == "acquirer-considerations-mitigants":
            content_block["requires"] = ["risk_mitigants", "risks_tagline"]
        elif entry.library_entry_id == "comparable-companies":
            content_block["requires"] = ["comps_takeaway"]
            content_block["optional"] = ["comps_chart_or_table"]
        elif entry.library_entry_id == "key-investment-highlights":
            content_block["requires"] = ["investment_highlights"]
            content_block["optional"] = ["investment_highlights_tagline"]
        elif entry.library_entry_id == "market-entry-targets":
            content_block["requires"] = ["market_entry_row_labels", "market_entry_targets"]
            content_block["optional"] = ["market_entry_market"]
            content_block["deferred"] = ["target_logos"]
        slides.append(
            SlideEntry(
                library_entry_id=entry.library_entry_id,
                title=title,
                section=_section_for(entry.library_entry_id),
                order=idx,
                content_block=content_block,
            )
        )
    return SlidePlan(
        deliverable_type="pitch",
        deck_title=f"{name} Confidential Discussion Materials",
        slides=slides,
        notes="Phase 3 slide-library POC structure for INFOR Slide Library.pptx.",
    )


def _section_for(entry_id: str) -> str:
    if entry_id in {"pitch-cover", "executive-summary"}:
        return "Executive Summary"
    if entry_id.startswith("infor-"):
        return "INFOR Credentials"
    if entry_id in {"section-divider", "public-company-overview", "financial-summary"}:
        return "Overview"
    if entry_id in {"acquirer-considerations-mitigants", "comparable-companies"}:
        return "Valuation"
    if entry_id in {"key-investment-highlights", "market-entry-targets"}:
        return "Appendix"
    return "Appendix"


def write_slide_plan(plan: SlidePlan, path: Path | str) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(plan.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return target
