"""Build the SlidePlan for the INFOR slide-library pitch deck."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from schemas import Company, SlideEntry, SlidePlan
from slide_library_registry import load_slide_library_registry


def _company_name(company: Company | dict[str, Any]) -> str:
    if isinstance(company, dict):
        return str(company.get("legal_name") or company.get("name") or "Company")
    return company.legal_name


def _content_block(entry, *, labels: list[str], current: str) -> dict[str, Any]:
    block: dict[str, Any] = {"slide_number": entry.slide_number}
    if entry.static:
        block["template_static"] = True
    if entry.library_entry_id == "pitch-cover":
        block["requires"] = ["client_name", "presentation_date"]
    elif entry.library_entry_id == "executive-summary":
        block["requires"] = ["executive_summary_bullets"]
    elif entry.library_entry_id == "section-divider":
        block.update({"section_labels": labels, "current_section": current})
    elif entry.library_entry_id == "public-company-overview":
        block["requires"] = ["company_overview_bullets", "cap_table"]
        block["deferred"] = ["ltm_revenue_breakdown"]
    elif entry.library_entry_id == "financial-summary":
        block["requires"] = ["financial_metric_labels"]
        block["deferred"] = ["financial_summary_charts"]
    elif entry.library_entry_id == "acquirer-considerations-mitigants":
        block["requires"] = ["risk_mitigants", "risks_tagline"]
    elif entry.library_entry_id == "comparable-companies":
        block["requires"] = ["comps_takeaway"]
        block["optional"] = ["comps_chart_or_table"]
    elif entry.library_entry_id == "key-investment-highlights":
        block["requires"] = ["investment_highlights"]
        block["optional"] = ["investment_highlights_tagline"]
    elif entry.library_entry_id == "market-entry-targets":
        block["requires"] = ["market_entry_row_labels", "market_entry_targets"]
        block["optional"] = ["market_entry_market"]
        block["deferred"] = ["target_logos"]
    elif entry.library_entry_id == "insider-ownership":
        # The insider table is a picture of the ownership workbook (built by the
        # `ownership` skill from a SEDI filing); the institutional side stays a
        # Bloomberg-sourced placeholder, like the overview slide's revenue pie.
        block["requires"] = ["ownership_table"]
        block["deferred"] = ["institutional_ownership"]
    return block


def build_pitch_deck_slide_plan(
    *,
    company: Company | dict[str, Any],
    section_labels: list[str] | None = None,
    current_section: str | None = None,
    market_entry_target_count: int | None = None,
) -> SlidePlan:
    """Return the canonical pitch plan for the slide-library deck.

    The blank library is 15 slides (incl. the insider-ownership slide). The
    market-entry section expands to ``ceil(market_entry_target_count / 2)``
    slides (two targets per slide). When
    the count is unknown at wireframe time — the live pipeline drafts targets
    later, in pitch-content — it defaults to a single market-entry slide; the
    deck-assembler clones to the true count from the content bundle regardless.
    """
    name = _company_name(company)
    labels = section_labels or ["Overview", "Financial Summary", "Valuation", "Process"]
    current = current_section or labels[0]
    n_market_entry = (
        max(1, math.ceil(market_entry_target_count / 2)) if market_entry_target_count else 1
    )

    slides: list[SlideEntry] = []
    order = 0
    for entry in load_slide_library_registry():
        repeat = n_market_entry if entry.library_entry_id == "market-entry-targets" else 1
        for k in range(repeat):
            title = entry.title.replace("[Client Name]", name)
            block = _content_block(entry, labels=labels, current=current)
            if entry.library_entry_id == "market-entry-targets" and n_market_entry > 1:
                block["market_entry_slide"] = k + 1
                block["market_entry_slide_count"] = n_market_entry
                title = f"{title} ({k + 1} of {n_market_entry})"
            slides.append(
                SlideEntry(
                    library_entry_id=entry.library_entry_id,
                    title=title,
                    section=_section_for(entry.library_entry_id),
                    order=order,
                    content_block=block,
                )
            )
            order += 1
    return SlidePlan(
        deliverable_type="pitch",
        deck_title=f"{name} Confidential Discussion Materials",
        slides=slides,
        notes="INFOR Slide Library pitch structure; market-entry expands two targets per slide.",
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
    if entry_id in {"key-investment-highlights", "market-entry-targets", "insider-ownership"}:
        return "Appendix"
    return "Appendix"


def write_slide_plan(plan: SlidePlan, path: Path | str) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(plan.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return target
