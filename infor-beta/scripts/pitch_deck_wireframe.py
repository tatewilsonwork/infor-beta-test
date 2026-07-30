"""Build the SlidePlan for the INFOR slide-library pitch deck."""

from __future__ import annotations

import math
from typing import Any

from schemas import Company, SlideEntry, SlidePlan
from slide_library_registry import load_slide_library_registry
from wireframe_common import company_display_name as _company_name, write_slide_plan


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
        # Metric labels come from the `financial-summary` stage output (not the
        # content bundle); the charts remain deferred placeholders.
        block["external"] = ["financial_metric_labels"]
        block["deferred"] = ["financial_summary_charts"]
    elif entry.library_entry_id == "acquirer-considerations-mitigants":
        block["requires"] = ["risk_mitigants", "risks_tagline"]
    elif entry.library_entry_id == "comparable-companies":
        block["requires"] = ["comps_takeaway"]
        block["optional"] = ["comps_chart_or_table"]
    elif entry.library_entry_id == "precedent-transactions":
        block["requires"] = ["precedents_takeaway"]
        block["optional"] = ["precedents_chart_or_table"]
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
    financial_metric_count: int | None = None,
    include_investment_highlights: bool | None = None,
) -> SlidePlan:
    """Return the canonical pitch plan for the slide-library deck.

    The blank library is 16 slides (incl. the insider-ownership slide, which
    follows Financial Summary). Three deck-spec options adjust the slide mix:

    - The market-entry section expands to ``ceil(market_entry_target_count / 2)``
      slides (two targets per slide). When the count is unspecified — the analyst
      didn't ask for a particular number — it **defaults to 8 targets (4
      market-entry slides)**, the standard pitch layout; the deck-assembler still
      clones to the true count from the content bundle regardless.
    - The Financial Summary section shows four metrics per slide:
      ``financial_metric_count`` (a positive multiple of 4; the deck spec offers
      4 or 8, **default 4**) grows it to ``count / 4`` slides. The deck-assembler
      clones the library's Financial Summary slide to match this plan.
    - ``include_investment_highlights=False`` drops the Key Investment Highlights
      slide entirely (**default: included**); the deck-assembler deletes the
      library slide when the plan omits its entry.
    """
    name = _company_name(company)
    labels = section_labels or ["Overview", "Financial Summary", "Valuation", "Process"]
    current = current_section or labels[0]
    # Default to 8 market-entry targets (4 slides, two per slide) when the analyst
    # doesn't specify a count.
    target_count = market_entry_target_count if market_entry_target_count else 8
    n_market_entry = max(1, math.ceil(target_count / 2))
    # Default to one Financial Summary slide (4 metrics) when unspecified.
    metric_count = financial_metric_count if financial_metric_count else 4
    if metric_count % 4 != 0 or metric_count <= 0:
        raise ValueError(
            f"financial_metric_count must be a positive multiple of 4 (each "
            f"Financial Summary slide shows four metric tiles); got {metric_count}"
        )
    n_financial_summary = metric_count // 4
    include_kih = True if include_investment_highlights is None else bool(include_investment_highlights)

    slides: list[SlideEntry] = []
    order = 0
    for entry in load_slide_library_registry():
        if entry.library_entry_id == "key-investment-highlights" and not include_kih:
            continue
        if entry.library_entry_id == "market-entry-targets":
            repeat = n_market_entry
        elif entry.library_entry_id == "financial-summary":
            repeat = n_financial_summary
        else:
            repeat = 1
        for k in range(repeat):
            title = entry.title.replace("[Client Name]", name)
            block = _content_block(entry, labels=labels, current=current)
            if entry.library_entry_id == "market-entry-targets" and n_market_entry > 1:
                block["market_entry_slide"] = k + 1
                block["market_entry_slide_count"] = n_market_entry
                title = f"{title} ({k + 1} of {n_market_entry})"
            elif entry.library_entry_id == "financial-summary" and n_financial_summary > 1:
                block["financial_summary_slide"] = k + 1
                block["financial_summary_slide_count"] = n_financial_summary
                title = f"{title} ({k + 1} of {n_financial_summary})"
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
        notes="INFOR Slide Library pitch structure; market-entry expands two targets "
        "per slide and Financial Summary four metrics per slide.",
    )


def _section_for(entry_id: str) -> str:
    """Which section of the DECK an entry belongs to, for the SlidePlan.

    This vocabulary — Executive Summary / INFOR Credentials / Overview /
    Valuation / Appendix — is deliberately **not** the divider slide's
    `section_labels` (Overview / Financial Summary / Valuation / Process). They
    are two different things and are kept apart on purpose:

    - These five are structural: they classify every library entry, including the
      front matter and appendix a divider never advertises, and downstream code
      groups by them.
    - The divider's labels are the analyst's agenda for the client. "Process" is
      on it because a pitch promises one; there is no `process` library entry, so
      no slide can map to it. "Financial Summary" is on it because the client
      cares about the financials; the slide itself is structurally part of the
      Overview section.

    So they overlap in three words and agree on none of their jobs. Reconciling
    them to one constant would mean either putting "Process" on a taxonomy no
    slide can satisfy, or dropping it from an agenda the analyst wants it on —
    which is a client-facing decision, not a refactor. The pairing is only ever
    read by a human, so the cost of the divergence is this comment.
    """
    if entry_id in {"pitch-cover", "executive-summary"}:
        return "Executive Summary"
    if entry_id.startswith("infor-"):
        return "INFOR Credentials"
    if entry_id in {"section-divider", "public-company-overview", "financial-summary", "insider-ownership"}:
        return "Overview"
    if entry_id in {"acquirer-considerations-mitigants", "comparable-companies", "precedent-transactions"}:
        return "Valuation"
    if entry_id in {"key-investment-highlights", "market-entry-targets"}:
        return "Appendix"
    return "Appendix"
