"""Build the fixed earnings-update SlidePlan for the Phase 3 POC."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from schemas import Company, SlideEntry, SlidePlan


def _company_name(company: Company | dict[str, Any]) -> str:
    if isinstance(company, dict):
        return str(company.get("legal_name") or company.get("name") or "Company")
    return company.legal_name


def build_earnings_update_slide_plan(
    *,
    company: Company | dict[str, Any],
    reporting_quarter: str,
    comparison_quarter: str,
) -> SlidePlan:
    """Return the fixed five-slide plan for the earnings-update POC.

    The stage plans structure only. It deliberately does not draft copy, select
    KPIs, parse EEO data, or touch PowerPoint.
    """
    name = _company_name(company)
    slides = [
        SlideEntry(
            library_entry_id="earnings-update-cover",
            title="Cover",
            section="Cover",
            order=0,
            content_block={"requires": ["cover_date"]},
        ),
        SlideEntry(
            library_entry_id="earnings-update-company-overview",
            title=f"Introduction to {name}",
            section="Company Overview",
            order=1,
            content_block={
                "requires": ["company_overview_bullets", "currency"],
                "cap_table_placeholder": "Leave Rectangle 3 untouched for analyst-pasted cap table (Capitalization Summary).",
                "ltm_revenue_pie_placeholder": "Leave the lower-left pie placeholder untouched; the companion LTM revenue workbook is built by the ltm-revenue stage.",
            },
        ),
        SlideEntry(
            library_entry_id="earnings-update-earnings-summary",
            title=f"{name} {reporting_quarter} Earnings Summary",
            section="Earnings Summary",
            order=2,
            content_block={
                "reporting_quarter": reporting_quarter,
                "comparison_quarter": comparison_quarter,
                "requires": [
                    "kpi_rows",
                    "business_updates",
                    "broker_rows",
                    "management_quotes",
                    "performance_summary",
                ],
            },
        ),
        SlideEntry(
            library_entry_id="earnings-update-disclaimer",
            title="Disclaimer",
            section="Appendix",
            order=3,
            content_block={"template_static": True},
        ),
        SlideEntry(
            library_entry_id="earnings-update-contact",
            title="Contact",
            section="Appendix",
            order=4,
            content_block={"template_static": True},
        ),
    ]
    return SlidePlan(
        deliverable_type="earnings-update",
        deck_title=f"{name} Earnings Update",
        slides=slides,
        notes="Phase 3 POC fixed structure cloned from INFOR Slide Library.pptx (library slides 1, 7, 8, 14, 15).",
    )


def write_slide_plan(plan: SlidePlan, path: Path | str) -> Path:
    """Write a SlidePlan JSON artefact and return its path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(plan.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return target
