"""Helpers shared by the earnings-update and pitch wireframe stages.

Both wireframes derive a display name from a `Company` (or raw dict) and write a
typed `SlidePlan` JSON artefact the same way; these used to be byte-identical
copies in each module. They are re-exported from each wireframe module so the
existing ``from <wireframe> import write_slide_plan`` import path still works.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from schemas import Company, SlidePlan


def company_display_name(company: Company | dict[str, Any]) -> str:
    """Best-effort display name from a `Company` model or a raw dict."""
    if isinstance(company, dict):
        return str(company.get("legal_name") or company.get("name") or "Company")
    return company.legal_name


def write_slide_plan(plan: SlidePlan, path: Path | str) -> Path:
    """Write a SlidePlan JSON artefact and return its path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(plan.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return target
