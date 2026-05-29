"""DealContext schema — Phase 1.

The DealContext is the conductor's runtime handle to a deal. It carries the
codename in its display form (per G3), the deal directory path, the (possibly
partial) subject Company, any persisted filings, and the deliverable type the
analyst answered in the deal-init prompt (G7).
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from .company import Company
from .filing import Filing

# Deliverable choices as listed in the deal-init prompt (G7). Scoped to the
# deliverables with a plan today (earnings-update, pitch) plus overview (stub,
# built out later); one-off-skill is the direct-invocation escape hatch.
DeliverableType = Literal[
    "pitch",
    "earnings-update",
    "overview",
    "one-off-skill",
]


class DealContext(BaseModel):
    """Runtime context for a single deal. Created at deal-init by the conductor."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True, arbitrary_types_allowed=True)

    codename: str = Field(
        ...,
        min_length=1,
        description="Display-form codename, e.g. 'Project OpenText'. Case is preserved (G3).",
    )
    deal_dir: Path = Field(
        ...,
        description="Path to the deal directory (typically under ~/Documents/INFOR Deals/). Relative or absolute.",
    )
    deliverable_type: DeliverableType = Field(
        ...,
        description="Deliverable the analyst is producing for this deal. See G7 prompt.",
    )
    subject_company: Company | None = Field(
        default=None,
        description="Subject company, populated once deal-init facts are confirmed.",
    )
    filings: list[Filing] = Field(
        default_factory=list,
        description="Filings persisted into the deal directory.",
    )
    notes: str | None = Field(
        default=None,
        description="Optional analyst notes captured at deal-init (G7 item 7).",
    )
