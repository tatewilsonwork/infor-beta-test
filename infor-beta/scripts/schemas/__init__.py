"""Typed I/O contract for the infor-beta plugin (Phase 1).

All cross-skill data passes through these pydantic v2 models. Skills that compose
each other consume and produce instances of these types — no free-form prompt glue.

Locked by Obsidian note `12 — Locked Decisions`, G-series.
"""

from .company import Company
from .filing import Filing, FilingType
from .slide_plan import SlidePlan, SlideEntry
from .deal_context import DealContext, DeliverableType
from .plan import Plan, Stage, CheckpointMode
from .skill_manifest import (
    SkillManifest,
    InputSpec,
    OutputSpec,
    SideEffectSpec,
)

__all__ = [
    "Company",
    "Filing",
    "FilingType",
    "SlidePlan",
    "SlideEntry",
    "DealContext",
    "DeliverableType",
    "Plan",
    "Stage",
    "CheckpointMode",
    "SkillManifest",
    "InputSpec",
    "OutputSpec",
    "SideEffectSpec",
]
