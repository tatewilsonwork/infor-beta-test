"""Plan + Stage schemas — Phase 2.

Per Obsidian note 12, H1: start minimal. v1 conductor executes stages
**sequentially in declaration order**. No parallel execution, no DAG
`depends_on` / `parallel_with`, no expression language inside reference strings.

Reference resolution is a pure string-templating pass over `Stage.inputs`:

- `$plan_inputs.<name>`        analyst-supplied plan input
- `$deal.<field>`              fields on DealContext: codename, deal_dir,
                               deliverable_type, subject_company, filings, notes,
                               or dotted access like `subject_company.ticker`
- `$stages.<stage_id>.<name>`  named output of an earlier stage

The resolver itself lives in `plan_refs.py` so this module stays a pure schema.
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .deal_context import DeliverableType
from .skill_manifest import InputSpec, OutputSpec

CheckpointMode = Literal["required", "informational", "silent"]


class Stage(BaseModel):
    """One stage of a plan. Dispatched to a single skill via the `Agent` tool."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    id: str = Field(
        ...,
        min_length=1,
        description="Stage id, unique within the plan. Lowercase + underscores by convention.",
    )
    skill: str = Field(
        ...,
        min_length=1,
        description="Name of the skill to dispatch (must match a skill directory name).",
    )
    inputs: dict[str, Any] = Field(
        default_factory=dict,
        description=(
            "Inputs to pass to the skill. Values may be literals or reference strings "
            "(`$plan_inputs.x`, `$deal.x`, `$stages.<id>.<name>`); strings are resolved "
            "by `plan_refs.resolve_refs` at dispatch time."
        ),
    )
    outputs: list[OutputSpec] = Field(
        default_factory=list,
        description="Named outputs the stage is expected to write to its outputs.json.",
    )
    checkpoint: CheckpointMode = Field(
        default="informational",
        description=(
            "Checkpoint mode per H2 / A2: `required` halts and awaits explicit approval; "
            "`informational` surfaces a summary and proceeds; `silent` is reserved for "
            "autonomous mode later."
        ),
    )


class Plan(BaseModel):
    """A conductor plan. One file per deliverable under `infor-beta/plans/`."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    deliverable_type: DeliverableType = Field(
        ...,
        description="Which deliverable this plan produces. Must match a value in DeliverableType.",
    )
    description: str = Field(
        ...,
        min_length=1,
        description="One-line description shown in the conductor's plan summary.",
    )
    plan_inputs: list[InputSpec] = Field(
        default_factory=list,
        description=(
            "Plan-specific inputs the conductor collects after deal-init "
            "(e.g. reporting quarter, EEO snip path). Owned by the plan, NOT deal-init."
        ),
    )
    stages: list[Stage] = Field(
        ...,
        min_length=1,
        description="Stages, in execution order. Phase 2 executes them sequentially.",
    )

    @field_validator("stages")
    @classmethod
    def _unique_stage_ids(cls, stages: list[Stage]) -> list[Stage]:
        seen: set[str] = set()
        for s in stages:
            if s.id in seen:
                raise ValueError(f"duplicate stage id {s.id!r} — stage ids must be unique within a plan")
            seen.add(s.id)
        return stages
