"""Plan + Stage schemas.

Per Obsidian note 12, H1: start minimal. There is no DAG `depends_on` /
`parallel_with` field and no expression language inside reference strings —
a stage declares its dependencies implicitly, by referencing an earlier
stage's output (`$stages.<id>.<name>`). The conductor derives an execution
schedule from those references and dispatches independent stages in concurrent
waves (see `plan_schedule.compute_waves`). Since Phase D every ordering
constraint is a real reference — there is no hardcoded barrier.

Reference resolution is a pure string-templating pass over `Stage.inputs`:

- `$plan_inputs.<name>`        analyst-supplied plan input
- `$deal.<field>`              fields on DealContext: codename, deal_dir,
                               deliverable_type, subject_company, filings, notes,
                               or dotted access like `subject_company.ticker`
- `$stages.<stage_id>.<name>`  named output of an earlier stage

The resolver itself lives in `plan_refs.py` so this module stays a pure schema.
So does the load-time reference pre-flight (`plan_refs.validate_plan_references`),
which the conductor's load paths run right after pydantic validation: every
`$stages.<id>` must name a real stage, every `$stages.<id>.<name>` a declared
output of that stage, and every `$plan_inputs.<name>` a declared plan input.
It is deliberately not a validator on `Plan` itself — the scheduler keeps its
lenient ignore-unknown-refs behaviour on hand-built plans as defense-in-depth.
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .deal_context import DeliverableType

CheckpointMode = Literal["required", "informational", "silent"]


class InputSpec(BaseModel):
    """One typed input a skill or plan consumes."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1, description="Parameter name as referenced inside the skill.")
    type: str = Field(
        ...,
        min_length=1,
        description="Type label. Free-form — e.g. 'Company', 'list[Filing]', 'str', 'Path'.",
    )
    required: bool = Field(default=True, description="Whether the conductor must supply this input.")
    description: str | None = Field(default=None, description="Human-readable explanation.")


class OutputSpec(BaseModel):
    """One typed output a skill or stage produces."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1, description="Output name as exported by the skill.")
    type: str = Field(..., min_length=1, description="Type label. Free-form.")
    description: str | None = Field(default=None, description="Human-readable explanation.")


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
        description=(
            "Named outputs the stage must write to its outputs.json. Every declared "
            "name is checked for presence when the conductor collects the stage "
            "(null values pass — omitting the key does not; extra undeclared keys "
            "are allowed; the type label is documentation, not validated)."
        ),
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
        description=(
            "Stages in declaration order. The conductor derives a concurrent wave "
            "schedule from their `$stages.*` references (plan_schedule.compute_waves) "
            "rather than running them strictly top-to-bottom; declaration order only "
            "breaks ties within a wave."
        ),
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
