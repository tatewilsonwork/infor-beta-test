"""SkillManifest schema — Phase 1.

Per Obsidian note 12, G8: minimum-viable manifest with seven required fields.
Extension fields (`cost_class`, `expected_duration`, `gate_mode_default`,
`idempotent`) are deferred to Phase 2 when the conductor needs them.

`version` must be a SemVer-ish `x.y.z` string and, by single-version policy
(E3), must equal the plugin version at install time. The latter check is a
runtime concern (Phase 2 conductor), not a schema-level validator.
"""

from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field, field_validator

_SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+$")


class InputSpec(BaseModel):
    """One typed input a skill consumes."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1, description="Parameter name as referenced inside the skill.")
    type: str = Field(
        ...,
        min_length=1,
        description="Type label. Free-form in Phase 1 — e.g. 'Company', 'list[Filing]', 'str', 'Path'.",
    )
    required: bool = Field(default=True, description="Whether the conductor must supply this input.")
    description: str | None = Field(default=None, description="Human-readable explanation.")


class OutputSpec(BaseModel):
    """One typed output a skill produces."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(..., min_length=1, description="Output name as exported by the skill.")
    type: str = Field(..., min_length=1, description="Type label. Free-form in Phase 1.")
    description: str | None = Field(default=None, description="Human-readable explanation.")


class SideEffectSpec(BaseModel):
    """One observable side effect (file written, external call, etc.)."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    kind: str = Field(
        ...,
        min_length=1,
        description="Side-effect class, e.g. 'file_write', 'web_search', 'shell'.",
    )
    target: str | None = Field(
        default=None,
        description="Specific target where meaningful (e.g. relative path, hostname).",
    )
    description: str | None = Field(default=None, description="Human-readable explanation.")


class SkillManifest(BaseModel):
    """The seven-field Phase 1 manifest a skill declares to the conductor."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(
        ...,
        min_length=1,
        description="Skill name. Must match the skill directory name.",
    )
    description: str = Field(
        ...,
        min_length=1,
        description="One-paragraph description; lead with the trigger phrases.",
    )
    version: str = Field(
        ...,
        description="SemVer x.y.z. Single-version policy (E3): must equal plugin version at install.",
    )
    inputs: list[InputSpec] = Field(default_factory=list, description="Typed inputs the skill consumes.")
    outputs: list[OutputSpec] = Field(default_factory=list, description="Typed outputs the skill produces.")
    side_effects: list[SideEffectSpec] = Field(
        default_factory=list,
        description="Observable side effects (file writes, web calls, shell, etc.).",
    )
    allowed_tools: list[str] = Field(
        default_factory=list,
        description="Claude Code allowed-tools allowlist (Read, Bash, Write, Glob, WebSearch, ...).",
    )

    @field_validator("version")
    @classmethod
    def _validate_semver(cls, v: str) -> str:
        if not _SEMVER_RE.match(v):
            raise ValueError(f"version must match x.y.z (got {v!r})")
        return v
