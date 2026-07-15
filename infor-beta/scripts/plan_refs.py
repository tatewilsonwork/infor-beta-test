"""Reference resolver for Plan/Stage inputs — Phase 2.

The conductor's plan YAML lets stages reference values from three sources
using simple `$prefix.path` strings (Obsidian note 12, H1):

  - `$plan_inputs.<name>`       analyst-supplied plan inputs
  - `$deal.<field>` (dotted)    fields on the DealContext, including nested
                                attribute access like `$deal.subject_company.ticker`
  - `$stages.<stage_id>.<name>` a named output of an earlier stage

This module is intentionally a tiny pure-Python helper — no pydantic
import, no I/O, no logging. `resolve_refs(value, ...)` walks dicts, lists,
tuples, and strings and replaces refs in place. Non-string scalars
(ints, bools, paths, pydantic models) pass through unchanged.

A reference string is detected by a leading `$`. Mid-string interpolation
(e.g. `"hello $deal.codename"`) is NOT supported — the v1 design uses
whole-string refs only. This is intentional: it avoids the fragility of a
mini-templating language, and matches how the conductor uses these (every
stage input is either a literal or one named reference).

Unknown prefixes raise `UnknownReferenceError`. Missing fields raise
`ReferenceResolutionError`. Both are subclasses of `ValueError` so callers
can catch the broad class.

This module is also the single owner of the reference **grammar**: `parse_ref`
splits a candidate string into `(prefix, path_parts)` and `iter_input_strings`
walks a stage's inputs for candidate strings. The scheduler
(`plan_schedule.stage_dependencies`) derives the execution DAG through these
same two functions — the refs ARE the DAG, and there is deliberately no second
regex anywhere that could drift from this one.

`validate_plan_references(plan)` is the load-time pre-flight over that grammar:
it checks every `$stages.<id>` names a real stage, every `$stages.<id>.<name>`
names a declared output of that stage, and every `$plan_inputs.<name>` names a
declared plan input — so a typo'd plan is rejected when it is loaded, not
mid-run when the reference fails to resolve. It is a standalone helper (invoked
on the conductor's load paths) rather than a pydantic validator on `Plan`, so
the scheduler keeps its documented leniency toward unknown-stage refs as
defense-in-depth on hand-built plans.

Optional plan inputs: a plan may declare `InputSpec`s with `required=False`
(e.g. `section_labels`, `valuation_range`, `risk_notes`). When the analyst
doesn't supply one, the conductor's `plan_inputs` dict simply won't carry that
key — and a stage that references it via `$plan_inputs.<name>` would otherwise
fail resolution. Pass the set of optional plan-input names as
`optional_plan_inputs`; a missing `$plan_inputs.<name>` whose name is in that
set resolves to `None` instead of raising. Missing *required* plan inputs,
unknown plan inputs, and any missing `$deal.*` / `$stages.*` reference still
raise — only the explicitly-declared-optional ones are softened.
"""

from __future__ import annotations

import re
from typing import AbstractSet, Any, Iterable, Mapping

_REF_RE = re.compile(r"^\$(plan_inputs|deal|stages)\.(.+)$")


class ReferenceResolutionError(ValueError):
    """A `$prefix.path` reference could not be resolved."""


class UnknownReferenceError(ReferenceResolutionError):
    """The reference uses a prefix this resolver doesn't know about."""


class PlanReferenceError(ValueError):
    """A plan's stage inputs carry a reference that can never resolve
    (unknown stage id, undeclared stage output, or unknown plan input)."""


def parse_ref(value: str) -> tuple[str, list[str]] | None:
    """Parse a whole-string reference into `(prefix, path_parts)`.

    Returns `("stages", ["captable", "workbook_path"])` for
    `"$stages.captable.workbook_path"`, and `None` when the string is not a
    reference in the grammar (`$plan_inputs.<name>` / `$deal.<field>` /
    `$stages.<id>.<name>`). This is the ONLY parser of the reference grammar —
    the resolver, the wave scheduler, and the load-time validator all go
    through it, so they cannot disagree about what counts as a reference.
    """
    m = _REF_RE.match(value)
    if not m:
        return None
    return m.group(1), m.group(2).split(".")


def iter_input_strings(value: Any) -> Iterable[str]:
    """Yield every string anywhere inside a (possibly nested) inputs value.

    Walks dicts and lists/tuples so references buried in a sub-structure — e.g.
    the pitch aggregator's `workbooks: {captable: $stages.captable.workbook_path,
    ...}` mapping — are still found. Shared by the wave scheduler and the
    load-time validator.
    """
    if isinstance(value, str):
        yield value
    elif isinstance(value, Mapping):
        for v in value.values():
            yield from iter_input_strings(v)
    elif isinstance(value, (list, tuple)):
        for v in value:
            yield from iter_input_strings(v)


def _walk_attribute_path(root: Any, parts: list[str], full_ref: str) -> Any:
    """Walk dotted attribute / key path on `root`, raising on miss."""
    current = root
    for i, part in enumerate(parts):
        if current is None:
            raise ReferenceResolutionError(
                f"reference {full_ref!r} bottomed out at None before reaching {'.'.join(parts[i:])!r}"
            )
        if isinstance(current, Mapping):
            if part not in current:
                raise ReferenceResolutionError(
                    f"reference {full_ref!r}: key {part!r} not in mapping (have: {sorted(current.keys())})"
                )
            current = current[part]
        else:
            if not hasattr(current, part):
                raise ReferenceResolutionError(
                    f"reference {full_ref!r}: attribute {part!r} not on {type(current).__name__}"
                )
            current = getattr(current, part)
    return current


def _resolve_one(
    ref: str,
    *,
    plan_inputs: Mapping[str, Any],
    deal_context: Any,
    stage_outputs: Mapping[str, Mapping[str, Any]],
    optional_plan_inputs: AbstractSet[str],
) -> Any:
    parsed = parse_ref(ref)
    if parsed is None:
        raise UnknownReferenceError(
            f"reference {ref!r} does not match `$plan_inputs.<name>`, `$deal.<field>`, or `$stages.<id>.<name>`"
        )
    prefix, parts = parsed

    if prefix == "plan_inputs":
        # A declared-optional plan input the analyst didn't supply resolves to
        # None rather than raising. Only the TOP-level name is softened — a
        # supplied-but-malformed optional input (e.g. dotted access past a None)
        # still surfaces through the normal walk below.
        top = parts[0]
        if top not in plan_inputs and top in optional_plan_inputs:
            return None
        return _walk_attribute_path(plan_inputs, parts, ref)
    if prefix == "deal":
        return _walk_attribute_path(deal_context, parts, ref)
    if prefix == "stages":
        if len(parts) < 2:
            raise ReferenceResolutionError(
                f"reference {ref!r} must look like `$stages.<stage_id>.<output_name>` (got {parts!r})"
            )
        stage_id, *output_parts = parts
        if stage_id not in stage_outputs:
            raise ReferenceResolutionError(
                f"reference {ref!r}: stage {stage_id!r} has not run (or produced no outputs). "
                f"Available: {sorted(stage_outputs.keys())}"
            )
        return _walk_attribute_path(stage_outputs[stage_id], output_parts, ref)

    # Unreachable — _REF_RE constrains the prefix — but kept for safety.
    raise UnknownReferenceError(f"reference {ref!r}: unknown prefix {prefix!r}")


def resolve_refs(
    value: Any,
    *,
    plan_inputs: Mapping[str, Any],
    deal_context: Any,
    stage_outputs: Mapping[str, Mapping[str, Any]],
    optional_plan_inputs: AbstractSet[str] | None = None,
) -> Any:
    """Recursively walk `value` and resolve any `$ref` strings within it.

    - dicts and lists are walked in-place (a new container of the same type is returned).
    - strings starting with `$` and matching one of the three prefixes are replaced.
    - any other string is returned verbatim.
    - non-string scalars (int, bool, float, None, Path, pydantic model) pass through.

    `plan_inputs` should be a plain dict[name -> resolved value].
    `deal_context` may be either a pydantic DealContext or a dict-shaped equivalent.
    `stage_outputs` is `{stage_id: {output_name: value}}` from prior stages.
    `optional_plan_inputs` is the set of plan-input names declared `required=False`
    (typically `{spec.name for spec in plan.plan_inputs if not spec.required}`).
    A `$plan_inputs.<name>` whose name is in this set but absent from `plan_inputs`
    resolves to `None` instead of raising; defaults to the empty set, so existing
    callers keep the strict behaviour. Missing required/unknown plan inputs and any
    missing `$deal.*` / `$stages.*` reference still raise.
    """
    optional = optional_plan_inputs if optional_plan_inputs is not None else frozenset()
    if isinstance(value, str):
        if value.startswith("$"):
            return _resolve_one(
                value,
                plan_inputs=plan_inputs,
                deal_context=deal_context,
                stage_outputs=stage_outputs,
                optional_plan_inputs=optional,
            )
        return value
    if isinstance(value, Mapping):
        return {
            k: resolve_refs(
                v,
                plan_inputs=plan_inputs,
                deal_context=deal_context,
                stage_outputs=stage_outputs,
                optional_plan_inputs=optional,
            )
            for k, v in value.items()
        }
    if isinstance(value, (list, tuple)):
        resolved = [
            resolve_refs(
                v,
                plan_inputs=plan_inputs,
                deal_context=deal_context,
                stage_outputs=stage_outputs,
                optional_plan_inputs=optional,
            )
            for v in value
        ]
        return type(value)(resolved) if isinstance(value, tuple) else resolved
    return value


def validate_plan_references(plan: Any) -> None:
    """Load-time pre-flight over every reference string in a plan's stage inputs.

    Rejects, with `PlanReferenceError` listing every problem at once:

    - `$stages.<id>.<name>` where `<id>` is not a stage in the plan, or `<name>`
      is not one of that stage's declared outputs (only the first path segment
      is checked — dotted access *into* an output value is resolved at runtime);
    - `$plan_inputs.<name>` where `<name>` is not a declared plan input
      (required or optional — both are declared);
    - any whole string starting with `$` that doesn't parse as a reference at
      all (it would raise `UnknownReferenceError` at dispatch time anyway).

    `$deal.<field>` references are left to resolve time — they are checked
    against the live DealContext, which doesn't exist at plan load.

    `plan` is duck-typed (needs `.stages` / `.plan_inputs` shaped like the
    pydantic `Plan`) so this module keeps its no-pydantic-import property.
    Deliberately NOT a pydantic validator on `Plan`: the wave scheduler keeps
    its lenient ignore-unknown-stage-refs behaviour on hand-built plans as
    defense-in-depth, and callers opt into strictness on the load paths
    (`conductor_cli.load_plan` and the conductor's Step 3).
    """
    stage_ids = {s.id for s in plan.stages}
    declared_outputs = {s.id: {o.name for o in s.outputs} for s in plan.stages}
    plan_input_names = {spec.name for spec in plan.plan_inputs}

    problems: list[str] = []
    for stage in plan.stages:
        for candidate in iter_input_strings(stage.inputs):
            if not candidate.startswith("$"):
                continue
            parsed = parse_ref(candidate)
            if parsed is None:
                problems.append(
                    f"stage {stage.id!r}: {candidate!r} does not match `$plan_inputs.<name>`, "
                    f"`$deal.<field>`, or `$stages.<id>.<name>`"
                )
                continue
            prefix, parts = parsed
            if prefix == "plan_inputs":
                if parts[0] not in plan_input_names:
                    problems.append(
                        f"stage {stage.id!r}: {candidate!r} references unknown plan input "
                        f"{parts[0]!r} (declared: {sorted(plan_input_names)})"
                    )
            elif prefix == "stages":
                if len(parts) < 2:
                    problems.append(
                        f"stage {stage.id!r}: {candidate!r} must look like "
                        f"`$stages.<stage_id>.<output_name>`"
                    )
                elif parts[0] not in stage_ids:
                    problems.append(
                        f"stage {stage.id!r}: {candidate!r} references unknown stage "
                        f"{parts[0]!r} (stages: {sorted(stage_ids)})"
                    )
                elif parts[1] not in declared_outputs[parts[0]]:
                    problems.append(
                        f"stage {stage.id!r}: {candidate!r} references undeclared output "
                        f"{parts[1]!r} of stage {parts[0]!r} "
                        f"(declared: {sorted(declared_outputs[parts[0]])})"
                    )

    if problems:
        raise PlanReferenceError(
            "plan failed reference pre-flight:\n" + "\n".join(f"- {p}" for p in problems)
        )
