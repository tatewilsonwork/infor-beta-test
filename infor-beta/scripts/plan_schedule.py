"""Execution-wave scheduler for conductor plans — parallel dispatch.

The v1 conductor ran every stage sequentially in declaration order (Obsidian
note 12, H1: "start minimal"). This module derives the *actual* dependency DAG
from the `$stages.<id>.<name>` references already present in each stage's
inputs and groups the stages into ordered **waves** of mutually-independent
stages. The conductor dispatches each wave concurrently (one Agent/`Task` call
per stage, all issued in a single message), waits for the whole wave to finish,
then moves to the next — collapsing the critical path from the *sum* of stage
durations to the *longest dependency chain*.

Two kinds of edge feed the DAG:

1. **Data edges — auto-derived.** If stage B's inputs contain the reference
   string `$stages.A.<name>`, then B depends on A. This reuses the very same
   reference grammar the resolver (`plan_refs`) understands, so the schedule can
   never disagree with what actually resolves at dispatch time. No `depends_on`
   field is added to the Stage schema — the references *are* the DAG.

2. **The aggregator barrier — hardcoded.** The `workbook-aggregator` stage
   *mutates* the run: it merges every companion `.xlsx` into one combined
   workbook and deletes the individual sources. The deck-assembler reads one of
   those sources (the standalone cap table) before it is folded in, and that
   ordering is a filesystem side-effect, NOT a value reference — so it is
   invisible to the data-edge derivation above. We therefore force the
   aggregator to depend on every other stage EXCEPT its own downstream
   consumers (stages that reference `$stages.workbook-aggregation.*`, e.g. the
   pitch plan's post-aggregation `financial-charts`, which must run after it —
   including them would be a cycle). The aggregator is thus always alone in its
   wave, with any post-aggregation consumers in later waves; see
   `stage_dependencies` (generalized in v0.5.16 from the original strict
   final-barrier form).

The scheduler only *orders* stages. It does not validate references: a typo'd
`$stages.<id>` is ignored here and left for `plan_refs.resolve_refs` to reject
at dispatch time with a clearer message.
"""

from __future__ import annotations

import re
from typing import Any, Iterable

from schemas import Plan

# Skill name whose stage acts as a strict final barrier (see module docstring).
_AGGREGATOR_SKILL = "workbook-aggregator"

# Matches a whole-string stage reference `$stages.<stage_id>.<output_name>`.
# Mirrors the `stages` branch of plan_refs._REF_RE: the stage id is everything
# between `$stages.` and the first dot (stage ids carry no dots — they are
# lowercase + hyphens/underscores by convention). Kept local so this module
# stays import-light, but the grammar must not drift from the resolver's.
_STAGE_REF_RE = re.compile(r"^\$stages\.([^.]+)\.(.+)$")


class PlanCycleError(ValueError):
    """The plan's stage dependencies form a cycle — it cannot be scheduled."""


def _iter_strings(value: Any) -> Iterable[str]:
    """Yield every string anywhere inside a (possibly nested) inputs value.

    Walks dicts and lists/tuples so references buried in a sub-structure — e.g.
    the pitch aggregator's `workbooks: {captable: $stages.captable.workbook_path,
    ...}` mapping — are still found.
    """
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for v in value.values():
            yield from _iter_strings(v)
    elif isinstance(value, (list, tuple)):
        for v in value:
            yield from _iter_strings(v)


def _transitive_deps(deps: dict[str, set[str]]) -> dict[str, set[str]]:
    """Transitive closure of a direct-dependency map (fixpoint iteration)."""
    closure = {k: set(v) for k, v in deps.items()}
    changed = True
    while changed:
        changed = False
        for k in closure:
            add: set[str] = set()
            for d in closure[k]:
                add |= closure.get(d, set())
            if not add <= closure[k]:
                closure[k] |= add
                changed = True
    return closure


def stage_dependencies(plan: Plan) -> dict[str, set[str]]:
    """Map each stage id to the set of stage ids it depends on.

    Data edges are derived from `$stages.<id>.<name>` references found anywhere
    in a stage's inputs. The `workbook-aggregator` stage (if present) is then
    forced to depend on every other stage **except its own downstream
    consumers** — the barrier rule (see module docstring). Self-references and
    references to unknown stage ids are dropped.

    Excluding the aggregator's consumers matters once a stage runs *after*
    aggregation (e.g. `financial-charts`, which charts the combined workbook):
    such a stage references `$stages.<aggregator>.…`, so blindly making the
    aggregator depend on *every* stage would form a cycle. A stage that consumes
    the combined workbook must run after the aggregator, never before it. Plans
    with no post-aggregation consumer are unaffected (the consumer set is empty,
    so the aggregator still depends on every other stage and is alone last).
    """
    ids = {s.id for s in plan.stages}
    deps: dict[str, set[str]] = {s.id: set() for s in plan.stages}

    for stage in plan.stages:
        for s in _iter_strings(stage.inputs):
            m = _STAGE_REF_RE.match(s)
            if not m:
                continue
            dep_id = m.group(1)
            if dep_id in ids and dep_id != stage.id:
                deps[stage.id].add(dep_id)

    # Aggregator barrier: it consolidates + deletes every source workbook, so it
    # must run strictly last among its *producers* — after the deck reads the
    # standalone cap table — but strictly *before* any stage that consumes the
    # combined workbook it emits. Use the data-edge closure to find those
    # consumers and leave them out of the forced dependency set.
    trans = _transitive_deps(deps)
    for stage in plan.stages:
        if stage.skill == _AGGREGATOR_SKILL:
            consumers = {sid for sid in ids if stage.id in trans.get(sid, set())}
            deps[stage.id] |= ids - {stage.id} - consumers

    return deps


def compute_waves(plan: Plan) -> list[list[str]]:
    """Topologically sort the plan into ordered waves of independent stage ids.

    Wave *i* contains every stage whose dependencies are all satisfied by waves
    0..i-1. Stages within a wave have no dependency between them and are safe to
    dispatch concurrently. Within a wave, ids are returned in plan declaration
    order so the schedule is deterministic and reads predictably in logs.

    Raises ``PlanCycleError`` if the dependencies form a cycle (no stage ever
    becomes ready while stages remain).
    """
    deps = stage_dependencies(plan)
    order = [s.id for s in plan.stages]  # declaration order, for stable output
    done: set[str] = set()
    waves: list[list[str]] = []

    while len(done) < len(order):
        ready = [sid for sid in order if sid not in done and deps[sid] <= done]
        if not ready:
            remaining = [sid for sid in order if sid not in done]
            raise PlanCycleError(
                f"cannot schedule stages {remaining!r}: dependency cycle "
                f"(or a stage depends on one that never becomes ready)"
            )
        waves.append(ready)
        done.update(ready)

    return waves
