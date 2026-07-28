"""Execution-wave scheduler for conductor plans — parallel dispatch.

The v1 conductor ran every stage sequentially in declaration order (Obsidian
note 12, H1: "start minimal"). This module derives the *actual* dependency DAG
from the `$stages.<id>.<name>` references already present in each stage's
inputs and groups the stages into ordered **waves** of mutually-independent
stages. The conductor dispatches each wave concurrently (one Agent/`Task` call
per stage, all issued in a single message), waits for the whole wave to finish,
then moves to the next — collapsing the critical path from the *sum* of stage
durations to the *longest dependency chain*.

**One kind of edge feeds the DAG: data edges, auto-derived.** If stage B's
inputs contain the reference string `$stages.A.<name>`, then B depends on A.
This reuses the very same reference grammar the resolver (`plan_refs`)
understands, so the schedule can never disagree with what actually resolves at
dispatch time. No `depends_on` field is added to the Stage schema — the
references *are* the DAG.

Phase D removed the second kind. A hardcoded **aggregator barrier** used to force
the `workbook-aggregator` stage to depend on every other stage except its own
downstream consumers, because that stage merged every companion `.xlsx` into one
combined workbook and deleted the sources — an ordering the deck-assembler
depended on through the *filesystem*, invisible to a reference scan. The deal now
owns one workbook from stage one, so there is no merge, no deletion, and no
side-effect edge: every ordering constraint in every shipped plan is a real data
reference again.

The scheduler only *orders* stages. It does not validate references: a typo'd
`$stages.<id>` is ignored here and left for the load-time pre-flight
(`plan_refs.validate_plan_references`, run on every conductor load path) — or,
on a hand-built plan that skipped it, for `plan_refs.resolve_refs` at dispatch
time — to reject with a clearer message. The grammar itself is imported from
`plan_refs` (`parse_ref` / `iter_input_strings`), so the edges derived here can
never disagree with what the resolver and the validator understand — the refs
ARE the DAG.
"""

from __future__ import annotations

from plan_refs import iter_input_strings, parse_ref
from schemas import Plan


class PlanCycleError(ValueError):
    """The plan's stage dependencies form a cycle — it cannot be scheduled."""


def stage_dependencies(plan: Plan) -> dict[str, set[str]]:
    """Map each stage id to the set of stage ids it depends on.

    Edges are derived from `$stages.<id>.<name>` references found anywhere in a
    stage's inputs. Self-references and references to unknown stage ids are
    dropped. That is the whole rule: since Phase D deleted the aggregator there
    is no forced barrier, so every ordering constraint is a real data reference
    and this function is a pure read of the plan.
    """
    ids = {s.id for s in plan.stages}
    deps: dict[str, set[str]] = {s.id: set() for s in plan.stages}

    for stage in plan.stages:
        for s in iter_input_strings(stage.inputs):
            parsed = parse_ref(s)
            if parsed is None:
                continue
            prefix, parts = parsed
            if prefix != "stages" or len(parts) < 2:
                continue
            dep_id = parts[0]
            if dep_id in ids and dep_id != stage.id:
                deps[stage.id].add(dep_id)

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
