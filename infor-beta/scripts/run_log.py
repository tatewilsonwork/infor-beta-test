"""Run-log helpers — Phase 2.

The conductor writes a per-run log under `<deal_dir>/runs/<run-id>/` per
locked decision H4. Layout:

    <deal_dir>/runs/2026-05-15-earnings-update-7f3a/
    ├── plan.yaml          (frozen snapshot of the plan that ran)
    ├── stages/
    │   └── <stage-id>/
    │       ├── inputs.json
    │       ├── outputs.json
    │       └── log.txt    (sub-agent transcript)
    └── summary.md         (analyst-readable end-of-run summary)

This module is filesystem-only — no Agent dispatch, no pydantic validation.
Validation is the conductor's / schema layer's job. We just create the
directory tree and serialise dicts predictably.
"""

from __future__ import annotations

import json
import uuid
from datetime import date
from pathlib import Path
from typing import Any, Mapping


def make_run_id(plan_id: str, *, today: date | None = None) -> str:
    """Compose a unique run id of the form `YYYY-MM-DD-<plan-id>-<short>`.

    `plan_id` is sanitised to a filesystem-safe form: lowercase, runs of
    non-alphanumerics collapsed to a single hyphen, leading/trailing
    hyphens stripped. Matches the convention used by `sanitize_name.sh`.

    `short` is the first 8 hex chars of a fresh uuid4 — short enough to read,
    long enough for collision-free per-deal-per-day runs.
    """
    import re

    if today is None:
        today = date.today()
    # Lowercase first, then collapse runs of non-alnum to a single hyphen,
    # then strip leading/trailing hyphens.
    safe = re.sub(r"[^a-z0-9]+", "-", plan_id.lower()).strip("-") or "run"
    short = uuid.uuid4().hex[:8]
    return f"{today.isoformat()}-{safe}-{short}"


def create_run_dir(deal_dir: Path | str, run_id: str) -> Path:
    """Create `<deal_dir>/runs/<run-id>/` and its `stages/` subdir. Idempotent.

    Returns the absolute path to the run directory.
    """
    root = Path(deal_dir).expanduser()
    run_dir = root / "runs" / run_id
    (run_dir / "stages").mkdir(parents=True, exist_ok=True)
    return run_dir


def write_plan_snapshot(run_dir: Path, plan_yaml_text: str) -> Path:
    """Persist the plan YAML verbatim under `<run_dir>/plan.yaml`."""
    target = Path(run_dir) / "plan.yaml"
    target.write_text(plan_yaml_text, encoding="utf-8")
    return target


def stage_dir(run_dir: Path | str, stage_id: str) -> Path:
    """Return `<run_dir>/stages/<stage_id>/`, creating it if missing."""
    d = Path(run_dir) / "stages" / stage_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _json_default(o: Any) -> Any:
    """Best-effort JSON serialiser for Path / pydantic / sets."""
    if isinstance(o, Path):
        return str(o)
    if isinstance(o, set):
        return sorted(o)
    if hasattr(o, "model_dump"):
        # pydantic v2 BaseModel
        return o.model_dump(mode="json")
    raise TypeError(f"object of type {type(o).__name__} is not JSON serialisable")


def write_stage_inputs(run_dir: Path | str, stage_id: str, inputs: Mapping[str, Any]) -> Path:
    """Write the stage's resolved inputs as JSON."""
    target = stage_dir(run_dir, stage_id) / "inputs.json"
    target.write_text(
        json.dumps(dict(inputs), indent=2, sort_keys=True, default=_json_default) + "\n",
        encoding="utf-8",
    )
    return target


def read_stage_outputs(run_dir: Path | str, stage_id: str) -> dict[str, Any]:
    """Read a stage's outputs.json. Raises FileNotFoundError if the file is missing
    (so the conductor refuses to proceed past a stage that didn't write outputs)."""
    path = stage_dir(run_dir, stage_id) / "outputs.json"
    if not path.exists():
        raise FileNotFoundError(
            f"stage {stage_id!r} did not write outputs.json at {path} — "
            f"the conductor cannot proceed."
        )
    return json.loads(path.read_text(encoding="utf-8"))


def write_stage_log(run_dir: Path | str, stage_id: str, transcript: str) -> Path:
    """Write the sub-agent transcript for this stage."""
    target = stage_dir(run_dir, stage_id) / "log.txt"
    target.write_text(transcript, encoding="utf-8")
    return target


def write_summary(run_dir: Path | str, summary_md: str) -> Path:
    """Write the end-of-run summary as `<run_dir>/summary.md`."""
    target = Path(run_dir) / "summary.md"
    target.write_text(summary_md, encoding="utf-8")
    return target
