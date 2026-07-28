"""Stage handoff — three arguments in, one JSON file out. No environment variables.

Every conductor-dispatched skill needs the same four things: where the plugin is,
where its resolved inputs are, where its outputs must go, and where the deal
directory is. Until Phase E it got them from `$STAGE_INPUTS` / `$STAGE_OUTPUTS` /
`$DEAL_DIR` / `$CLAUDE_PLUGIN_ROOT`, which the sub-agent had to `export` itself as
its first action — because the `Task` tool has no parameter for environment
variables. That made the whole handoff depend on the exports surviving every
later tool call in the sub-agent's session: the most fragile contract in the
system, and one that fails *silently* (an unset `DEAL_DIR` writes the deliverable
to whatever cwd the shell happened to have).

So the paths are arguments now. The dispatch envelope renders the exact command
line, and each snippet reads it back:

    python <your_script.py> "<plugin_root>" "<stage_inputs.json>" "<stage_outputs.json>"

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(sys.argv[1]) / "scripts"))
    from stage_io import stage_io

    io = stage_io()
    company = io.inputs["company"]
    ...
    io.write({"workbook_path": str(path)})

Nothing is inherited and nothing persists, so a fresh shell, a different tool
call, and a retry all behave identically.

`deal_dir` is *derived*, not passed: the run directory always sits at
`<deal_dir>/runs/<run-id>/stages/<stage-id>/`, so walking up for the directory
holding `deal.json` finds it and proves it at the same time. A fourth argument
would only be another thing to get wrong.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

USAGE = 'python <script.py> "<plugin_root>" "<stage_inputs.json>" "<stage_outputs.json>"'

DEAL_CONTEXT_FILENAME = "deal.json"


class StageIOError(RuntimeError):
    """The stage handoff arguments are missing, malformed, or don't point anywhere."""


def deal_dir_for(path: Path | str) -> Path:
    """Return the deal directory containing `path` — the nearest ancestor with `deal.json`.

    Walks up from the stage's inputs.json (`<deal_dir>/runs/<run-id>/stages/<id>/`).
    Finding the file is also the check that we landed in a real deal directory,
    which a fixed "four levels up" would not be.
    """
    start = Path(path).resolve()
    for candidate in (start, *start.parents):
        if candidate.is_dir() and (candidate / DEAL_CONTEXT_FILENAME).is_file():
            return candidate
    raise StageIOError(
        f"no {DEAL_CONTEXT_FILENAME} in any parent of {start} — the stage inputs path "
        f"is not inside a deal directory"
    )


@dataclass(frozen=True)
class StageIO:
    """One stage's resolved handoff: the paths, the inputs, and how to answer."""

    plugin_root: Path
    inputs_path: Path
    outputs_path: Path
    inputs: dict[str, Any]

    @property
    def stage_dir(self) -> Path:
        """`<run_dir>/stages/<stage-id>/` — where intermediate artefacts belong."""
        return self.inputs_path.parent

    @property
    def run_dir(self) -> Path:
        """`<run_dir>` — this run's directory, two levels up from the stage dir.

        Derived like `deal_dir` rather than passed: the layout is fixed
        (`<deal_dir>/runs/<run-id>/stages/<stage-id>/`), so a stage that needs the
        whole run — `deckcheck` consolidating every stage's provenance fragment —
        asks for it here instead of doing path arithmetic in a SKILL.md snippet,
        where it would rot the next time the layout moved.
        """
        return self.stage_dir.parent.parent

    @property
    def deal_dir(self) -> Path:
        """`<deal_dir>` — derived from the inputs path, never passed in."""
        return deal_dir_for(self.inputs_path)

    @property
    def artefacts_dir(self) -> Path:
        """`<deal_dir>/artefacts/`, created if absent. Deliverables go here, NOT cwd."""
        d = self.deal_dir / "artefacts"
        d.mkdir(parents=True, exist_ok=True)
        return d

    def write(self, outputs: Mapping[str, Any]) -> Path:
        """Write the stage's structured handoff. The conductor reads this, not the reply."""
        self.outputs_path.parent.mkdir(parents=True, exist_ok=True)
        self.outputs_path.write_text(
            json.dumps(dict(outputs), indent=2, sort_keys=True, default=str) + "\n",
            encoding="utf-8",
        )
        return self.outputs_path

    def fail(self, reason: str) -> Path:
        """Write `{"error": reason}` and return the path.

        A stage that cannot do its work must still write outputs.json — the
        conductor halts on a missing file with far less to tell the analyst than
        it can with a reason.
        """
        return self.write({"error": reason})


def stage_io(argv: list[str] | None = None) -> StageIO:
    """Parse the three handoff arguments and load the resolved inputs.

    `argv` defaults to `sys.argv`, so a snippet run as
    `python script.py <plugin_root> <inputs.json> <outputs.json>` needs no
    arguments here.
    """
    if argv is None:
        import sys as _sys

        argv = _sys.argv
    if len(argv) < 4:
        raise StageIOError(
            f"expected three handoff arguments, got {max(len(argv) - 1, 0)}. Usage:\n  {USAGE}\n"
            f"The dispatch envelope for your stage prints the exact command line."
        )
    plugin_root, inputs_path, outputs_path = (Path(a).expanduser() for a in argv[1:4])
    if not (plugin_root / "scripts").is_dir():
        raise StageIOError(f"{plugin_root} is not a plugin root (no scripts/ directory)")
    if not inputs_path.is_file():
        raise StageIOError(f"stage inputs not found at {inputs_path}")
    inputs = json.loads(inputs_path.read_text(encoding="utf-8"))
    if not isinstance(inputs, dict):
        raise StageIOError(f"{inputs_path} must hold a JSON object, got {type(inputs).__name__}")
    return StageIO(
        plugin_root=plugin_root.resolve(),
        inputs_path=inputs_path.resolve(),
        outputs_path=outputs_path.resolve(),
        inputs=inputs,
    )
