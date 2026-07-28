"""Unit tests for the argv-based stage handoff (scripts/stage_io.py).

Plus the drift lock that keeps it that way: no SKILL.md may go back to telling a
sub-agent to `export` its handoff paths, or to read them from `os.environ`. That
contract only held while every one of the sub-agent's later tool calls shared one
shell session, and it failed silently — an unset `DEAL_DIR` wrote the client
deliverable to whatever cwd the shell happened to have.
"""

import json
import re
from pathlib import Path

import pytest

from stage_io import StageIOError, deal_dir_for, stage_io

_PLUGIN_ROOT = Path(__file__).resolve().parents[2]  # infor-beta/
_SKILLS_DIR = _PLUGIN_ROOT / "skills"


@pytest.fixture
def handoff(tmp_path: Path):
    """A realistic run layout: <deal_dir>/runs/<run-id>/stages/<stage-id>/."""
    deal_dir = tmp_path / "Project Test"
    sdir = deal_dir / "runs" / "2026-07-28-pitch-abc12345" / "stages" / "comps"
    sdir.mkdir(parents=True)
    (deal_dir / "deal.json").write_text(json.dumps({"codename": "Project Test"}), encoding="utf-8")
    inputs = sdir / "inputs.json"
    inputs.write_text(json.dumps({"ticker": "OTEX", "deal_workbook": "/wb.xlsx"}), encoding="utf-8")
    return deal_dir, inputs, sdir / "outputs.json"


def _argv(inputs: Path, outputs: Path) -> list[str]:
    return ["script.py", str(_PLUGIN_ROOT), str(inputs), str(outputs)]


# ─── Parsing ─────────────────────────────────────────────────────────────────


def test_reads_the_three_paths_and_loads_the_inputs(handoff):
    _, inputs, outputs = handoff
    io = stage_io(_argv(inputs, outputs))
    assert io.plugin_root == _PLUGIN_ROOT
    assert io.inputs["ticker"] == "OTEX"
    assert io.inputs_path == inputs.resolve()
    assert io.outputs_path == outputs.resolve()
    assert io.stage_dir == inputs.parent.resolve()


def test_defaults_to_sys_argv(handoff, monkeypatch):
    """A snippet run as `python script.py <root> <in> <out>` needs no arguments here."""
    _, inputs, outputs = handoff
    monkeypatch.setattr("sys.argv", _argv(inputs, outputs))
    assert stage_io().inputs["ticker"] == "OTEX"


def test_too_few_arguments_names_the_usage(handoff):
    _, inputs, _ = handoff
    with pytest.raises(StageIOError, match="three handoff arguments"):
        stage_io(["script.py", str(_PLUGIN_ROOT), str(inputs)])


def test_rejects_a_plugin_root_that_is_not_one(handoff, tmp_path):
    _, inputs, outputs = handoff
    with pytest.raises(StageIOError, match="not a plugin root"):
        stage_io(["script.py", str(tmp_path / "nope"), str(inputs), str(outputs)])


def test_rejects_missing_inputs_file(handoff, tmp_path):
    _, _, outputs = handoff
    missing = tmp_path / "absent.json"
    with pytest.raises(StageIOError, match="stage inputs not found"):
        stage_io(["script.py", str(_PLUGIN_ROOT), str(missing), str(outputs)])


def test_rejects_non_object_inputs(handoff):
    _, inputs, outputs = handoff
    inputs.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(StageIOError, match="must hold a JSON object"):
        stage_io(_argv(inputs, outputs))


# ─── Derived paths ───────────────────────────────────────────────────────────


def test_deal_dir_is_derived_from_the_inputs_path(handoff):
    deal_dir, inputs, outputs = handoff
    assert stage_io(_argv(inputs, outputs)).deal_dir == deal_dir.resolve()


def test_artefacts_dir_is_created_under_the_deal_dir(handoff):
    deal_dir, inputs, outputs = handoff
    artefacts = stage_io(_argv(inputs, outputs)).artefacts_dir
    assert artefacts == (deal_dir / "artefacts").resolve()
    assert artefacts.is_dir()


def test_deal_dir_raises_outside_a_deal_directory(tmp_path):
    """Finding deal.json is also the proof we landed somewhere real — a fixed
    'four levels up' would happily return a directory that isn't a deal."""
    stray = tmp_path / "somewhere" / "inputs.json"
    stray.parent.mkdir(parents=True)
    stray.write_text("{}", encoding="utf-8")
    with pytest.raises(StageIOError, match="not inside a deal directory"):
        deal_dir_for(stray)


# ─── Writing back ────────────────────────────────────────────────────────────


def test_write_persists_the_handoff(handoff):
    _, inputs, outputs = handoff
    io = stage_io(_argv(inputs, outputs))
    assert io.write({"workbook_path": Path("/wb.xlsx")}) == outputs.resolve()
    # Path values serialize (json.dumps default=str) rather than raising.
    assert json.loads(outputs.read_text(encoding="utf-8")) == {
        "workbook_path": str(Path("/wb.xlsx"))
    }


def test_fail_writes_an_error_key(handoff):
    _, inputs, outputs = handoff
    stage_io(_argv(inputs, outputs)).fail("missing input: ticker")
    assert json.loads(outputs.read_text(encoding="utf-8")) == {"error": "missing input: ticker"}


# ─── Drift lock ──────────────────────────────────────────────────────────────

#: The conductor's own docs are exempt on both counts. It is the top-level skill
#: the plugin harness invokes, not a dispatched sub-agent, so `CLAUDE_PLUGIN_ROOT`
#: is the only plugin root it can have; and `stage-envelope.md` has to be able to
#: *name* the four variables in order to say they are gone. What a sub-agent
#: actually receives is the envelope TEMPLATE, checked separately below.
_EXEMPT_PREFIX = "conductor/"

_BANNED_PATTERNS = {
    "an export of a handoff path": re.compile(r"(export|\$env:)\s*(STAGE_INPUTS|STAGE_OUTPUTS|DEAL_DIR)"),
    "a handoff path read from the environment": re.compile(
        r"os\.environ(\.get)?[\[(]\s*[\"'](STAGE_INPUTS|STAGE_OUTPUTS|DEAL_DIR)[\"']"
    ),
    "a shell expansion of a handoff path": re.compile(r"\$\{?(STAGE_INPUTS|STAGE_OUTPUTS|DEAL_DIR)\b"),
}

_PLUGIN_ROOT_PATTERN = re.compile(r"CLAUDE_PLUGIN_ROOT")


def _dispatched_skill_docs() -> list[Path]:
    """Every skill doc a dispatched sub-agent reads — i.e. all but the conductor's."""
    return [
        p
        for p in sorted(_SKILLS_DIR.rglob("*.md"))
        if not p.relative_to(_SKILLS_DIR).as_posix().startswith(_EXEMPT_PREFIX)
    ]


def test_no_skill_doc_reaches_for_a_handoff_environment_variable():
    """The Phase E contract, locked.

    Phase E replaced the `export STAGE_INPUTS / STAGE_OUTPUTS / DEAL_DIR /
    CLAUDE_PLUGIN_ROOT` block with three command-line arguments. The exports were
    the most fragile contract in the system — they survived only if the sub-agent
    ran everything in one shell session — so a reappearance is a regression, not a
    style question.
    """
    offenders = []
    for doc in _dispatched_skill_docs():
        text = doc.read_text(encoding="utf-8")
        for what, pattern in _BANNED_PATTERNS.items():
            for match in pattern.finditer(text):
                line = text[: match.start()].count("\n") + 1
                offenders.append(
                    f"{doc.relative_to(_SKILLS_DIR)}:{line} — {what}: {match.group(0)!r}"
                )
    assert not offenders, "handoff paths must come from sys.argv:\n" + "\n".join(offenders)


def test_no_dispatched_skill_reads_claude_plugin_root():
    """A dispatched stage takes the plugin root as `sys.argv[1]`.

    The old `os.environ.get("CLAUDE_PLUGIN_ROOT", "./infor-beta")` idiom carried a
    **cwd-relative** fallback, so a sub-agent whose cwd was anywhere else resolved
    the plugin to a directory that does not exist — and did it silently.
    """
    offenders = []
    for doc in _dispatched_skill_docs():
        text = doc.read_text(encoding="utf-8")
        for match in _PLUGIN_ROOT_PATTERN.finditer(text):
            line = text[: match.start()].count("\n") + 1
            offenders.append(f"{doc.relative_to(_SKILLS_DIR)}:{line}")
    assert not offenders, (
        "a dispatched stage must take the plugin root from sys.argv[1], not the "
        "environment:\n" + "\n".join(offenders)
    )


def test_the_rendered_envelope_template_is_free_of_the_env_handoff():
    """The template is what a sub-agent actually receives.

    `stage-envelope.md`'s surrounding prose is allowed to name the four variables
    in order to record that they are gone; the fenced template between
    "## Template" and its closing fence is not.
    """
    from conductor import _extract_template

    doc = (_SKILLS_DIR / "conductor" / "references" / "stage-envelope.md").read_text(
        encoding="utf-8"
    )
    template = _extract_template(doc)

    for what, pattern in _BANNED_PATTERNS.items():
        assert not pattern.search(template), f"the envelope template still carries {what}"
    assert not _PLUGIN_ROOT_PATTERN.search(template)
    assert (
        'python <your_script.py> "{{plugin_root}}" "{{stage_inputs_path}}" '
        '"{{stage_outputs_path}}"' in template
    )
    assert "{{resolved_inputs_block}}" in template
