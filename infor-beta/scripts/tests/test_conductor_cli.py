"""Unit tests for the thin conductor driver (conductor_cli).

Exercises prep-wave reference resolution (including the pydantic `Company` / `Path`
serialization that previously crashed a live run, and the optional-plan-input ->
None softening) and collect-wave validation. No Agent calls; everything runs off a
hand-built run directory.
"""

import json
from pathlib import Path

import pytest

import conductor_cli
from conductor_cli import (
    collect_wave,
    load_plan,
    load_plan_inputs,
    prep_wave,
    write_plan_inputs,
)
from deal_init import save_deal_context
from plan_refs import PlanReferenceError
from run_log import create_run_dir, stage_dir, write_plan_snapshot
from schemas import Company, DealContext

_PLAN_YAML = """\
deliverable_type: pitch
description: conductor_cli test plan
plan_inputs:
  - name: reporting_quarter
    type: str
    required: true
  - name: risk_notes
    type: str
    required: false
stages:
  - id: alpha
    skill: comps
    inputs:
      ticker: $deal.subject_company.ticker
      company: $deal.subject_company
      quarter: $plan_inputs.reporting_quarter
      notes: $plan_inputs.risk_notes
      template: $deal.deal_dir
    outputs:
      - name: workbook_path
        type: Path
  - id: beta
    skill: deck-assembler
    inputs:
      upstream: $stages.alpha.workbook_path
    outputs:
      - name: deck_path
        type: Path
"""

# The real plugin root (so the stage-envelope template resolves deterministically
# regardless of CLAUDE_PLUGIN_ROOT in the environment).
_PLUGIN_ROOT = Path(conductor_cli.__file__).resolve().parent.parent


@pytest.fixture
def run_dir(tmp_path: Path) -> Path:
    deal_dir = tmp_path / "Project OpenText"
    ctx = DealContext(
        codename="Project OpenText",
        deal_dir=deal_dir,
        deliverable_type="pitch",
        subject_company=Company(legal_name="OpenText Corporation", ticker="OTEX"),
    )
    save_deal_context(ctx)  # writes deal.json + bootstraps subdirs
    rd = create_run_dir(deal_dir, "2026-06-30-pitch-test01")
    write_plan_snapshot(rd, _PLAN_YAML)
    write_plan_inputs(rd, {"reporting_quarter": "Q4 2025"})  # risk_notes left unsupplied
    return rd


def test_load_plan_round_trips(run_dir: Path):
    plan = load_plan(run_dir)
    assert [s.id for s in plan.stages] == ["alpha", "beta"]


def test_prep_wave_resolves_refs_and_serializes_company_and_path(run_dir: Path):
    prepared = prep_wave(run_dir, 1, plugin_root=_PLUGIN_ROOT)
    assert [p["stage_id"] for p in prepared] == ["alpha"]

    # inputs.json was written and reloads cleanly — the Company (from
    # $deal.subject_company) serialized to a dict and the Path (from $deal.deal_dir)
    # to a string, instead of raising "object of type Company is not JSON
    # serializable".
    inputs_file = stage_dir(run_dir, "alpha") / "inputs.json"
    assert inputs_file.exists()
    data = json.loads(inputs_file.read_text(encoding="utf-8"))
    assert data["ticker"] == "OTEX"
    assert isinstance(data["company"], dict)
    assert data["company"]["legal_name"] == "OpenText Corporation"
    assert data["quarter"] == "Q4 2025"
    assert isinstance(data["template"], str)  # Path serialized
    # Optional plan input the analyst didn't supply -> None, not a crash.
    assert data["notes"] is None


def test_prep_wave_renders_envelope_with_absolute_paths(run_dir: Path):
    prepared = prep_wave(run_dir, 1, plugin_root=_PLUGIN_ROOT)
    env = prepared[0]["envelope"]
    assert env is not None
    # The export block carries the absolute inputs path (no env var on the Task call).
    assert "export STAGE_INPUTS=" in env
    assert prepared[0]["inputs_path"] in env
    # The task points at the stage's skill SKILL.md.
    assert "skills/comps/SKILL.md" in env


def test_prep_wave_two_resolves_stage_outputs(run_dir: Path):
    # Wave 1 must have produced alpha's outputs before wave 2 can resolve.
    prep_wave(run_dir, 1, plugin_root=_PLUGIN_ROOT)
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"workbook_path": "/deals/wb.xlsx"}), encoding="utf-8"
    )

    prepared = prep_wave(run_dir, 2, plugin_root=_PLUGIN_ROOT)
    assert [p["stage_id"] for p in prepared] == ["beta"]
    data = json.loads((stage_dir(run_dir, "beta") / "inputs.json").read_text(encoding="utf-8"))
    assert data["upstream"] == "/deals/wb.xlsx"


def test_prep_wave_two_raises_when_prior_outputs_missing(run_dir: Path):
    # Wave 2 needs alpha's outputs.json; without it, resolution can't proceed.
    with pytest.raises(FileNotFoundError):
        prep_wave(run_dir, 2, plugin_root=_PLUGIN_ROOT)


def test_prep_wave_out_of_range(run_dir: Path):
    with pytest.raises(IndexError):
        prep_wave(run_dir, 3, plugin_root=_PLUGIN_ROOT)


def test_collect_wave_ok(run_dir: Path):
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"workbook_path": "/deals/wb.xlsx"}), encoding="utf-8"
    )
    results = collect_wave(run_dir, 1)
    assert results == [
        {"stage_id": "alpha", "ok": True, "outputs": {"workbook_path": "/deals/wb.xlsx"}}
    ]


def test_collect_wave_missing_outputs_reports_failure(run_dir: Path):
    results = collect_wave(run_dir, 1)
    assert results[0]["stage_id"] == "alpha"
    assert results[0]["ok"] is False
    assert "outputs.json" in results[0]["error"]


def test_collect_wave_error_key_reports_failure(run_dir: Path):
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"error": "missing input: ticker"}), encoding="utf-8"
    )
    results = collect_wave(run_dir, 1)
    assert results[0]["ok"] is False
    assert results[0]["error"] == "missing input: ticker"


def test_collect_wave_malformed_outputs_reports_failure(run_dir: Path):
    """A sub-agent that truncates outputs.json mid-write must produce a [FAIL]
    result, not crash the driver with an unhandled json.JSONDecodeError."""
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        '{"workbook_path": "/deals/wb', encoding="utf-8"  # truncated JSON
    )
    results = collect_wave(run_dir, 1)
    assert results[0]["stage_id"] == "alpha"
    assert results[0]["ok"] is False
    assert "malformed" in results[0]["error"]


def test_cli_main_collect_malformed_outputs_engages_fail_path(run_dir: Path, capsys):
    """End-to-end through the argv entrypoint: malformed JSON prints [FAIL] and
    exits non-zero instead of raising."""
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text("not json at all", encoding="utf-8")
    rc = conductor_cli.main(["collect-wave", str(run_dir), "1"])
    assert rc == 1
    assert "[FAIL] alpha" in capsys.readouterr().out


def test_collect_wave_missing_declared_output_reports_failure(run_dir: Path):
    """alpha declares `workbook_path`; an outputs.json without that key fails
    with an error that names the missing output."""
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"something_else": 1}), encoding="utf-8"
    )
    results = collect_wave(run_dir, 1)
    assert results[0]["ok"] is False
    assert "workbook_path" in results[0]["error"]
    assert results[0]["outputs"] == {"something_else": 1}  # kept for debugging


def test_collect_wave_null_declared_output_passes(run_dir: Path):
    """Null values are legal per the v0.5.21 contract (ltm-metrics emits null —
    never omits — ltm_revenue/ltm_adj_ebitda): presence is checked, not truthiness."""
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"workbook_path": None}), encoding="utf-8"
    )
    results = collect_wave(run_dir, 1)
    assert results == [{"stage_id": "alpha", "ok": True, "outputs": {"workbook_path": None}}]


def test_collect_wave_extra_undeclared_keys_pass(run_dir: Path):
    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"workbook_path": "/deals/wb.xlsx", "notes": "extra is fine"}),
        encoding="utf-8",
    )
    results = collect_wave(run_dir, 1)
    assert results[0]["ok"] is True


def test_load_plan_rejects_typod_stage_reference(run_dir: Path):
    """Reference pre-flight at load time: a `$stages` ref to a stage id that
    doesn't exist in the plan is rejected when the snapshot is loaded, before
    any wave is prepped or dispatched."""
    bad_yaml = _PLAN_YAML.replace("$stages.alpha.workbook_path", "$stages.alhpa.workbook_path")
    write_plan_snapshot(run_dir, bad_yaml)
    with pytest.raises(PlanReferenceError, match="alhpa"):
        load_plan(run_dir)
    # prep_wave and collect_wave load through the same path, so both refuse too.
    with pytest.raises(PlanReferenceError):
        prep_wave(run_dir, 1, plugin_root=_PLUGIN_ROOT)
    with pytest.raises(PlanReferenceError):
        collect_wave(run_dir, 1)


def test_load_plan_rejects_undeclared_output_reference(run_dir: Path):
    """A ref to a real stage but an output name it never declares is equally dead
    at load time."""
    bad_yaml = _PLAN_YAML.replace("$stages.alpha.workbook_path", "$stages.alpha.deck_path")
    write_plan_snapshot(run_dir, bad_yaml)
    with pytest.raises(PlanReferenceError, match="deck_path"):
        load_plan(run_dir)


def test_plan_inputs_round_trip_serializes_path(run_dir: Path):
    # write_plan_inputs routes through _json_default, so a Path value is stored as a
    # string and reloads without error.
    write_plan_inputs(run_dir, {"reporting_quarter": "Q1 2026", "snip": Path("/tmp/x.png")})
    loaded = load_plan_inputs(run_dir)
    assert loaded["reporting_quarter"] == "Q1 2026"
    assert loaded["snip"] == str(Path("/tmp/x.png"))


def test_cli_main_prep_and_collect(run_dir: Path, capsys):
    # End-to-end through the argv entrypoint.
    rc = conductor_cli.main(["prep-wave", str(run_dir), "1", "--plugin-root", str(_PLUGIN_ROOT)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "stage `alpha`" in out

    (stage_dir(run_dir, "alpha") / "outputs.json").write_text(
        json.dumps({"workbook_path": "/deals/wb.xlsx"}), encoding="utf-8"
    )
    rc = conductor_cli.main(["collect-wave", str(run_dir), "1"])
    assert rc == 0
    assert "[ok]   alpha" in capsys.readouterr().out
