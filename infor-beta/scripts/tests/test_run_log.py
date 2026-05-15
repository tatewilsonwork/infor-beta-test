"""Unit tests for the run_log helpers."""

import json
from datetime import date
from pathlib import Path

import pytest

from run_log import (
    create_run_dir,
    make_run_id,
    read_stage_outputs,
    stage_dir,
    write_plan_snapshot,
    write_stage_inputs,
    write_stage_log,
    write_summary,
)


def test_run_id_format():
    rid = make_run_id("earnings-update", today=date(2026, 5, 15))
    parts = rid.split("-")
    # Expected: YYYY-MM-DD-earnings-update-<8hex>
    assert rid.startswith("2026-05-15-earnings-update-")
    assert len(parts[-1]) == 8  # short uuid
    # Hex check on the last segment
    int(parts[-1], 16)


def test_run_id_sanitises_plan_id():
    rid = make_run_id("Earnings Update / V2", today=date(2026, 5, 15))
    assert rid.startswith("2026-05-15-earnings-update-v2-")
    assert "/" not in rid


def test_create_run_dir_makes_tree(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "2026-05-15-test-abcd1234")
    assert run_dir == tmp_path / "runs" / "2026-05-15-test-abcd1234"
    assert run_dir.is_dir()
    assert (run_dir / "stages").is_dir()


def test_create_run_dir_idempotent(tmp_path: Path):
    a = create_run_dir(tmp_path, "rid")
    b = create_run_dir(tmp_path, "rid")
    assert a == b
    assert a.is_dir()


def test_write_plan_snapshot(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    yaml_text = "deliverable_type: earnings-update\n"
    out = write_plan_snapshot(run_dir, yaml_text)
    assert out.read_text(encoding="utf-8") == yaml_text


def test_stage_inputs_round_trip(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    inputs = {"ticker": "OTEX", "quarter": "Q4 2025", "path": Path("/tmp/x.pdf")}
    target = write_stage_inputs(run_dir, "s1", inputs)
    assert target == run_dir / "stages" / "s1" / "inputs.json"
    parsed = json.loads(target.read_text(encoding="utf-8"))
    assert parsed["ticker"] == "OTEX"
    assert parsed["path"] == "/tmp/x.pdf"  # Path serialised as str


def test_read_stage_outputs(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    sd = stage_dir(run_dir, "s1")
    (sd / "outputs.json").write_text('{"deck_path": "/tmp/deck.pptx"}', encoding="utf-8")
    out = read_stage_outputs(run_dir, "s1")
    assert out == {"deck_path": "/tmp/deck.pptx"}


def test_read_stage_outputs_missing_raises(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    with pytest.raises(FileNotFoundError):
        read_stage_outputs(run_dir, "did-not-run")


def test_write_stage_log(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    out = write_stage_log(run_dir, "s1", "transcript line 1\ntranscript line 2\n")
    assert "transcript line 1" in out.read_text(encoding="utf-8")


def test_write_summary(tmp_path: Path):
    run_dir = create_run_dir(tmp_path, "rid")
    out = write_summary(run_dir, "# Run summary\n\nAll good.\n")
    assert out == run_dir / "summary.md"
    assert "Run summary" in out.read_text(encoding="utf-8")
