"""Unit tests for the Plan / Stage schemas."""

import pytest
import yaml
from pathlib import Path
from pydantic import ValidationError

from schemas import InputSpec, OutputSpec, Plan, Stage


def _base_stage(**overrides):
    kwargs = dict(
        id="earnings_update",
        skill="earningsupdate-infor",
        inputs={"company": "$deal.subject_company"},
        outputs=[OutputSpec(name="deck_path", type="Path")],
    )
    kwargs.update(overrides)
    return Stage(**kwargs)


def _base_plan(**overrides):
    kwargs = dict(
        deliverable_type="earnings-update",
        description="Quarterly earnings update deck + companion cap table.",
        plan_inputs=[InputSpec(name="reporting_quarter", type="str")],
        stages=[_base_stage()],
    )
    kwargs.update(overrides)
    return Plan(**kwargs)


def test_minimal_plan():
    p = _base_plan()
    assert p.deliverable_type == "earnings-update"
    assert len(p.stages) == 1
    assert p.stages[0].id == "earnings_update"


def test_default_checkpoint_is_informational():
    s = _base_stage()
    assert s.checkpoint == "informational"


def test_explicit_checkpoint_modes_accepted():
    for mode in ("required", "informational", "silent"):
        s = _base_stage(checkpoint=mode)
        assert s.checkpoint == mode


def test_invalid_checkpoint_rejected():
    with pytest.raises(ValidationError):
        _base_stage(checkpoint="halt")  # type: ignore[arg-type]


def test_plan_requires_at_least_one_stage():
    with pytest.raises(ValidationError):
        _base_plan(stages=[])


def test_duplicate_stage_ids_rejected():
    with pytest.raises(ValidationError):
        _base_plan(
            stages=[
                _base_stage(id="s1"),
                _base_stage(id="s1"),
            ]
        )


def test_two_stage_plan_round_trip():
    p = _base_plan(
        stages=[
            _base_stage(id="earnings_update"),
            _base_stage(
                id="captable",
                skill="captable-infor",
                inputs={"ticker": "$deal.subject_company.ticker"},
                outputs=[OutputSpec(name="workbook_path", type="Path")],
            ),
        ]
    )
    raw = p.model_dump_json()
    p2 = Plan.model_validate_json(raw)
    assert p2 == p
    assert [s.id for s in p2.stages] == ["earnings_update", "captable"]


def test_phase3_earnings_update_plan_has_decomposed_stage_order():
    plan_path = Path(__file__).resolve().parents[2] / "plans" / "earnings-update.yaml"
    plan = Plan.model_validate(yaml.safe_load(plan_path.read_text(encoding="utf-8")))

    assert [stage.id for stage in plan.stages] == ["wireframe", "content", "captable", "deck"]
    assert [stage.skill for stage in plan.stages] == [
        "earningsupdate-wireframe-infor",
        "earningsupdate-content-infor",
        "captable-infor",
        "deck-assembler",
    ]
    assert plan.stages[1].inputs["slide_plan_path"] == "$stages.wireframe.slide_plan_path"
    assert plan.stages[3].inputs["content_bundle_path"] == "$stages.content.content_bundle_path"
    assert plan.stages[3].inputs["captable_workbook_path"] == "$stages.captable.workbook_path"


def test_invalid_deliverable_type_rejected():
    with pytest.raises(ValidationError):
        _base_plan(deliverable_type="management-presentation")  # removed B2


def test_empty_stage_id_rejected():
    with pytest.raises(ValidationError):
        _base_stage(id="")


def test_empty_skill_name_rejected():
    with pytest.raises(ValidationError):
        _base_stage(skill="")


def test_extra_fields_rejected_on_stage():
    with pytest.raises(ValidationError):
        Stage(
            id="x",
            skill="y",
            depends_on=["other-stage"],  # type: ignore[call-arg]
        )
