"""Unit tests for the plan_refs reference resolver."""

from pathlib import Path

import pytest

from plan_refs import (
    ReferenceResolutionError,
    UnknownReferenceError,
    resolve_refs,
)
from schemas import Company, DealContext


@pytest.fixture
def ctx():
    return DealContext(
        codename="Project OpenText",
        deal_dir=Path("/tmp/Project OpenText"),
        deliverable_type="earnings-update",
        subject_company=Company(legal_name="OpenText Corporation", ticker="OTEX"),
    )


def _resolve(value, *, plan_inputs=None, ctx=None, stage_outputs=None):
    return resolve_refs(
        value,
        plan_inputs=plan_inputs or {},
        deal_context=ctx,
        stage_outputs=stage_outputs or {},
    )


def test_plan_input_lookup(ctx):
    out = _resolve("$plan_inputs.reporting_quarter", plan_inputs={"reporting_quarter": "Q4 2025"}, ctx=ctx)
    assert out == "Q4 2025"


def test_deal_top_level_attribute(ctx):
    out = _resolve("$deal.codename", ctx=ctx)
    assert out == "Project OpenText"


def test_deal_dotted_attribute(ctx):
    out = _resolve("$deal.subject_company.ticker", ctx=ctx)
    assert out == "OTEX"


def test_stage_output_lookup(ctx):
    stage_outputs = {"earnings_update": {"deck_path": "/tmp/deck.pptx"}}
    out = _resolve("$stages.earnings_update.deck_path", ctx=ctx, stage_outputs=stage_outputs)
    assert out == "/tmp/deck.pptx"


def test_dict_walked_recursively(ctx):
    inputs = {
        "ticker": "$deal.subject_company.ticker",
        "quarter": "$plan_inputs.q",
        "nested": {"codename": "$deal.codename"},
    }
    out = _resolve(inputs, plan_inputs={"q": "Q4 2025"}, ctx=ctx)
    assert out == {
        "ticker": "OTEX",
        "quarter": "Q4 2025",
        "nested": {"codename": "Project OpenText"},
    }


def test_list_walked_recursively(ctx):
    out = _resolve(["$deal.codename", "literal", 42], ctx=ctx)
    assert out == ["Project OpenText", "literal", 42]


def test_plain_string_passes_through(ctx):
    assert _resolve("hello world", ctx=ctx) == "hello world"


def test_non_string_scalars_pass_through(ctx):
    assert _resolve(42, ctx=ctx) == 42
    assert _resolve(True, ctx=ctx) is True
    assert _resolve(None, ctx=ctx) is None


def test_string_with_dollar_in_middle_passes_through(ctx):
    """Mid-string interpolation is not supported — these strings are treated as literals."""
    assert _resolve("price is $5.00", ctx=ctx) == "price is $5.00"


def test_unknown_prefix_raises(ctx):
    with pytest.raises(UnknownReferenceError):
        _resolve("$config.something", ctx=ctx)


def test_missing_plan_input_raises(ctx):
    with pytest.raises(ReferenceResolutionError):
        _resolve("$plan_inputs.does_not_exist", plan_inputs={}, ctx=ctx)


def test_missing_deal_field_raises(ctx):
    with pytest.raises(ReferenceResolutionError):
        _resolve("$deal.no_such_field", ctx=ctx)


def test_missing_stage_raises(ctx):
    with pytest.raises(ReferenceResolutionError):
        _resolve("$stages.unrun.output", ctx=ctx, stage_outputs={})


def test_missing_stage_output_field_raises(ctx):
    stage_outputs = {"earnings_update": {"deck_path": "/tmp/deck.pptx"}}
    with pytest.raises(ReferenceResolutionError):
        _resolve("$stages.earnings_update.workbook_path", ctx=ctx, stage_outputs=stage_outputs)


def test_dotted_path_through_none_raises(ctx):
    # subject_company is set, but its ticker is None when company is private
    private_ctx = DealContext(
        codename="Project Private",
        deal_dir=Path("/tmp/Project Private"),
        deliverable_type="pitch",
        subject_company=Company(legal_name="Acme LLC"),
    )
    # ticker IS None but exists as an attribute → returns None, doesn't raise
    out = _resolve("$deal.subject_company.ticker", ctx=private_ctx)
    assert out is None
    # but trying to traverse PAST None should raise
    with pytest.raises(ReferenceResolutionError):
        _resolve("$deal.subject_company.ticker.length", ctx=private_ctx)


def test_stages_requires_id_and_field(ctx):
    """`$stages.x` alone (no output name) is invalid."""
    with pytest.raises(ReferenceResolutionError):
        _resolve("$stages.earnings_update", ctx=ctx, stage_outputs={"earnings_update": {}})
