"""Unit tests for the locked deck-spec questionnaires + answer converters."""

from pathlib import Path

import pytest
import yaml

from deck_spec import (
    EARNINGS_UPDATE_ITEM_PLAN_INPUTS,
    PITCH_ITEM_PLAN_INPUTS,
    market_entry_targets_from_slides,
    metric_count_from_slides,
    render_deck_spec_prompt,
)
from schemas import Plan

PLUGIN_ROOT = Path(__file__).resolve().parents[2]


def _plan(name: str) -> Plan:
    text = (PLUGIN_ROOT / "plans" / name).read_text(encoding="utf-8")
    return Plan.model_validate(yaml.safe_load(text))


def test_pitch_prompt_covers_every_questionnaire_topic():
    prompt = render_deck_spec_prompt("pitch")
    for token in (
        "Client name",
        "Presentation date",
        "Analyst notes",
        "Reporting quarter",
        "Comparison quarter",
        "Financial Summary slides",
        "Acquisition-target slides",
        "Key Investment Highlights",
        "Section divider labels",
        "Valuation range",
        "Risk notes",
        "CIM",
        "SEDI",
        "Bloomberg ownership export",
    ):
        assert token in prompt, f"pitch deck-spec prompt lost its {token!r} item"
    # Rendered verbatim, twice the same — the consistency contract.
    assert render_deck_spec_prompt("pitch") == prompt


def test_earnings_update_prompt_covers_every_questionnaire_topic():
    prompt = render_deck_spec_prompt("earnings-update")
    for token in ("Reporting quarter", "Comparison quarter", "Bloomberg EEO snip"):
        assert token in prompt
    # The EU deck has no slide options — the prompt must say so, not offer any.
    assert "no slide options" in prompt


def test_unknown_deliverable_raises():
    with pytest.raises(ValueError):
        render_deck_spec_prompt("overview")


def test_pitch_item_table_maps_onto_real_plan_inputs():
    names = {spec.name for spec in _plan("pitch.yaml").plan_inputs}
    assert set(PITCH_ITEM_PLAN_INPUTS.values()) <= names


def test_earnings_item_table_maps_onto_real_plan_inputs():
    names = {spec.name for spec in _plan("earnings-update.yaml").plan_inputs}
    assert set(EARNINGS_UPDATE_ITEM_PLAN_INPUTS.values()) <= names


def test_slide_count_converters():
    assert metric_count_from_slides(1) == 4
    assert metric_count_from_slides(2) == 8
    assert market_entry_targets_from_slides(1) == 2
    assert market_entry_targets_from_slides(4) == 8


def test_slide_count_converters_reject_out_of_range():
    with pytest.raises(ValueError):
        metric_count_from_slides(0)
    with pytest.raises(ValueError):
        market_entry_targets_from_slides(0)
    # PitchDeckContent caps market-entry targets at 8 (4 slides).
    with pytest.raises(ValueError):
        market_entry_targets_from_slides(5)
