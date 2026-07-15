"""Unit tests for the locked deck-spec dialogs, prompts + answer converters."""

from datetime import date
from pathlib import Path

import pytest
import yaml

from deck_spec import (
    EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS,
    EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS,
    EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
    EARNINGS_UPDATE_ITEM_PLAN_INPUTS,
    NO_NOTES_ANALYST_NOTES,
    PITCH_DEFAULT_SUPPLIED_INPUTS,
    PITCH_DEFAULT_UNSET_INPUTS,
    PITCH_DIALOG_PLAN_INPUTS,
    PITCH_ITEM_PLAN_INPUTS,
    default_presentation_date,
    market_entry_targets_from_slides,
    metric_count_from_slides,
    prior_year_quarter,
    render_deck_spec_defaults,
    render_deck_spec_dialogs,
    render_deck_spec_documents_note,
    render_deck_spec_prompt,
)
from schemas import Plan

PLUGIN_ROOT = Path(__file__).resolve().parents[2]


def _plan(name: str) -> Plan:
    text = (PLUGIN_ROOT / "plans" / name).read_text(encoding="utf-8")
    return Plan.model_validate(yaml.safe_load(text))


def _assert_askuserquestion_shape(dialogs: list[list[dict]]) -> None:
    """Every dialog must be a valid AskUserQuestion `questions` payload."""
    assert dialogs, "no dialogs"
    for dialog in dialogs:
        assert 1 <= len(dialog) <= 4, "AskUserQuestion holds at most 4 questions"
        for q in dialog:
            assert set(q) == {"question", "header", "multiSelect", "options"}
            assert q["question"].strip().endswith("?")
            assert 1 <= len(q["header"]) <= 12, f"header too long: {q['header']!r}"
            assert q["multiSelect"] is False
            assert 2 <= len(q["options"]) <= 4
            for opt in q["options"]:
                assert set(opt) == {"label", "description"}
                assert opt["label"].strip()
                assert opt["description"].strip()


# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------


def test_pitch_dialogs_are_valid_askuserquestion_payloads():
    _assert_askuserquestion_shape(render_deck_spec_dialogs("pitch"))


def test_earnings_update_dialogs_are_valid_askuserquestion_payloads():
    _assert_askuserquestion_shape(render_deck_spec_dialogs("earnings-update"))


def test_pitch_dialog_headers_match_plan_input_table():
    headers = [
        q["header"] for dialog in render_deck_spec_dialogs("pitch") for q in dialog
    ]
    assert len(headers) == len(set(headers)), "dialog headers must be unique"
    assert set(headers) == set(PITCH_DIALOG_PLAN_INPUTS)


def test_earnings_update_dialog_headers_match_plan_input_table():
    headers = [
        q["header"]
        for dialog in render_deck_spec_dialogs("earnings-update")
        for q in dialog
    ]
    assert len(headers) == len(set(headers))
    assert set(headers) == set(EARNINGS_UPDATE_DIALOG_PLAN_INPUTS)


def test_dialogs_render_verbatim_and_immutably():
    first = render_deck_spec_dialogs("pitch")
    first[0][0]["question"] = "mutated"
    first[0][0]["options"].clear()
    # The consistency contract: a re-render is untouched by caller mutation.
    again = render_deck_spec_dialogs("pitch")
    assert again == render_deck_spec_dialogs("pitch")
    assert again[0][0]["question"] != "mutated"
    assert again[0][0]["options"]


def test_unknown_deliverable_raises_everywhere():
    with pytest.raises(ValueError):
        render_deck_spec_dialogs("overview")
    with pytest.raises(ValueError):
        render_deck_spec_prompt("overview")
    with pytest.raises(ValueError):
        render_deck_spec_documents_note("overview")
    with pytest.raises(ValueError):
        render_deck_spec_defaults("overview")


# ---------------------------------------------------------------------------
# Asked / defaulted split vs. the real plans
# ---------------------------------------------------------------------------


def test_pitch_asked_and_defaulted_inputs_cover_the_plan():
    plan = _plan("pitch.yaml")
    names = {spec.name for spec in plan.plan_inputs}
    asked = set(PITCH_DIALOG_PLAN_INPUTS.values())
    supplied = set(PITCH_DEFAULT_SUPPLIED_INPUTS)
    unset = set(PITCH_DEFAULT_UNSET_INPUTS)
    assert asked | supplied | unset <= names
    # No input is both asked and defaulted.
    assert not asked & (supplied | unset)
    assert not supplied & unset
    # Every REQUIRED plan input is either asked or default-supplied — never
    # left to chance.
    required = {spec.name for spec in plan.plan_inputs if spec.required}
    assert required <= asked | supplied
    # Default-supplied inputs are all required (optional defaults stay unset).
    assert supplied <= required


def test_earnings_update_asked_and_defaulted_inputs_cover_the_plan():
    plan = _plan("earnings-update.yaml")
    names = {spec.name for spec in plan.plan_inputs}
    asked = set(EARNINGS_UPDATE_DIALOG_PLAN_INPUTS.values())
    supplied = set(EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS)
    assert asked | supplied <= names
    assert not asked & supplied
    assert not EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS
    required = {spec.name for spec in plan.plan_inputs if spec.required}
    assert required <= asked | supplied


# ---------------------------------------------------------------------------
# Text fallback prompts
# ---------------------------------------------------------------------------


def test_pitch_prompt_covers_every_questionnaire_topic():
    prompt = render_deck_spec_prompt("pitch")
    for token in (
        # Asked items.
        "Analyst notes",
        "CIM",
        "Valuation range",
        "Risk notes",
        "Acquisition-target slides",
        "Key Investment Highlights",
        # Defaulted items (listed so a reply can override them).
        "Client name on the cover",
        "Presentation date",
        "Reporting quarter",
        "Comparison quarter",
        "Financial Summary slides",
        "Section divider labels",
        # Documents checklist.
        "SEDI",
        "Bloomberg ownership export",
    ):
        assert token in prompt, f"pitch deck-spec prompt lost its {token!r} item"
    # Rendered verbatim, twice the same — the consistency contract.
    assert render_deck_spec_prompt("pitch") == prompt


def test_earnings_update_prompt_covers_every_questionnaire_topic():
    prompt = render_deck_spec_prompt("earnings-update")
    for token in ("Bloomberg EEO snip", "Reporting quarter", "Comparison quarter"):
        assert token in prompt
    # The EU deck has no slide options — the prompt must say so, not offer any.
    assert "no slide options" in prompt


def test_fallback_prompt_numbering_matches_dialog_order():
    """The numbered text items and the dialogs are the same list, same order."""
    labels = {
        "analyst_notes": "Analyst notes",
        "cim_path": "CIM / management pres.",
        "valuation_range": "Valuation range",
        "risk_notes": "Risk notes",
        "market_entry_target_count": "Acquisition-target slides",
        "include_investment_highlights": "Key Investment Highlights",
        "eeo_snip_path": "Bloomberg EEO snip",
    }
    for deliverable, table in (
        ("pitch", PITCH_ITEM_PLAN_INPUTS),
        ("earnings-update", EARNINGS_UPDATE_ITEM_PLAN_INPUTS),
    ):
        prompt = render_deck_spec_prompt(deliverable)
        dialog_inputs = [
            (
                PITCH_DIALOG_PLAN_INPUTS
                if deliverable == "pitch"
                else EARNINGS_UPDATE_DIALOG_PLAN_INPUTS
            )[q["header"]]
            for dialog in render_deck_spec_dialogs(deliverable)
            for q in dialog
        ]
        assert table == {i + 1: name for i, name in enumerate(dialog_inputs)}
        for number, name in table.items():
            assert f"{number}. {labels[name]}" in prompt, (
                f"{deliverable} prompt item {number} is not {labels[name]!r}"
            )


def test_prompts_embed_the_documents_note():
    for deliverable in ("pitch", "earnings-update"):
        note = render_deck_spec_documents_note(deliverable)
        assert note in render_deck_spec_prompt(deliverable)
    assert "SEDI" not in render_deck_spec_documents_note("earnings-update")


# ---------------------------------------------------------------------------
# Defaults echo + computed defaults
# ---------------------------------------------------------------------------


def test_pitch_defaults_echo_lists_every_default():
    echo = render_deck_spec_defaults(
        "pitch",
        client_name="ACME Corp",
        presentation_date="July 2026",
        reporting_quarter="Q2 2026",
        comparison_quarter="Q2 2025",
    )
    for token in (
        "ACME Corp",
        "July 2026",
        "Q2 2026 vs Q2 2025",
        "Financial Summary slides",
        "Section divider labels",
        "override",
    ):
        assert token in echo, f"pitch defaults echo lost {token!r}"


def test_earnings_update_defaults_echo():
    echo = render_deck_spec_defaults(
        "earnings-update", reporting_quarter="Q2 2026", comparison_quarter="Q2 2025"
    )
    assert "Q2 2026 vs Q2 2025" in echo
    assert "override" in echo


def test_defaults_echo_requires_the_computed_values():
    with pytest.raises(ValueError, match="presentation_date"):
        render_deck_spec_defaults(
            "pitch",
            client_name="ACME Corp",
            reporting_quarter="Q2 2026",
            comparison_quarter="Q2 2025",
        )
    with pytest.raises(ValueError):
        render_deck_spec_defaults("earnings-update", reporting_quarter="Q2 2026")


def test_default_presentation_date():
    assert default_presentation_date(date(2026, 7, 15)) == "July 2026"
    assert default_presentation_date(date(2025, 12, 1)) == "December 2025"


def test_prior_year_quarter():
    assert prior_year_quarter("Q2 2026") == "Q2 2025"
    assert prior_year_quarter(" q4 2000 ") == "Q4 1999"


def test_prior_year_quarter_rejects_malformed():
    for bad in ("Q5 2026", "FY2026", "2026 Q2", "Q2", ""):
        with pytest.raises(ValueError):
            prior_year_quarter(bad)


def test_no_notes_literal_is_nonempty_and_stable():
    assert "draft" in NO_NOTES_ANALYST_NOTES.lower()
    assert "filings" in NO_NOTES_ANALYST_NOTES


# ---------------------------------------------------------------------------
# Slide-count converters
# ---------------------------------------------------------------------------


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


def test_target_slide_dialog_options_span_the_valid_range():
    """The Targets question offers exactly the 1–4-slide space the converter allows."""
    dialogs = render_deck_spec_dialogs("pitch")
    targets_q = next(
        q for dialog in dialogs for q in dialog if q["header"] == "Targets"
    )
    slide_counts = sorted(
        int(opt["label"].split()[0]) for opt in targets_q["options"]
    )
    assert slide_counts == [1, 2, 3, 4]
    for n in slide_counts:
        market_entry_targets_from_slides(n)  # none of the options may raise
