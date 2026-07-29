"""Unit tests for the locked deck-spec dialogs, prompts + answer converters."""

from datetime import date
from pathlib import Path

import pytest
import yaml

from deal_init import INIT_INTAKE
from deck_spec import (
    EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS,
    EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS,
    EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS,
    EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
    EARNINGS_UPDATE_INTAKE,
    EARNINGS_UPDATE_ITEM_PLAN_INPUTS,
    NO_NOTES_ANALYST_NOTES,
    PITCH_ATTACHMENT_PLAN_INPUTS,
    PITCH_DEFAULT_SUPPLIED_INPUTS,
    PITCH_DEFAULT_UNSET_INPUTS,
    PITCH_DIALOG_PLAN_INPUTS,
    PITCH_INTAKE,
    PITCH_ITEM_PLAN_INPUTS,
    default_presentation_date,
    market_entry_targets_from_slides,
    metric_count_from_slides,
    prior_year_quarter,
    render_deck_spec_dialogs,
    render_deck_spec_prompt,
    render_run_attachment_request,
    render_run_defaults,
    render_run_dialogs,
)
from intake_spec import (
    ATTACHMENT_OPTIONAL_HEADER,
    ATTACHMENT_REQUIRED_HEADER,
    render_attachment_request,
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
    # And merged with deal-init's — the payload a slash-command run actually
    # sends, which is the one that has to satisfy the tool's shape contract.
    _assert_askuserquestion_shape(render_run_dialogs("pitch"))
    _assert_askuserquestion_shape(render_run_dialogs("earnings-update"))


def test_a_slash_command_run_asks_once():
    """One `AskUserQuestion` call per `/pitch` or `/earnings-update` run.

    The same property `test_intake_spec.test_a_slash_command_run_renders_exactly_
    one_dialog` pins from the spec side, asserted here through the deck-spec API
    the conductor actually calls, plus the per-question count that makes it fit.
    """
    for deliverable, questions in (("pitch", 4), ("earnings-update", 1)):
        dialogs = render_run_dialogs(deliverable)
        assert len(dialogs) == 1, f"/{deliverable} asks {len(dialogs)} times, not once"
        assert len(dialogs[0]) == questions


def test_earnings_update_asks_nothing_at_all():
    """The earnings-update deck spec has no questions left (v0.5.50).

    Both quarters are inferred from the attached interim filing and echoed for
    override; the EEO snip — its one former question, and a status gate at that
    — is an attachment. Step 4 for this deliverable is the defaults echo plus the
    attachment request, and no dialog.
    """
    assert render_deck_spec_dialogs("earnings-update") == []
    assert EARNINGS_UPDATE_DIALOG_PLAN_INPUTS == {}
    assert EARNINGS_UPDATE_ITEM_PLAN_INPUTS == {}
    # Its one input still arrives, as an attachment carrying a path.
    assert EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS == {
        "eeo_snip_path": "Bloomberg EEO snip"
    }


def test_pitch_dialog_headers_match_plan_input_table():
    headers = [
        q["header"] for dialog in render_deck_spec_dialogs("pitch") for q in dialog
    ]
    assert len(headers) == len(set(headers)), "dialog headers must be unique"
    assert set(headers) == set(PITCH_DIALOG_PLAN_INPUTS)


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
    # The per-deliverable renderers need a questionnaire to render.
    with pytest.raises(ValueError):
        render_deck_spec_dialogs("overview")
    with pytest.raises(ValueError):
        render_deck_spec_prompt("overview")
    # The run-level renderers degrade for a declared stub instead — the deal
    # still has a listing to ask, a sector to default, and filings to request.
    assert render_run_dialogs("overview") == render_run_dialogs()
    assert "Sector / industry" in render_run_defaults("overview", sector="Software")
    # A typo is still an error on all three.
    for render in (
        render_run_attachment_request,
        render_run_dialogs,
        render_run_defaults,
    ):
        with pytest.raises(ValueError, match="unknown deliverable type"):
            render("pitchh")


# ---------------------------------------------------------------------------
# The attachment request — one message, two sections, no questions
# ---------------------------------------------------------------------------


def test_no_deliverable_asks_a_question_about_an_attachment():
    """The v0.5.50 change, stated as the invariant it is.

    A pitch run put three `AskUserQuestion` status gates in front of the analyst
    — SEDI PDF, Bloomberg export, CIM — one dialog at a time, each asking whether
    a file was attached, would be attached next message, or did not exist. All
    three are gone; no rendering of either questionnaire asks about a file.
    """
    for deliverable, spec in (
        ("pitch", PITCH_INTAKE),
        ("earnings-update", EARNINGS_UPDATE_INTAKE),
    ):
        attachment_keys = {f.key for f in spec.attachment_fields()}
        assert attachment_keys, f"{deliverable} declares no attachments"
        headers = {
            q["header"] for d in render_deck_spec_dialogs(deliverable) for q in d
        }
        assert not headers & attachment_keys
        prompt = render_deck_spec_prompt(deliverable)
        for f in spec.attachment_fields():
            # Listed as a bullet, never as a numbered item with options.
            assert f"{f.prompt_label} —" in _flat(prompt)
            assert f not in spec.prompt_fields()


def test_pitch_request_lists_the_filings_then_its_own_documents():
    request = render_run_attachment_request("pitch")
    assert ATTACHMENT_REQUIRED_HEADER in request
    assert ATTACHMENT_OPTIONAL_HEADER in request
    flat = _flat(request)
    # deal-init's G7 filings are the run's one REQUIRED document for a pitch;
    # the three deliverable documents are all degrade-gracefully OPTIONAL.
    assert "Financial statements / filings" in flat
    required_section, optional_section = flat.split(ATTACHMENT_OPTIONAL_HEADER)
    assert "Financial statements / filings" in required_section
    for label in (
        "CIM / management presentation",
        'SEDI "Insider Information by Issuer" PDF',
        "Bloomberg ownership export (.xlsm)",
    ):
        assert label in optional_section, f"{label} is not in the OPTIONAL section"
    # Each bullet carries what the run loses — that warning used to live in a
    # status dialog's option descriptions, which no longer exist.
    for consequence in (
        "insider side stays a placeholder",
        "institutions side stays a placeholder",
        "drafts from the filings and public sources",
    ):
        assert consequence in flat


def test_earnings_update_request_makes_the_eeo_snip_required():
    request = render_run_attachment_request("earnings-update")
    flat = _flat(request)
    required_section = flat.split(ATTACHMENT_OPTIONAL_HEADER)[0]
    assert "Bloomberg EEO snip" in required_section
    assert "Financial statements / filings" in required_section
    # Nothing optional for this deliverable, so no OPTIONAL section at all.
    assert ATTACHMENT_OPTIONAL_HEADER not in request
    # The bullet states the halt, since there is no dialog option to state it in.
    assert "halts here until it arrives" in flat
    assert "SEDI" not in flat


def test_a_deliverable_without_a_questionnaire_still_gets_the_filings():
    """The overview stub and one-off-skill have no deck spec — the deal still
    needs its filings, so the request degrades to deal-init's half."""
    filings_only = render_attachment_request(INIT_INTAKE)
    for deliverable in ("overview", "one-off-skill"):
        assert render_run_attachment_request(deliverable) == filings_only
    assert render_run_attachment_request() == filings_only
    assert "Financial statements / filings" in filings_only
    assert "SEDI" not in filings_only


# ---------------------------------------------------------------------------
# Asked / defaulted split vs. the real plans
# ---------------------------------------------------------------------------


def test_pitch_asked_attached_and_defaulted_inputs_cover_the_plan():
    plan = _plan("pitch.yaml")
    names = {spec.name for spec in plan.plan_inputs}
    asked = set(PITCH_DIALOG_PLAN_INPUTS.values())
    attached = set(PITCH_ATTACHMENT_PLAN_INPUTS)
    supplied = set(PITCH_DEFAULT_SUPPLIED_INPUTS)
    unset = set(PITCH_DEFAULT_UNSET_INPUTS)
    assert asked | attached | supplied | unset <= names
    # An input arrives exactly one way: asked, attached, or defaulted.
    assert not asked & (attached | supplied | unset)
    assert not attached & (supplied | unset)
    assert not supplied & unset
    # Every REQUIRED plan input is asked, attached, or default-supplied — never
    # left to chance.
    required = {spec.name for spec in plan.plan_inputs if spec.required}
    assert required <= asked | attached | supplied
    # Default-supplied inputs are all required (optional defaults stay unset).
    assert supplied <= required
    # The pitch CIM is the optional attachment: no file, no key in plan_inputs.
    assert attached == {"cim_path"}
    assert not attached & required


def test_earnings_update_asked_attached_and_defaulted_inputs_cover_the_plan():
    plan = _plan("earnings-update.yaml")
    names = {spec.name for spec in plan.plan_inputs}
    asked = set(EARNINGS_UPDATE_DIALOG_PLAN_INPUTS.values())
    attached = set(EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS)
    supplied = set(EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS)
    assert asked | attached | supplied <= names
    assert not asked & (attached | supplied)
    assert not attached & supplied
    assert not EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS
    required = {spec.name for spec in plan.plan_inputs if spec.required}
    assert required <= asked | attached | supplied
    # Nothing is asked, so `attached` is carrying the whole non-defaulted half:
    # the EEO snip is REQUIRED, which is why a missing file must halt the run
    # rather than resolve `$plan_inputs.eeo_snip_path` to None.
    assert not asked
    assert attached == {"eeo_snip_path"}
    assert attached <= required


# ---------------------------------------------------------------------------
# Text fallback prompts
# ---------------------------------------------------------------------------


def _flat(text: str) -> str:
    """Collapse whitespace — the generated prompt is wrapped, so a token can
    straddle a line break. Wrapping is a rendering detail; the wording is not."""
    return " ".join(text.split())


def test_pitch_prompt_covers_every_questionnaire_topic():
    prompt = _flat(render_deck_spec_prompt("pitch"))
    for token in (
        # Asked items — three of them since v0.5.51.
        "Analyst notes",
        "CIM",
        "Acquisition-target slides",
        "Key Investment Highlights",
        # Defaulted items (listed so a reply can override them). Valuation range
        # and risk notes moved here from the asked list in v0.5.51: both
        # defaulted to "None" as questions, so nothing is lost by listing them.
        "Valuation range",
        "Risk notes",
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
    # The demoted pair are default lines now, not numbered questions.
    for gone in (
        "Valuation range language for the executive summary?",
        "Any specific risks / mitigants",
    ):
        assert gone not in prompt, f"{gone!r} is still asked as a question"
    # Generated deterministically, twice the same — the consistency contract.
    assert _flat(render_deck_spec_prompt("pitch")) == prompt


def test_earnings_update_prompt_covers_every_questionnaire_topic():
    prompt = _flat(render_deck_spec_prompt("earnings-update"))
    for token in ("Bloomberg EEO snip", "Reporting quarter", "Comparison quarter"):
        assert token in prompt
    # The EU deck has no slide options — the prompt must say so, not offer any.
    assert "no slide options" in prompt
    # And it has no questions either, so the prompt says that rather than
    # instructing the analyst to answer items it does not list.
    assert "Nothing to answer here." in prompt
    assert "1. " not in prompt


def test_fallback_prompt_numbering_matches_dialog_order():
    """The numbered text items and the dialogs are the same list, same order."""
    labels = {
        "analyst_notes": "Analyst notes",
        "market_entry_target_count": "Acquisition-target slides",
        "include_investment_highlights": "Key Investment Highlights",
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
    # The CIM and the EEO snip left this numbering when they became attachments;
    # neither is a numbered item in any prompt any more. The valuation range and
    # the risk notes left it in v0.5.51 when they became defaults — the numbering
    # closed up behind them, which is why the tables are derived and not written.
    assert "cim_path" not in PITCH_ITEM_PLAN_INPUTS.values()
    assert EARNINGS_UPDATE_ITEM_PLAN_INPUTS == {}
    assert PITCH_ITEM_PLAN_INPUTS == {
        1: "analyst_notes",
        2: "market_entry_target_count",
        3: "include_investment_highlights",
    }


def test_prompts_embed_their_half_of_the_attachment_request():
    """Same documents, same order, in the fallback as in the live message.

    The live request merges deal-init's filings with the deliverable's; the text
    fallback is two prompts, each embedding its own spec's half — so between them
    the analyst sees the same list, generated by the same function.
    """
    for deliverable, spec in (
        ("pitch", PITCH_INTAKE),
        ("earnings-update", EARNINGS_UPDATE_INTAKE),
    ):
        half = render_attachment_request(spec)
        assert half.rstrip("\n") in render_deck_spec_prompt(deliverable)
        # Within each section of the merged message, bullets follow spec order:
        # deal-init's filings, then the deliverable's own documents.
        flat = _flat(render_run_attachment_request(deliverable))
        for required in (True, False):
            expected = [
                _flat(f.checklist)
                for f in (
                    *INIT_INTAKE.attachment_fields(required=required),
                    *spec.attachment_fields(required=required),
                )
            ]
            assert all(c in flat for c in expected)
            positions = [flat.index(c) for c in expected]
            assert positions == sorted(positions), (
                f"{deliverable}'s "
                f"{'REQUIRED' if required else 'OPTIONAL'} bullets are out of "
                f"spec order"
            )
    assert "SEDI" not in render_attachment_request(EARNINGS_UPDATE_INTAKE)


# ---------------------------------------------------------------------------
# Defaults echo + computed defaults
# ---------------------------------------------------------------------------


def test_pitch_defaults_echo_lists_every_default():
    """One merged echo per run: deal-init's sector, then the deliverable's."""
    echo = render_run_defaults(
        "pitch",
        sector="Enterprise software",
        client_name="ACME Corp",
        presentation_date="July 2026",
        reporting_quarter="Q2 2026",
        comparison_quarter="Q2 2025",
    )
    for token in (
        "Enterprise software",
        "ACME Corp",
        "July 2026",
        "Q2 2026 vs Q2 2025",
        # The two v0.5.51 demotions — asked through v0.5.50, echoed now.
        "Valuation range",
        "Risk notes",
        "Financial Summary slides",
        "Section divider labels",
        "override",
    ):
        assert token in echo, f"pitch defaults echo lost {token!r}"


def test_earnings_update_defaults_echo():
    echo = render_run_defaults(
        "earnings-update",
        sector="Enterprise software",
        reporting_quarter="Q2 2026",
        comparison_quarter="Q2 2025",
    )
    assert "Q2 2026 vs Q2 2025" in echo
    assert "Enterprise software" in echo
    assert "override" in echo


def test_defaults_echo_requires_the_computed_values():
    with pytest.raises(ValueError, match="presentation_date"):
        render_run_defaults(
            "pitch",
            sector="Enterprise software",
            client_name="ACME Corp",
            reporting_quarter="Q2 2026",
            comparison_quarter="Q2 2025",
        )
    with pytest.raises(ValueError):
        render_run_defaults("earnings-update", reporting_quarter="Q2 2026")
    # The sector is deal-init's, so every run needs it — including a deliverable
    # with no questionnaire of its own.
    with pytest.raises(ValueError, match="sector"):
        render_run_defaults("overview")


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
