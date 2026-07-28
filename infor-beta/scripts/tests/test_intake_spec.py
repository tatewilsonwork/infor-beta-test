"""The H1 drift lock: one declarative spec behind every intake rendering.

The locked-questionnaire principle — *every run asks the same questions, in the
same order, with the same options* — used to be half structural and half
conventional. `_dialog_item_plan_inputs` derived the item -> plan-input
**mapping** from the dialog order, so answer mapping could not drift, but
`_PITCH_SPEC_PROMPT` was a hand-written string literal carrying its own copy of
every question's wording, defaults and option labels. Change a dialog option
and nothing forced the prompt to follow; the analyst on a surface without
`AskUserQuestion` would be offered a choice that no longer existed, and the
suite stayed green.

These tests are the replacement guarantee, in two layers:

  1. **Both renderings are generated** — each public renderer returns exactly
     what `intake_spec` produces from the spec. Reintroducing a hand-written
     prompt literal fails here, in the change that reintroduces it.
  2. **The two renderings describe the same items with the same wording** —
     every dialog question, option label and option description appears in the
     text prompt, in the same order, with nothing in one that is missing from
     the other.

Plus the invariants the intake spec has to keep carrying: the
`*_DIALOG_PLAN_INPUTS` answer mapping, and the attachment gates whose answers
are deliberately NOT plan inputs.
"""

from __future__ import annotations

import pytest
import yaml
from pathlib import Path

import deal_init
import deck_spec
from intake_spec import (
    DIALOG_MAX_QUESTIONS,
    IntakeDefault,
    IntakeField,
    IntakeNote,
    IntakeOption,
    IntakeSpec,
    render_defaults_echo,
    render_dialogs,
    render_note,
    render_prompt,
)
from schemas import Plan

PLUGIN_ROOT = Path(__file__).resolve().parents[2]

# Every shipped questionnaire: (label, spec).
SPECS = [
    ("deal-init", deal_init.INIT_INTAKE),
    ("pitch", deck_spec.PITCH_INTAKE),
    ("earnings-update", deck_spec.EARNINGS_UPDATE_INTAKE),
]
SPEC_IDS = [label for label, _ in SPECS]


def _flat(text: str) -> str:
    """Collapse whitespace so a wrapped line break doesn't hide a token.

    The prompt is wrapped at a fixed width; wrapping is a rendering detail,
    the wording is what must match.
    """
    return " ".join(text.split())


def _plan(name: str) -> Plan:
    text = (PLUGIN_ROOT / "plans" / name).read_text(encoding="utf-8")
    return Plan.model_validate(yaml.safe_load(text))


# ---------------------------------------------------------------------------
# Layer 1 — every rendering is generated from the spec
# ---------------------------------------------------------------------------


def test_deal_init_renderings_are_generated():
    spec = deal_init.INIT_INTAKE
    assert deal_init.render_init_prompt() == render_prompt(spec)
    assert deal_init.render_init_filings_note() == render_note(spec)
    assert deal_init.render_init_dialogs(include_deliverable=True) == render_dialogs(
        spec
    )
    assert deal_init.render_init_dialogs() == render_dialogs(
        spec, omit=(deal_init.DELIVERABLE_FIELD_KEY,)
    )
    assert deal_init.INIT_DIALOG_FIELDS == spec.targets("deal-context")


@pytest.mark.parametrize(
    "deliverable,spec",
    [("pitch", deck_spec.PITCH_INTAKE), ("earnings-update", deck_spec.EARNINGS_UPDATE_INTAKE)],
)
def test_deck_spec_renderings_are_generated(deliverable: str, spec: IntakeSpec):
    assert deck_spec.render_deck_spec_prompt(deliverable) == render_prompt(spec)
    assert deck_spec.render_deck_spec_documents_note(deliverable) == render_note(spec)
    assert deck_spec.render_deck_spec_dialogs(deliverable) == render_dialogs(
        spec, target_kinds=("plan-input",)
    )
    assert deck_spec.render_deck_spec_documents_dialogs(
        deliverable
    ) == render_dialogs(spec, target_kinds=("attachment",))
    echo_values = {
        "client_name": "Example Target Inc.",
        "presentation_date": "July 2026",
        "reporting_quarter": "Q2 2026",
        "comparison_quarter": "Q2 2025",
    }
    assert deck_spec.render_deck_spec_defaults(
        deliverable, **echo_values
    ) == render_defaults_echo(spec, echo_values)


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_generation_is_deterministic(label: str, spec: IntakeSpec):
    """Same spec in, same bytes out — the locked questionnaire, mechanised."""
    assert render_prompt(spec) == render_prompt(spec)
    assert render_dialogs(spec) == render_dialogs(spec)
    assert render_note(spec) == render_note(spec)
    # Fresh payloads every call: a caller mutating one render cannot reach the
    # next, so no deep-copy discipline is needed at the call sites.
    first = render_dialogs(spec)
    first[0][0]["question"] = "mutated"
    first[0][0]["options"].clear()
    again = render_dialogs(spec)
    assert again[0][0]["question"] != "mutated"
    assert again[0][0]["options"]


# ---------------------------------------------------------------------------
# Layer 2 — the deliverable: the two renderings say the same thing
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_dialogs_and_prompt_describe_the_same_items_in_the_same_order(
    label: str, spec: IntakeSpec
):
    """The numbered prompt items and the dialog questions are one list.

    This is the guarantee that replaced `_PITCH_SPEC_PROMPT`: a question the
    dialogs ask is a numbered item in the prompt, at the same position, and
    nothing else is.
    """
    prompt = _flat(render_prompt(spec))
    items = spec.prompt_fields()

    # Same items, same order. Attachment gates are in neither list — they are
    # status gates described by the note (asserted separately below).
    dialog_keys = [
        q["header"]
        for dialog in render_dialogs(
            spec, target_kinds=("plan-input", "deal-context")
        )
        for q in dialog
    ]
    assert dialog_keys == [f.key for f in items if f.is_dialog]

    for number, f in enumerate(items, start=1):
        assert f"{number}. {f.prompt_label}:" in prompt, (
            f"{label} prompt item {number} is not {f.prompt_label!r}"
        )
    # No stray item numbers beyond the field list.
    assert f"{len(items) + 1}. " not in prompt


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_prompt_quotes_every_dialog_question_and_option_verbatim(
    label: str, spec: IntakeSpec
):
    """Same wording, not merely the same topics.

    The old failure mode was a renamed option label: the dialog offered
    "Include — draft from my notes" while the prompt still said
    "include — drafted from your notes". Both now come from one string.
    """
    prompt = _flat(render_prompt(spec))
    for f in spec.prompt_fields():
        if f.is_dialog:
            assert _flat(f.question) in prompt, (
                f"{label} prompt does not ask {f.key!r}'s question verbatim"
            )
            for opt in f.options:
                assert _flat(opt.label) in prompt, (
                    f"{label} prompt is missing {f.key!r} option {opt.label!r}"
                )
                assert _flat(opt.description) in prompt, (
                    f"{label} prompt is missing {f.key!r} option "
                    f"{opt.label!r}'s description"
                )
        else:
            # A free-text item has no dialog wording to share — the hint is it.
            assert _flat(f.hint) in prompt


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_prompt_marks_required_items_and_brackets_defaults(
    label: str, spec: IntakeSpec
):
    """REQUIRED and [default] are derived, so they cannot contradict the options."""
    prompt = _flat(render_prompt(spec))
    for number, f in enumerate(spec.prompt_fields(), start=1):
        head = f"{number}. {f.prompt_label}:"
        if f.required:
            assert f"{head} REQUIRED" in prompt
            assert f.default_option is None
        elif f.default_option is not None:
            assert f"{head} [{f.default_option.label}]" in prompt
    if any(f.required for f in spec.prompt_fields()):
        assert "Items marked REQUIRED have no default." in prompt


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_every_default_states_one_rule_in_the_prompt_and_the_echo(
    label: str, spec: IntakeSpec
):
    """A default's `rule` is the single statement both renderings carry."""
    if not spec.defaults:
        return
    prompt = _flat(render_prompt(spec))
    assert "Defaulted unless you override here" in prompt
    for d in spec.defaults:
        assert f"- {d.label}:" in prompt
        assert _flat(d.rule) in prompt, f"{label} prompt lost the {d.name} rule"
    # A static default (no computed value) echoes the same rule the prompt
    # lists; a computed one echoes its filled-in template.
    echo = render_defaults_echo(
        spec,
        {
            "client_name": "Example Target Inc.",
            "presentation_date": "July 2026",
            "reporting_quarter": "Q2 2026",
            "comparison_quarter": "Q2 2025",
        },
    )
    for d in spec.defaults:
        if not d.echoed:
            continue
        assert f"- {d.echo_line_label}:" in echo
        if d.echo is None:
            assert d.rule in echo


def test_defaults_echo_names_every_missing_computed_value():
    with pytest.raises(ValueError, match="client_name, presentation_date"):
        render_defaults_echo(deck_spec.PITCH_INTAKE, {"reporting_quarter": "Q2 2026"})


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_note_describes_every_attachment_gate(label: str, spec: IntakeSpec):
    """A document the run asks about is a document the checklist explains.

    Attachment gates are not numbered prompt items — file bytes cannot come
    through a dialog, so the note is where the analyst reads what to attach and
    what degrades without it.
    """
    note = _flat(render_note(spec))
    gates = [f for f in spec.fields if f.target_kind == "attachment"]
    prompt = _flat(render_prompt(spec))
    assert note in prompt, "the text prompt must embed the checklist"
    for f in gates:
        assert _flat(f.checklist) in note, (
            f"{label} note does not describe the {f.key!r} attachment"
        )
        assert f.key not in spec.targets("plan-input")
        assert f.key in spec.targets("attachment")
        # The status answer never becomes a plan input, and the gate is not a
        # numbered item competing with one.
        assert f not in spec.prompt_fields()


# ---------------------------------------------------------------------------
# The surviving invariants: answer mapping and the attachment/plan-input split
# ---------------------------------------------------------------------------


def test_plan_input_tables_still_key_off_the_dialog_headers():
    for deliverable, table, spec in (
        ("pitch", deck_spec.PITCH_DIALOG_PLAN_INPUTS, deck_spec.PITCH_INTAKE),
        (
            "earnings-update",
            deck_spec.EARNINGS_UPDATE_DIALOG_PLAN_INPUTS,
            deck_spec.EARNINGS_UPDATE_INTAKE,
        ),
    ):
        headers = [
            q["header"]
            for dialog in deck_spec.render_deck_spec_dialogs(deliverable)
            for q in dialog
        ]
        assert set(headers) == set(table)
        assert len(headers) == len(set(headers))
        # Status gates share the header namespace but never the table.
        assert not set(spec.targets("attachment")) & set(table)


def test_item_tables_number_the_prompt_the_dialogs_render():
    for table, spec in (
        (deck_spec.PITCH_ITEM_PLAN_INPUTS, deck_spec.PITCH_INTAKE),
        (
            deck_spec.EARNINGS_UPDATE_ITEM_PLAN_INPUTS,
            deck_spec.EARNINGS_UPDATE_INTAKE,
        ),
    ):
        assert table == {
            number: f.target
            for number, f in enumerate(spec.prompt_fields(), start=1)
        }


def test_declared_requiredness_matches_the_plans():
    """`required` on a field means the plan input has no default — check it."""
    for plan_name, spec in (
        ("pitch.yaml", deck_spec.PITCH_INTAKE),
        ("earnings-update.yaml", deck_spec.EARNINGS_UPDATE_INTAKE),
    ):
        plan_required = {s.name: s.required for s in _plan(plan_name).plan_inputs}
        for f in spec.fields:
            if f.target_kind != "plan-input":
                continue
            assert f.target in plan_required, f"{f.key} asks for an unknown input"
            assert f.required is plan_required[f.target], (
                f"{f.key} is {'REQUIRED' if f.required else 'optional'} in the "
                f"questionnaire but the opposite in {plan_name}"
            )


def test_dialog_batches_respect_the_declared_groups():
    """Batching is by declared group, then capped — not incidental chunking."""
    for _, spec in SPECS:
        for dialog in render_dialogs(spec):
            assert 1 <= len(dialog) <= DIALOG_MAX_QUESTIONS
        batched = [
            [q["header"] for q in dialog] for dialog in render_dialogs(spec)
        ]
        by_key = {f.key: f for f in spec.fields}
        for dialog in batched:
            groups = {by_key[key].group for key in dialog}
            assert len(groups) == 1, f"dialog {dialog} mixes groups {groups}"
    # The pitch questionnaire's two declared groups are its two dialogs.
    assert [
        [q["header"] for q in dialog]
        for dialog in deck_spec.render_deck_spec_dialogs("pitch")
    ] == [["Notes", "CIM", "Valuation", "Risk notes"], ["Targets", "Highlights"]]


# ---------------------------------------------------------------------------
# The spec model rejects a questionnaire that could not render consistently
# ---------------------------------------------------------------------------


def _field(**overrides) -> IntakeField:
    kwargs = dict(
        key="Notes",
        prompt_label="Analyst notes",
        target="analyst_notes",
        question="Notes?",
        options=(IntakeOption("A", "First."), IntakeOption("B", "Second.")),
    )
    kwargs.update(overrides)
    return IntakeField(**kwargs)


def test_header_longer_than_the_dialog_allows_is_rejected():
    with pytest.raises(ValueError, match="AskUserQuestion allows 12"):
        _field(key="Much too long a header")


def test_option_count_outside_the_dialog_range_is_rejected():
    with pytest.raises(ValueError, match="AskUserQuestion allows 2-4"):
        _field(options=(IntakeOption("Only one", "Sole choice."),))
    with pytest.raises(ValueError, match="AskUserQuestion allows 2-4"):
        _field(options=tuple(IntakeOption(f"O{i}", "Choice.") for i in range(5)))


def test_required_field_cannot_also_declare_a_default_option():
    with pytest.raises(ValueError, match="REQUIRED means there is no default"):
        _field(
            required=True,
            options=(
                IntakeOption("A", "First.", default=True),
                IntakeOption("B", "Second."),
            ),
        )


def test_two_default_options_are_rejected():
    with pytest.raises(ValueError, match="more than one default"):
        _field(
            options=(
                IntakeOption("A", "First.", default=True),
                IntakeOption("B", "Second.", default=True),
            )
        )


def test_a_dialog_field_cannot_carry_prompt_only_prose():
    """The hint is the drift surface H1 removed — a dialog field may not have one."""
    with pytest.raises(ValueError, match="drift surface"):
        _field(hint="the raw notes behind the executive summary")


def test_a_free_text_field_needs_a_hint_and_no_question():
    with pytest.raises(ValueError, match="needs a hint"):
        IntakeField(
            key="Company", prompt_label="Subject company name", target="legal_name"
        )
    with pytest.raises(ValueError, match="put its wording in `hint`"):
        IntakeField(
            key="Company",
            prompt_label="Subject company name",
            target="legal_name",
            question="Which company?",
            hint="e.g. ACME",
        )


def test_an_attachment_gate_needs_a_checklist_line():
    with pytest.raises(ValueError, match="needs a checklist line"):
        _field(key="SEDI PDF", target_kind="attachment", prompt_label="")


def test_only_an_attachment_gate_contributes_a_checklist_bullet():
    with pytest.raises(ValueError, match="not an attachment gate"):
        _field(checklist="Attach the thing.")


def _spec(**overrides) -> IntakeSpec:
    kwargs = dict(name="probe", title="Probe", fields=(_field(),))
    kwargs.update(overrides)
    return IntakeSpec(**kwargs)


def test_duplicate_dialog_headers_are_rejected():
    """A reused header would route one question's answer through another's target."""
    with pytest.raises(ValueError, match="reuses dialog header"):
        _spec(fields=(_field(), _field(target="risk_notes")))


def test_a_split_group_is_rejected():
    with pytest.raises(ValueError, match="splits group"):
        _spec(
            fields=(
                _field(key="A", group="one"),
                _field(key="B", group="two"),
                _field(key="C", group="one"),
            )
        )


def test_an_input_cannot_be_both_asked_and_defaulted():
    with pytest.raises(ValueError, match="both asks for and defaults"):
        _spec(
            defaults=(
                IntakeDefault(
                    name="analyst_notes",
                    label="Analyst notes",
                    rule="drafted from the filings",
                    supplied=True,
                ),
            )
        )


def test_an_attachment_gate_without_a_note_is_rejected():
    gate = _field(
        key="SEDI PDF",
        prompt_label="",
        target_kind="attachment",
        checklist="Attach the SEDI report.",
    )
    with pytest.raises(ValueError, match="no note to carry"):
        _spec(fields=(gate,))
    # With a note, the gate's bullet lands in it.
    spec = _spec(fields=(gate,), note=IntakeNote(header="Documents:"))
    assert "- Attach the SEDI report." in render_note(spec)
