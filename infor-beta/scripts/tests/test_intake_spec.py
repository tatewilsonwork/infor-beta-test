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
`*_DIALOG_PLAN_INPUTS` answer mapping, and the attachments — which as of
v0.5.50 are asked about in no rendering at all, and reach the analyst as the
two-section request instead.

And, as of v0.5.51, the property the whole intake exists to have: **a
slash-command run renders exactly ONE `AskUserQuestion` call.** Three renderers
merge deal-init's spec with the deliverable's now (dialogs, attachment request,
defaults echo), and the merged dialog fits in one call only while the run's
question count stays within `DIALOG_MAX_QUESTIONS`. Nothing raises when a fifth
question is added — it silently splits the run back into two dialogs — so
`test_a_slash_command_run_renders_exactly_one_dialog` is the guard, and it is
deliberately assertive about the exact headers rather than just the count.
"""

from __future__ import annotations

import pytest
import yaml
from pathlib import Path

import deal_init
import deck_spec
from intake_spec import (
    ATTACHMENT_OPTIONAL_HEADER,
    ATTACHMENT_REQUIRED_HEADER,
    DEFAULTS_ECHO_HEADER,
    DIALOG_MAX_QUESTIONS,
    IntakeDefault,
    IntakeField,
    IntakeOption,
    IntakeSpec,
    render_attachment_request,
    render_defaults_echo,
    render_dialogs,
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

# Every computed value any spec's echo templates can ask for. Passed whole so a
# per-spec test does not have to know which subset that spec needs — the
# renderer's own missing-value check is exercised separately.
ECHO_VALUES = {
    "sector": "Enterprise software",
    "client_name": "Example Target Inc.",
    "presentation_date": "July 2026",
    "reporting_quarter": "Q2 2026",
    "comparison_quarter": "Q2 2025",
}


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
    assert deal_init.render_init_dialogs(include_deliverable=True) == render_dialogs(
        spec
    )
    assert deal_init.render_init_dialogs() == render_dialogs(
        spec, omit=(deal_init.DELIVERABLE_FIELD_KEY,)
    )
    assert deal_init.INIT_DIALOG_FIELDS == spec.targets("deal-context")
    assert deal_init.INIT_DEFAULT_FIELDS == spec.default_rules(
        supplied=True, target_kind="deal-context"
    )


def test_the_run_renderings_are_generated_from_the_two_specs_merged():
    """Every merged renderer is `<intake_spec generator>(INIT_INTAKE, <deck spec>)`.

    The three live surfaces a run puts in front of the analyst — the dialog, the
    attachment request, the defaults echo — share one shape: varargs over the
    specs, merged in spec order. A hand-assembled second rendering of any of
    them fails here, in the change that adds it.
    """
    for deliverable, spec in (
        ("pitch", deck_spec.PITCH_INTAKE),
        ("earnings-update", deck_spec.EARNINGS_UPDATE_INTAKE),
    ):
        specs = (deal_init.INIT_INTAKE, spec)
        assert deck_spec.render_run_dialogs(deliverable) == render_dialogs(
            *specs, omit=(deal_init.DELIVERABLE_FIELD_KEY,)
        )
        assert deck_spec.render_run_dialogs(
            deliverable, include_deliverable=True
        ) == render_dialogs(*specs)
        assert deck_spec.render_run_attachment_request(
            deliverable
        ) == render_attachment_request(*specs)
        assert deck_spec.render_run_defaults(
            deliverable, **ECHO_VALUES
        ) == render_defaults_echo(*specs, values=ECHO_VALUES)
    # A deliverable with no questionnaire degrades to deal-init's half on all
    # three, identically — the `overview` stub still has a listing to ask about,
    # a sector to default, and filings to request.
    init_only = (deal_init.INIT_INTAKE,)
    for deliverable in (None, "overview", "one-off-skill"):
        assert deck_spec.render_run_dialogs(deliverable) == render_dialogs(
            *init_only, omit=(deal_init.DELIVERABLE_FIELD_KEY,)
        )
        assert deck_spec.render_run_attachment_request(
            deliverable
        ) == render_attachment_request(*init_only)
        assert deck_spec.render_run_defaults(
            deliverable, **ECHO_VALUES
        ) == render_defaults_echo(*init_only, values=ECHO_VALUES)
    with pytest.raises(ValueError, match="unknown deliverable type"):
        deck_spec.render_run_dialogs("pitchh")
    with pytest.raises(ValueError, match="unknown deliverable type"):
        deck_spec.render_run_defaults("pitchh")


@pytest.mark.parametrize(
    "deliverable,spec",
    [("pitch", deck_spec.PITCH_INTAKE), ("earnings-update", deck_spec.EARNINGS_UPDATE_INTAKE)],
)
def test_deck_spec_renderings_are_generated(deliverable: str, spec: IntakeSpec):
    assert deck_spec.render_deck_spec_prompt(deliverable) == render_prompt(spec)
    assert deck_spec.render_deck_spec_dialogs(deliverable) == render_dialogs(
        spec, target_kinds=("plan-input",)
    )
    # The merged renderings are asserted in
    # `test_the_run_renderings_are_generated_from_the_two_specs_merged`.


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_generation_is_deterministic(label: str, spec: IntakeSpec):
    """Same spec in, same bytes out — the locked questionnaire, mechanised."""
    assert render_prompt(spec) == render_prompt(spec)
    assert render_dialogs(spec) == render_dialogs(spec)
    assert render_attachment_request(spec) == render_attachment_request(spec)
    # Fresh payloads every call: a caller mutating one render cannot reach the
    # next, so no deep-copy discipline is needed at the call sites.
    first = render_dialogs(spec)
    if not first:
        # One spec has no questions left at all: the earnings-update deck spec
        # defaults both quarters and takes the EEO snip as an attachment.
        assert spec is deck_spec.EARNINGS_UPDATE_INTAKE
        return
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

    # Same items, same order. Attachments are in neither list — nothing is
    # asked about them (asserted separately below).
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


@pytest.mark.parametrize(
    "deliverable,spec",
    [
        ("pitch", deck_spec.PITCH_INTAKE),
        ("earnings-update", deck_spec.EARNINGS_UPDATE_INTAKE),
    ],
)
def test_the_fallback_is_two_prompts_asking_the_merged_dialog_in_order(
    deliverable: str, spec: IntakeSpec
):
    """The dialogs merge; the text fallback stays TWO prompts. This says which.

    `render_prompt` takes one spec deliberately. A merged prompt would have to
    renumber items across specs, and the numbers are the fallback's answer
    mapping (`*_ITEM_PLAN_INPUTS`), so a merged one would key deal-init's items
    off numbers the deliverable's table also claims. Instead the analyst reads
    `render_init_prompt()` then `render_deck_spec_prompt(...)` — each numbering
    its own half, and between them asking exactly the merged dialog's questions
    in the merged dialog's order.
    """
    merged = [
        q["header"] for dialog in deck_spec.render_run_dialogs(deliverable) for q in dialog
    ]
    halves = [
        f.key
        for one in (deal_init.INIT_INTAKE, spec)
        for f in one.dialog_fields(omit=(deal_init.DELIVERABLE_FIELD_KEY,))
    ]
    assert halves == merged
    # Each half's own prompt numbers its own items, from 1 — and both prompts
    # together carry every question the one dialog asks.
    for one in (deal_init.INIT_INTAKE, spec):
        prompt = _flat(render_prompt(one))
        for number, f in enumerate(one.prompt_fields(), start=1):
            assert f"{number}. {f.prompt_label}:" in prompt
        for f in one.dialog_fields():
            if f.key in merged:
                assert _flat(f.question) in prompt


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
    echo = render_defaults_echo(spec, values=ECHO_VALUES)
    for d in spec.defaults:
        if not d.echoed:
            continue
        assert f"- {d.echo_line_label}:" in echo
        if d.echo is None:
            assert d.rule in echo


def test_defaults_echo_names_every_missing_computed_value():
    with pytest.raises(ValueError, match="client_name, presentation_date"):
        render_defaults_echo(
            deck_spec.PITCH_INTAKE, values={"reporting_quarter": "Q2 2026"}
        )
    # Merged, deal-init's own computed value is named the same way — its sector
    # is the first default the G7 spec ever had, so this is the first run where a
    # missing deal-init value could have echoed as a blank.
    with pytest.raises(ValueError, match="deal-init \\+ pitch defaults echo needs sector"):
        deck_spec.render_run_defaults(
            "pitch",
            client_name="Example Target Inc.",
            presentation_date="July 2026",
            reporting_quarter="Q2 2026",
            comparison_quarter="Q2 2025",
        )


def test_the_merged_defaults_echo_is_one_message_and_refuses_a_repeat():
    """One echo per run, deal-init's defaults first, and no target twice."""
    echo = deck_spec.render_run_defaults("pitch", **ECHO_VALUES)
    assert echo.count(DEFAULTS_ECHO_HEADER) == 1, "two echoes merged into one message"
    flat = _flat(echo)
    # Spec order: deal-init's sector leads the deliverable's client name.
    assert flat.index("Sector / industry:") < flat.index("Client name on the cover:")
    for label in ("Valuation range:", "Risk notes:"):
        assert label in flat, f"the demoted {label!r} default is not echoed"
    # A default's `name` IS its target, so two entries writing one target cannot
    # both be in effect — the echo refuses rather than listing it twice.
    with pytest.raises(ValueError, match="one target, one default"):
        render_defaults_echo(
            deck_spec.PITCH_INTAKE, deck_spec.PITCH_INTAKE, values=ECHO_VALUES
        )


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_the_request_describes_every_attachment_and_asks_about_none(
    label: str, spec: IntakeSpec
):
    """Every attachment is one request bullet, and a question in no rendering.

    This is the v0.5.50 guarantee. An attachment used to be a status dialog
    (attached / will-drop-next-message / none) whose option descriptions carried
    the "without it, X stays a placeholder" warning. There is no dialog to hold
    that any more, so the warning lives in the `checklist` line and the request
    is where the analyst reads it.
    """
    request = _flat(render_attachment_request(spec))
    prompt = _flat(render_prompt(spec))
    attachments = spec.attachment_fields()
    assert attachments, f"{label} declares no attachments"
    assert request in prompt, "the text prompt must embed the request"
    for f in attachments:
        assert _flat(f.checklist) in request, (
            f"{label} request does not describe the {f.key!r} attachment"
        )
        assert _flat(f.prompt_label) in request
        # Asked about nowhere: no dialog question, no numbered prompt item, and
        # no entry in any answer-mapping table.
        assert not f.is_dialog and not f.is_free_text and f.is_attachment
        assert not f.question and not f.options and not f.hint
        assert f.key not in spec.targets("plan-input")
        assert f.key not in spec.targets("deal-context")
        assert f not in spec.prompt_fields()


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_the_request_splits_required_from_optional(label: str, spec: IntakeSpec):
    """Both sections come from `required`, so neither can list the wrong file."""
    request = render_attachment_request(spec)
    required = spec.attachment_fields(required=True)
    optional = spec.attachment_fields(required=False)
    assert set(required) | set(optional) == set(spec.attachment_fields())
    assert not set(required) & set(optional)
    for header, section in (
        (ATTACHMENT_REQUIRED_HEADER, required),
        (ATTACHMENT_OPTIONAL_HEADER, optional),
    ):
        # A section header appears exactly when that section has a bullet.
        assert (header in request) is bool(section)
    # Requiredness orders the message, not just labels it.
    if required and optional:
        assert request.index(ATTACHMENT_REQUIRED_HEADER) < request.index(
            ATTACHMENT_OPTIONAL_HEADER
        )
    for f in required:
        assert _flat(f.checklist) in _flat(
            request.split(ATTACHMENT_OPTIONAL_HEADER)[0]
        )


def test_the_request_merges_the_specs_in_order_and_refuses_a_repeat():
    """One message per run: deal-init's filings, then the deliverable's own."""
    merged = deck_spec.render_run_attachment_request("pitch")
    for spec in (deal_init.INIT_INTAKE, deck_spec.PITCH_INTAKE):
        for f in spec.attachment_fields():
            assert _flat(f.checklist) in _flat(merged)
    # deal-init leads, so the G7 filings are the first REQUIRED bullet.
    filings = deal_init.INIT_INTAKE.attachment_fields()[0]
    cim = deck_spec.PITCH_INTAKE.attachment_fields()[0]
    assert _flat(merged).index(_flat(filings.prompt_label)) < _flat(merged).index(
        _flat(cim.prompt_label)
    )
    # The G7 filings were described in deal-init's prose AND again, in different
    # words, in each deck spec's checklist. Merging from the declarations is what
    # removed that; a re-declaration is refused rather than listed twice.
    with pytest.raises(ValueError, match="one document, one declaration"):
        render_attachment_request(deal_init.INIT_INTAKE, deal_init.INIT_INTAKE)


def test_an_attachment_carrying_a_plan_input_matches_the_plan():
    """The CIM and the EEO snip are attachments whose saved path IS a plan input."""
    assert deck_spec.PITCH_ATTACHMENT_PLAN_INPUTS == {
        "cim_path": "CIM / management presentation"
    }
    assert deck_spec.EARNINGS_UPDATE_ATTACHMENT_PLAN_INPUTS == {
        "eeo_snip_path": "Bloomberg EEO snip"
    }
    # A REQUIRED one halts the run when it never arrives, so requiredness has to
    # be the plan's — checked in full by
    # `test_declared_requiredness_matches_the_plans` below.
    assert deck_spec.EARNINGS_UPDATE_INTAKE.attachment_inputs(required=True) == {
        "eeo_snip_path": "Bloomberg EEO snip"
    }
    assert deck_spec.PITCH_INTAKE.attachment_inputs(required=True) == {}
    # The SEDI report and the Bloomberg export carry no path: the ownership
    # stage finds them under <deal_dir>/filings/ itself.
    assert set(deck_spec.PITCH_INTAKE.attachment_inputs()) == {"cim_path"}


@pytest.mark.parametrize("label,spec", SPECS, ids=SPEC_IDS)
def test_attachments_are_not_reachable_through_the_dialog_tables(
    label: str, spec: IntakeSpec
):
    """`targets("attachment")` used to answer `{}` — it now refuses.

    `IntakeSpec.targets` filters on `is_dialog`, so once attachments stopped
    being dialogs it would have returned an empty table to every caller asking
    for them, silently. That is how `PITCH_DOCUMENTS_DIALOG_TARGETS` would have
    survived this change as `{}` and nothing would have failed.
    """
    with pytest.raises(ValueError, match="attachments are not dialog questions"):
        spec.targets("attachment")
    with pytest.raises(ValueError, match="attachments are not numbered"):
        spec.item_targets("attachment")


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
        # Attachments share the header namespace but reach no answer table.
        assert not {f.key for f in spec.attachment_fields()} & set(table)
        # A plan input is collected exactly one way — asked, attached, or
        # defaulted, never two of them.
        assert not set(table.values()) & set(spec.attachment_inputs())
        assert not set(table.values()) & {d.name for d in spec.defaults}


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
    """`required` means the plan input has no default — for both ways in.

    An asked question and an attached document supply a plan input alike, so both
    are checked against the plan. This is what stops the earnings-update spec
    from listing the EEO snip as OPTIONAL while `earnings-update.yaml` declares
    `eeo_snip_path` required — a run that would then resolve the reference to
    None and hand a content stage no estimates source.
    """
    for plan_name, spec in (
        ("pitch.yaml", deck_spec.PITCH_INTAKE),
        ("earnings-update.yaml", deck_spec.EARNINGS_UPDATE_INTAKE),
    ):
        plan_required = {s.name: s.required for s in _plan(plan_name).plan_inputs}
        checked = 0
        for f in spec.fields:
            if f.is_attachment:
                if not f.plan_input:
                    continue  # discovered on disk; not a plan input at all
                name, how = f.plan_input, "attaches"
            elif f.target_kind == "plan-input":
                name, how = f.target, "asks for"
            else:
                continue
            checked += 1
            assert name in plan_required, f"{f.key} {how} an unknown input {name!r}"
            assert f.required is plan_required[name], (
                f"{f.key} is {'REQUIRED' if f.required else 'optional'} in the "
                f"questionnaire but the opposite in {plan_name}"
            )
        assert checked, f"{plan_name} questionnaire supplies no plan input at all"


def test_dialog_batches_are_spec_order_chunked_at_the_cap():
    """The batching rule, restated for v0.5.51: spec order, chunked at the cap.

    This replaced `test_dialog_batches_respect_the_declared_groups`. Fields used
    to be batched by a declared `group` and only *then* chunked, which is what
    forced the pitch questionnaire's content and slide-mix sections into two
    dialogs and a `/pitch` run into three calls. `group` is deleted: nothing
    reads a section label now, so questions merge across specs in declaration
    order and split only when they hit `DIALOG_MAX_QUESTIONS`.
    """
    for label, spec in SPECS:
        fields = spec.dialog_fields()
        batched = [[q["header"] for q in dialog] for dialog in render_dialogs(spec)]
        # Flattening the calls returns the declaration order, unpermuted.
        assert [h for dialog in batched for h in dialog] == [f.key for f in fields]
        # Every call is full but the last — that is what "chunked" means, and it
        # is the property a group-aware batcher did NOT have.
        for dialog in batched[:-1]:
            assert len(dialog) == DIALOG_MAX_QUESTIONS, f"{label} left a call short"
        for dialog in batched:
            assert 1 <= len(dialog) <= DIALOG_MAX_QUESTIONS
    # The merged run, chunked the same way: one call, because there are four
    # questions. The deleted section labels are still in deck_spec.py as comments
    # — they were always the documentation, which is why the field cost nothing.
    assert [
        [q["header"] for q in dialog]
        for dialog in deck_spec.render_run_dialogs("pitch")
    ] == [["Listing", "Notes", "Targets", "Highlights"]]
    # Over the cap, it splits rather than raising — hence the test below.
    over = IntakeSpec(
        name="probe",
        title="Probe",
        fields=tuple(_field(key=f"Q{i}", target=f"t{i}") for i in range(6)),
    )
    assert [len(dialog) for dialog in render_dialogs(over)] == [4, 2]
    assert [f.key for f in deck_spec.PITCH_INTAKE.attachment_fields()] == [
        "CIM",
        "SEDI PDF",
        "BBG export",
    ]


def test_a_slash_command_run_renders_exactly_one_dialog():
    """The property this change exists to create — one dialog per `/command` run.

    `/pitch` walked the analyst through THREE sequential `AskUserQuestion` calls
    and seven questions through v0.5.50: deal-init's `[Listing, Sector]`, then the
    deck spec's `[Notes, Valuation, Risk notes]` and `[Targets, Highlights]`. It
    is now one call of four, which is the tool's own per-call maximum — so this
    has no headroom by construction, and that is deliberate. A fifth question
    does not raise anywhere: `render_dialogs` chunks it into a silent second
    call. **This test is the only thing that fails.** When it does, the fix is
    to default the new item and echo it, not to accept two dialogs.
    """
    expected = {
        "pitch": ["Listing", "Notes", "Targets", "Highlights"],
        "earnings-update": ["Listing"],
    }
    for deliverable, headers in expected.items():
        dialogs = deck_spec.render_run_dialogs(deliverable)
        assert len(dialogs) == 1, (
            f"/{deliverable} renders {len(dialogs)} AskUserQuestion calls, not 1 — "
            f"a question was added; default it and echo it instead"
        )
        assert [q["header"] for q in dialogs[0]] == headers
        assert len(dialogs[0]) <= DIALOG_MAX_QUESTIONS
    # Generic conductor entry is two rounds and cannot be one: the Deliverable
    # answer decides which deck spec exists, so there is nothing to merge until
    # it is answered. Documented in the conductor SKILL.md, not designed around.
    assert len(deal_init.render_init_dialogs(include_deliverable=True)) == 1
    assert len(deck_spec.render_deck_spec_dialogs("pitch")) == 1


def test_the_merged_dialog_refuses_a_repeated_header():
    """The replacement for the split-group rejection, and a real cross-spec case.

    `IntakeSpec.__post_init__` already rejects a duplicate key *within* a spec,
    so the merge's own hazard is a collision *between* deal-init and a deck spec
    — one question's answer routed through another's target inside a single call.
    Mirrors `test_the_request_merges_the_specs_in_order_and_refuses_a_repeat`.
    """
    # A deck spec that re-asked one of deal-init's questions by header.
    colliding = IntakeSpec(
        name="probe",
        title="Probe",
        fields=(_field(key="Listing", target="analyst_notes"),),
    )
    with pytest.raises(ValueError, match="one question, one declaration"):
        render_dialogs(deal_init.INIT_INTAKE, colliding)
    # And the same spec twice, which is how a wrapper double-passing deal-init
    # would show up.
    with pytest.raises(ValueError, match="one question, one declaration"):
        render_dialogs(deal_init.INIT_INTAKE, deal_init.INIT_INTAKE)


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


def _attachment(**overrides) -> IntakeField:
    kwargs = dict(
        key="SEDI PDF",
        prompt_label="SEDI report",
        target="consumed by the ownership stage",
        target_kind="attachment",
        checklist="Canadian targets only; without it the insider side is blank.",
        question="",
        options=(),
    )
    kwargs.update(overrides)
    return IntakeField(**kwargs)


def test_an_attachment_needs_a_checklist_line():
    with pytest.raises(ValueError, match="needs a checklist line"):
        _attachment(checklist="")


def test_only_an_attachment_contributes_a_checklist_bullet():
    with pytest.raises(ValueError, match="not an attachment"):
        _field(checklist="Attach the thing.")


def test_an_attachment_is_its_own_kind_not_a_dialog_field_about_a_file():
    """The shape v0.5.50 made unrepresentable: an attachment that is asked about.

    Before, `attachment` was a rider on the dialog/free-text split — an
    attachment with no options fell through to the free-text branch and was told
    it needed a `hint`, so every one of them had to carry dialog wording to be
    constructible at all. Now each of the three kinds is validated as itself.
    """
    for wording in (
        {"question": "Attach the SEDI report?"},
        {"hint": "drop the PDF in chat"},
        {
            "question": "Attached?",
            "options": (IntakeOption("Yes", "It is."), IntakeOption("No", "Not yet.")),
        },
    ):
        with pytest.raises(ValueError, match="carries dialog wording"):
            _attachment(**wording)
    # And the shape that is now legal: no wording at all, just the bullet.
    field = _attachment()
    assert field.is_attachment
    assert not field.is_dialog and not field.is_free_text


def test_only_an_attachment_may_carry_a_plan_input():
    """A question's plan input is its `target`; `plan_input` is a file's path."""
    with pytest.raises(ValueError, match="is not an attachment"):
        _field(plan_input="cim_path")
    assert _attachment(plan_input="cim_path").plan_input == "cim_path"


def test_every_field_needs_a_prompt_label():
    """Including an attachment: the label is the handle its bullet leads with."""
    with pytest.raises(ValueError, match="needs a prompt_label"):
        _attachment(prompt_label="")
    with pytest.raises(ValueError, match="needs a prompt_label"):
        _field(prompt_label="")


def _spec(**overrides) -> IntakeSpec:
    kwargs = dict(name="probe", title="Probe", fields=(_field(),))
    kwargs.update(overrides)
    return IntakeSpec(**kwargs)


def test_duplicate_dialog_headers_are_rejected():
    """A reused header would route one question's answer through another's target."""
    with pytest.raises(ValueError, match="reuses dialog header"):
        _spec(fields=(_field(), _field(target="risk_notes")))


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


def _sector_default(**overrides) -> IntakeDefault:
    kwargs = dict(
        name="subject_company.sector / subject_company.industry",
        label="Sector / industry",
        rule="researched on the web",
        supplied=True,
        target_kind="deal-context",
    )
    kwargs.update(overrides)
    return IntakeDefault(**kwargs)


def test_a_deal_context_field_cannot_also_be_defaulted():
    """The check that made demoting Sector safe (v0.5.51).

    `IntakeDefault.name` IS the target, so the clash check is plain equality
    against the field's `target` — which is why the demotion needed no new
    bookkeeping. The check runs per kind, because a `plan-input` name and a
    `DealContext` path are different namespaces.
    """
    asked_sector = _field(
        key="Sector",
        prompt_label="Sector / industry",
        target="subject_company.sector / subject_company.industry",
        target_kind="deal-context",
        question="Sector?",
    )
    with pytest.raises(ValueError, match="both asks for and defaults"):
        _spec(fields=(_field(), asked_sector), defaults=(_sector_default(),))
    # Defaulted alone is the shipped shape; asked alone is the pre-v0.5.51 one.
    _spec(fields=(_field(),), defaults=(_sector_default(),))
    _spec(fields=(_field(), asked_sector))
    # A same-string default of the OTHER kind is not a clash — different
    # namespaces, so the per-kind split has to be real.
    _spec(
        fields=(_field(), asked_sector),
        defaults=(_sector_default(target_kind="plan-input"),),
    )


def test_a_deal_context_default_never_reaches_the_plan_input_table():
    """`default_rules` splits by kind, or the conductor sets a bogus plan input.

    `*_DEFAULT_SUPPLIED_INPUTS` is what the conductor turns into `plan_inputs`
    entries. A deal-context default leaking into one would have it write a plan
    input literally named `subject_company.sector / subject_company.industry`,
    which no plan declares — so the split is load-bearing, not tidiness.
    """
    spec = _spec(
        defaults=(
            _sector_default(),
            IntakeDefault(
                name="client_name",
                label="Client name",
                rule="the subject company",
                supplied=True,
            ),
        )
    )
    assert spec.default_rules(supplied=True) == {"client_name": "the subject company"}
    assert spec.default_rules(supplied=True, target_kind="deal-context") == {
        "subject_company.sector / subject_company.industry": "researched on the web"
    }
    # The shipped tables, checked the same way: deal-init's one default is
    # deal-context and appears in NO plan-input table.
    sector = next(iter(deal_init.INIT_DEFAULT_FIELDS))
    assert deal_init.INIT_INTAKE.default_rules(supplied=True) == {}
    assert deal_init.INIT_INTAKE.default_rules(supplied=False) == {}
    for table in (
        deck_spec.PITCH_DEFAULT_SUPPLIED_INPUTS,
        deck_spec.PITCH_DEFAULT_UNSET_INPUTS,
        deck_spec.EARNINGS_UPDATE_DEFAULT_SUPPLIED_INPUTS,
        deck_spec.EARNINGS_UPDATE_DEFAULT_UNSET_INPUTS,
    ):
        assert sector not in table
    # And a deliverable spec declares no deal-context default at all, so the
    # DealContext half belongs to deal-init alone.
    for spec in (deck_spec.PITCH_INTAKE, deck_spec.EARNINGS_UPDATE_INTAKE):
        assert spec.default_rules(supplied=True, target_kind="deal-context") == {}
        assert spec.default_rules(supplied=False, target_kind="deal-context") == {}
    with pytest.raises(ValueError, match="attachments have no defaults"):
        deal_init.INIT_INTAKE.default_rules(supplied=True, target_kind="attachment")


def test_an_attachment_cannot_be_defaulted():
    """A file has no default value — a missing OPTIONAL one is an absent path."""
    with pytest.raises(ValueError, match="targets an attachment"):
        _sector_default(target_kind="attachment")
    with pytest.raises(ValueError, match="known:"):
        _sector_default(target_kind="deal_context")


def test_an_attachment_needs_no_note_object_to_be_rendered():
    """The request's framing is code-owned; a spec declares only the bullets.

    Attachments used to require an `IntakeNote` on the spec to carry their
    checklist lines, and each deck spec declared one with the same boilerplate
    header. A merged request cannot honour two headers, and a note's free-prose
    `bullets` had no required/optional section to go in, so the object went.
    """
    spec = _spec(fields=(_attachment(checklist="Attach the SEDI report."),))
    request = render_attachment_request(spec)
    assert "- SEDI report — Attach the SEDI report." in request
    assert ATTACHMENT_OPTIONAL_HEADER in request
    assert ATTACHMENT_REQUIRED_HEADER not in request


def test_a_spec_with_no_attachments_renders_no_request():
    assert render_attachment_request(_spec()) == ""
    # And the text prompt then has no request section to embed.
    assert ATTACHMENT_OPTIONAL_HEADER not in render_prompt(_spec())


def test_a_plan_input_cannot_be_both_asked_and_attached():
    with pytest.raises(ValueError, match="collects plan input"):
        _spec(
            fields=(
                _field(target="cim_path"),
                _attachment(key="CIM", plan_input="cim_path"),
            )
        )
