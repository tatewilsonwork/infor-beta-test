"""Unit tests for the deal_init helper."""

from pathlib import Path

import pytest

from deal_init import (
    DEAL_SUBDIRS,
    INIT_DEFAULT_FIELDS,
    INIT_DIALOG_FIELDS,
    INIT_INTAKE,
    load_deal_context,
    load_or_locate_deal,
    render_init_dialogs,
    render_init_prompt,
    save_deal_context,
)
from intake_spec import (
    ATTACHMENT_REQUIRED_HEADER,
    DEFAULTS_PROMPT_HEADER,
    render_attachment_request,
)
from schemas import Company, DealContext


def _ctx(tmp_root: Path, codename: str = "Project OpenText", **overrides) -> DealContext:
    kwargs = dict(
        codename=codename,
        deal_dir=tmp_root / codename,
        deliverable_type="earnings-update",
        subject_company=Company(legal_name="OpenText Corporation", ticker="OTEX"),
    )
    kwargs.update(overrides)
    return DealContext(**kwargs)


def test_render_init_prompt_contains_three_questions():
    """The init prompt asks exactly three numbered items — verify all present.

    Five until v0.5.50, four until v0.5.51. The filings item was the fifth, and
    it was a status question rather than a real one: it asked the analyst to say
    whether files were attached, which the deal's `filings/` directory already
    knew. Sector was the fourth, and its dialog already defaulted to "Infer from
    the web" — so both are now *listed* rather than asked: the filings as a
    bullet of the request this prompt ends with, the sector as a default in the
    override list.
    """
    prompt = render_init_prompt()
    for n in range(1, 4):
        assert f"{n}." in prompt, f"prompt missing item {n}"
    assert "4. " not in prompt, "the sector question is not a prompt item"
    assert "5. " not in prompt, "the filings status question is not a prompt item"
    # The locked field names should appear verbatim
    for label in (
        "Deliverable type:",
        "Subject company name:",
        "Public or private?:",
    ):
        assert label in prompt, f"prompt missing label {label!r}"
    # The sector is still in the prompt — as a default the analyst can override,
    # under the defaults header rather than as a numbered question.
    assert DEFAULTS_PROMPT_HEADER in prompt
    assert "- Sector / industry:" in prompt
    assert "Sector / industry (one line)?" not in prompt
    # The filings are still asked for, as an attachment the prompt requests.
    assert ATTACHMENT_REQUIRED_HEADER in prompt
    assert "Financial statements / filings" in prompt
    # The codename is auto-derived, never asked — the prompt says so and no
    # longer carries the old items.
    assert "derived automatically" in prompt
    assert "Codename:" not in prompt
    assert "Anything else?:" not in prompt


def test_init_sector_is_a_default_not_a_question():
    """Sector left the dialogs in v0.5.51 — that is what got a run to one call.

    Its dialog defaulted to "Infer from the web — I'll look it up and use it, no
    confirmation needed", so the question's entire content was an invitation to
    type the sector instead, which a reply to the defaults echo does. Keeping it
    would have landed a `/pitch` run at five questions, over the
    `AskUserQuestion` cap, and back to two dialogs.
    """
    for dialogs in (
        render_init_dialogs(),
        render_init_dialogs(include_deliverable=True),
    ):
        assert "Sector" not in [q["header"] for dialog in dialogs for q in dialog]
    assert "Sector" not in INIT_DIALOG_FIELDS
    # It targets the DealContext, not a plan input — so it must never appear in
    # a plan-input defaults table (see test_intake_spec).
    (sector,) = INIT_INTAKE.defaults
    assert sector.target_kind == "deal-context"
    assert sector.supplied, "the conductor researches the sector and sets it"
    assert sector.name == "subject_company.sector / subject_company.industry"
    assert INIT_DEFAULT_FIELDS == {sector.name: sector.rule}
    assert INIT_INTAKE.default_rules(supplied=True) == {}
    # Listing cannot absorb the freed slot instead: it is required, and a
    # required field may declare no default option.
    listing = next(f for f in INIT_INTAKE.fields if f.key == "Listing")
    assert listing.required and listing.default_option is None


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


def test_init_dialogs_are_valid_askuserquestion_payloads():
    _assert_askuserquestion_shape(render_init_dialogs())
    _assert_askuserquestion_shape(render_init_dialogs(include_deliverable=True))


def test_init_dialog_headers_match_field_table():
    headers = [
        q["header"]
        for dialog in render_init_dialogs(include_deliverable=True)
        for q in dialog
    ]
    assert len(headers) == len(set(headers)), "dialog headers must be unique"
    assert set(headers) == set(INIT_DIALOG_FIELDS)


def test_init_dialogs_deliverable_question_is_optional():
    without = [
        q["header"] for dialog in render_init_dialogs() for q in dialog
    ]
    with_it = [
        q["header"]
        for dialog in render_init_dialogs(include_deliverable=True)
        for q in dialog
    ]
    assert "Deliverable" not in without
    # Slash-command entry presets the deliverable — the other questions are
    # identical and keep their order either way.
    assert [h for h in with_it if h != "Deliverable"] == without
    # Deliverable leads: the codename is auto-derived, not a dialog question.
    assert with_it.index("Deliverable") == 0


def test_init_dialogs_render_verbatim_and_immutably():
    first = render_init_dialogs()
    first[0][0]["question"] = "mutated"
    first[0][0]["options"].clear()
    again = render_init_dialogs()
    assert again == render_init_dialogs()
    assert again[0][0]["question"] != "mutated"
    assert again[0][0]["options"]


def test_init_filings_is_an_attachment_not_a_question():
    """Filings is asked about in no rendering — it is a REQUIRED request bullet.

    Through v0.5.49 this was a three-option status dialog (attached / will drop
    next message / none for now), and "None for now" was its default — so the
    one question standing between a run and its only source of financial data
    defaulted to "no data". The filings are now a REQUIRED attachment, and the
    analyst answers by dropping the files into chat.
    """
    headers = [q["header"] for dialog in render_init_dialogs() for q in dialog]
    assert "Filings" not in headers
    headers_with_deliverable = [
        q["header"]
        for dialog in render_init_dialogs(include_deliverable=True)
        for q in dialog
    ]
    assert "Filings" not in headers_with_deliverable
    filings = next(f for f in INIT_INTAKE.attachment_fields() if f.key == "Filings")
    assert filings.required
    assert not filings.question and not filings.options and not filings.hint
    assert INIT_INTAKE.attachment_fields(required=True) == (filings,)
    # No plan input: every data stage reads the filings off <deal_dir>/filings/.
    assert INIT_INTAKE.attachment_inputs() == {}


def test_init_filings_request_keeps_the_ltm_reminder():
    # The filings requirement is one generated bullet now, and it must keep the
    # LTM-bridge explanation that used to sit in the note's prose.
    request = render_attachment_request(INIT_INTAKE)
    for token in ("LTM", "10-Q", "10-Ks", "five fiscal years"):
        assert token in request, f"filings request lost {token!r}"
    # The text-fallback G7 prompt must carry the same requirements.
    prompt = render_init_prompt()
    for token in ("LTM", "10-Q", "10-Ks", "five fiscal years"):
        assert token in prompt
    assert request.rstrip("\n") in prompt


def test_save_deal_context_bootstraps_dirs(tmp_path: Path):
    ctx = _ctx(tmp_path)
    save_deal_context(ctx)
    deal_dir = tmp_path / "Project OpenText"
    assert (deal_dir / "deal.json").is_file()
    for sub in DEAL_SUBDIRS:
        assert (deal_dir / sub).is_dir(), f"missing subdir {sub}"


def test_save_then_load_round_trip(tmp_path: Path):
    ctx = _ctx(tmp_path)
    save_deal_context(ctx)
    loaded = load_deal_context(tmp_path / "Project OpenText")
    assert loaded.codename == ctx.codename
    assert loaded.subject_company == ctx.subject_company


def test_load_missing_raises(tmp_path: Path):
    (tmp_path / "Project OpenText").mkdir()  # dir exists, no deal.json yet
    with pytest.raises(FileNotFoundError):
        load_deal_context(tmp_path / "Project OpenText")


def test_load_or_locate_existing_deal(tmp_path: Path):
    save_deal_context(_ctx(tmp_path))
    loaded, path = load_or_locate_deal("Project OpenText", deals_root=tmp_path)
    assert loaded is not None
    assert path == tmp_path / "Project OpenText"


def test_load_or_locate_case_insensitive(tmp_path: Path):
    save_deal_context(_ctx(tmp_path))
    loaded, path = load_or_locate_deal("project opentext", deals_root=tmp_path)
    assert loaded is not None
    assert path == tmp_path / "Project OpenText"


def test_load_or_locate_new_deal(tmp_path: Path):
    ctx, path = load_or_locate_deal("Project Brand New", deals_root=tmp_path)
    assert ctx is None
    assert path == tmp_path / "Project Brand New"
    # Must NOT mutate disk for a brand-new deal
    assert not path.exists()


def test_load_or_locate_dir_exists_no_deal_json(tmp_path: Path):
    """Empty directory shouldn't count as an existing deal."""
    (tmp_path / "Project Half-Built").mkdir()
    ctx, path = load_or_locate_deal("Project Half-Built", deals_root=tmp_path)
    assert ctx is None
    # Path should be the existing directory (case-insensitive match)
    assert path == tmp_path / "Project Half-Built"
