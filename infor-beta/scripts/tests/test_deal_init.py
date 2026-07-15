"""Unit tests for the deal_init helper."""

from pathlib import Path

import pytest

from deal_init import (
    DEAL_SUBDIRS,
    INIT_DIALOG_FIELDS,
    load_deal_context,
    load_or_locate_deal,
    render_init_dialogs,
    render_init_filings_note,
    render_init_prompt,
    save_deal_context,
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


def test_render_init_prompt_contains_five_questions():
    """The init prompt asks exactly five numbered items — verify all present."""
    prompt = render_init_prompt()
    for n in range(1, 6):
        assert f"{n}." in prompt, f"prompt missing item {n}"
    # The locked field names should appear verbatim
    for label in (
        "Deliverable type:",
        "Subject company name:",
        "Public or private?:",
        "Sector / industry:",
        "Filings / attachments:",
    ):
        assert label in prompt, f"prompt missing label {label!r}"
    # The codename is auto-derived, never asked — the prompt says so and no
    # longer carries the old items.
    assert "derived automatically" in prompt
    assert "Codename:" not in prompt
    assert "Anything else?:" not in prompt


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


def test_init_filings_question_is_a_status_gate():
    """Filings is a fixed status question — files themselves come via chat."""
    questions = [q for dialog in render_init_dialogs() for q in dialog]
    filings_q = next(q for q in questions if q["header"] == "Filings")
    labels = [opt["label"] for opt in filings_q["options"]]
    assert labels == [
        "Attached in this chat",
        "I'll drop them in my next message",
        "None for now",
    ]


def test_init_filings_note_keeps_the_ltm_reminder():
    # The filings checklist stays a text note (attachments can't come
    # through a dialog) and must keep the LTM-bridge explanation.
    note = render_init_filings_note()
    for token in ("LTM", "10-Q", "10-Ks", "five fiscal years"):
        assert token in note, f"filings note lost {token!r}"
    # The text-fallback G7 prompt must carry the same requirements.
    prompt = render_init_prompt()
    for token in ("LTM", "10-Q", "10-Ks", "five fiscal years"):
        assert token in prompt


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
