"""Drift locks on `CLAUDE.md`'s standing claims about the repo.

The import block came first (see below). The rest were added in v0.5.48, when the
migration wrap-up audited the brief against the repo and found three claims that
prose alone had let rot — the same failure mode, one level up: the brief is what a
fresh session reads *instead of* looking, so a stale claim there costs more than no
claim at all.

The block is the first thing a fresh dev session reads to learn what to import,
so a symbol that has been renamed or deleted is worse than no documentation: it
sends the reader to write an import that cannot work. That rotted once already —
by v0.5.42 the block still named eight symbols Phases B/C/D had removed
(`combine_workbooks`, `workbook_aggregator`, `CombineResult`, `excel_com_app`,
`_ClipboardPasteError`, `palatino_text_width_in`, `OVERVIEW_SLIDE_INDEX`,
`_KEEP_LIBRARY_INDICES`), because nothing executed it.

Now something does: this test extracts the block and runs it. A deleted export
fails here, in the release that deletes it, instead of in a future session that
trusted the brief.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[3]
BRIEF = REPO_ROOT / "CLAUDE.md"


def _python_blocks() -> list[str]:
    text = BRIEF.read_text(encoding="utf-8")
    return re.findall(r"```python\n(.*?)```", text, re.S)


def test_the_brief_has_exactly_one_python_block():
    # If a second one appears, extend this test to cover it rather than letting
    # the new block go unchecked.
    assert len(_python_blocks()) == 1, "CLAUDE.md gained a python block nothing verifies"


def test_every_symbol_in_the_import_block_resolves():
    source = _python_blocks()[0]
    # conftest has already put `scripts/` on sys.path, which is what the block's
    # own two bootstrap lines do for a skill; drop them so the test does not
    # depend on CLAUDE_PLUGIN_ROOT being set.
    body = "\n".join(
        line
        for line in source.splitlines()
        if not line.startswith(("import sys, os", "sys.path.insert"))
    )
    try:
        exec(compile(body, "CLAUDE.md import block", "exec"), {})
    except ImportError as exc:  # a renamed or deleted export
        pytest.fail(
            f"CLAUDE.md's shared-helpers import block is stale: {exc}. "
            f"Update the block in the same change that renames or removes the symbol."
        )


def test_the_brief_keeps_its_bootstrap_lines_importable_by_a_skill():
    # The block must stay copy-pasteable into a skill, which means keeping the
    # CLAUDE_PLUGIN_ROOT bootstrap the dispatched stages rely on.
    source = _python_blocks()[0]
    assert "CLAUDE_PLUGIN_ROOT" in source
    assert 'sys.path.insert(0, os.environ.get("CLAUDE_PLUGIN_ROOT"' in source


# ─── The stage portfolio ──────────────────────────────────────────────────────

PLUGIN_ROOT = REPO_ROOT / "infor-beta"


def _brief_stage_list(lead: str) -> set[str]:
    """The bare backticked names in the brief's paragraph starting with `lead`.

    Tokens holding a `/` or a `.` are paths, not stage names (`skills/<name>/SKILL.md`,
    `skills/`), so they are dropped — which keeps the prose free to name a path in
    the same breath as the list.
    """
    for line in BRIEF.read_text(encoding="utf-8").splitlines():
        if line.startswith(lead):
            return {
                token
                for token in re.findall(r"`([^`]+)`", line)
                if "/" not in token and "." not in token
            }
    raise AssertionError(f"CLAUDE.md no longer has a paragraph starting {lead!r}")


def test_contributor_brief_stage_lists_match_the_repo():
    """The brief's two exhaustive stage lists are checked against the repo.

    Phase F turned four dispatched skills into in-process transforms and deleted
    their SKILL.md files. The brief's portfolio section was updated for *that*, but
    the judgment list it left behind still named four skills which have never
    existed (`buyerslist`, `lbo-model`, `deck-writing`, `brand-guidelines`) as
    though they had been refactored, and omitted the two that do
    (`earningsupdate-content`, `pitch-content`). A reader trusting it would go
    hunting for four SKILL.md files and never learn about two real ones — the exact
    cost the section exists to avoid, which is why it is now a checked list rather
    than a described one.
    """
    import stage_transforms

    judgment = _brief_stage_list("**Judgment —")
    transforms = _brief_stage_list("**Transform —")

    # `conductor` is a skills/ directory but not a stage — it is the dispatcher.
    on_disk = {
        p.parent.name
        for p in (PLUGIN_ROOT / "skills").glob("*/SKILL.md")
    } - {"conductor"}

    assert judgment == on_disk, (
        f"CLAUDE.md's judgment list and skills/ disagree. "
        f"Only in the brief (no SKILL.md — a reader would hunt for one): "
        f"{sorted(judgment - on_disk)}. "
        f"Only on disk (undocumented): {sorted(on_disk - judgment)}."
    )
    assert transforms == set(stage_transforms.TRANSFORMS), (
        f"CLAUDE.md's transform list and stage_transforms.TRANSFORMS disagree: "
        f"brief-only {sorted(transforms - set(stage_transforms.TRANSFORMS))}, "
        f"registry-only {sorted(set(stage_transforms.TRANSFORMS) - transforms)}."
    )
    assert not judgment & transforms, (
        f"a stage is in both of the brief's lists: {sorted(judgment & transforms)}"
    )


# ─── The provenance obligation ───────────────────────────────────────────────
#
# Asserted across the stage list rather than file by file, because the failure it
# is here to stop is a stage being ADDED without one. A run's provenance.json held
# 70 records from three stages — `captable`, `financial-summary`, `ltm-metrics` —
# and zero from `content`, `comps`, `precedents` and `ownership`, which is exactly
# the set whose SKILL.md never mentioned provenance. The obligation existed only
# where it was written down, so the writing-down is what gets checked.

#: What declaring the obligation means, concretely. Each token is load-bearing:
#: the heading (so a reader finds it), the record type (a citation string is not a
#: source), the ledger (the stage's own), and the write target (`io.stage_dir` — a
#: shared file would be a read-modify-write race between concurrent wave-mates).
_PROVENANCE_TOKENS = (
    "Provenance — REQUIRED",
    "provenance.FigureSource",
    "ProvenanceLedger",
    "ledger.write(io.stage_dir)",
)

#: The two `skills/` directories that record no figures, each for a stated reason.
#: Anything else — including a stage added tomorrow — is in by default.
_NOT_RECORDING_STAGES = {
    "conductor": "the meta-skill that dispatches stages; it is not a stage",
    "deckcheck": "consumes the merged ledger and writes the run record; records no figures",
}


def _skill_docs() -> dict[str, str]:
    return {
        path.parent.name: path.read_text(encoding="utf-8")
        for path in (PLUGIN_ROOT / "skills").glob("*/SKILL.md")
    }


def test_every_data_producing_stage_declares_the_provenance_obligation():
    docs = _skill_docs()
    unknown = set(_NOT_RECORDING_STAGES) - set(docs)
    assert not unknown, f"the exemption list names skills that do not exist: {sorted(unknown)}"

    missing = {
        name: [token for token in _PROVENANCE_TOKENS if token not in text]
        for name, text in sorted(docs.items())
        if name not in _NOT_RECORDING_STAGES
    }
    missing = {name: tokens for name, tokens in missing.items() if tokens}
    assert not missing, (
        f"stage(s) whose SKILL.md does not declare the provenance obligation: {missing}. "
        f"Every figure a stage writes carries a FigureSource recorded in the stage's own "
        f"ProvenanceLedger and written to io.stage_dir. Mirror the wording in "
        f"financial-summary / ltm-metrics rather than inventing a second dialect — or, if "
        f"the stage genuinely records nothing, add it to _NOT_RECORDING_STAGES with the "
        f"reason."
    )
    assert len(missing) == 0 and len(docs) - len(_NOT_RECORDING_STAGES) >= 8, (
        "the obligation is asserted over fewer stages than the repo has — check the glob"
    )


def test_the_stage_that_consumes_the_ledger_declares_the_merge_instead():
    # `deckcheck` is exempt from recording, not from provenance: it is the stage that
    # merges every fragment into <run_dir>/provenance.json, and an exemption with
    # nothing behind it would quietly let that obligation go too.
    doc = _skill_docs()["deckcheck"]
    for token in ("write_run_provenance(io.run_dir)", "read_run_provenance(io.run_dir)"):
        assert token in doc, f"deckcheck's SKILL.md no longer declares {token}"


def test_no_stage_doc_writes_provenance_to_a_shared_file():
    # The race the fragments exist to prevent. A doc that told a sub-agent to write
    # the run-level record would be telling two concurrent wave-mates to overwrite
    # each other, and the loser's figures would simply be absent.
    offenders = {
        name: [
            line.strip()
            for line in text.splitlines()
            if ".write(io.run_dir" in line or "write_run_provenance" in line
        ]
        for name, text in _skill_docs().items()
        if name != "deckcheck"  # the one stage whose job IS the merge
    }
    offenders = {name: lines for name, lines in offenders.items() if lines}
    assert not offenders, (
        f"stage doc(s) writing provenance outside their own stage directory: {offenders}"
    )


# ─── The COM boundary ────────────────────────────────────────────────────────

_COM_TOKENS = ("win32com", "pythoncom", "DispatchEx", "EnsureDispatch")


def _com_users(root: Path) -> list[str]:
    this_file = Path(__file__).resolve()
    return sorted(
        f"{py.relative_to(REPO_ROOT).as_posix()}:{n}"
        for py in root.rglob("*.py")
        if py.resolve() != this_file  # the scanner spells the tokens it looks for
        for n, line in enumerate(py.read_text(encoding="utf-8").splitlines(), 1)
        if any(token in line for token in _COM_TOKENS)
    )


def test_the_shipped_plugin_holds_no_excel_com():
    """Phase D's deletion, and `tools/`'s exemption from it, in one assertion.

    Both halves are load-bearing and neither was checked. Production is Linux, so a
    COM path reintroduced under `infor-beta/` is code that cannot run where the
    analyst runs it — and it would run *green* on this Windows dev box, which is the
    inversion Phase A existed to end. In the other direction, `tools/` keeps COM on
    purpose (an add-in-free Excel assembles the deal-workbook template, and Excel is
    the repair oracle for the stamped defined names); a later sweep "finishing the
    job" would take working prep tooling with it. So the boundary is asserted, not
    described.
    """
    assert _com_users(PLUGIN_ROOT) == [], (
        f"Excel COM is back in the shipped plugin: {_com_users(PLUGIN_ROOT)}. "
        f"Production (Cowork) is Linux and has no Excel — render through LibreOffice "
        f"and write workbooks through openpyxl."
    )
    assert _com_users(REPO_ROOT / "tools"), (
        "tools/ no longer drives Excel COM. If that is deliberate, update the "
        "'Office on the Windows dev box' section of CLAUDE.md and the module "
        "docstrings in tools/ in the same change, then delete this assertion."
    )


# ─── One plugin version, three files ─────────────────────────────────────────


def test_the_three_version_files_agree():
    """The bump checklist, executed.

    `CLAUDE.md` names exactly three files that carry the version and calls updating
    all of them a release requirement. That was a hand-check through 47 releases;
    a missed file ships a plugin whose manifest and package disagree, and nothing
    would have said so.
    """
    import json
    import tomllib

    marketplace = json.loads(
        (REPO_ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8")
    )
    plugin = json.loads(
        (PLUGIN_ROOT / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8")
    )
    pyproject = tomllib.loads((REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    found = {
        ".claude-plugin/marketplace.json": marketplace["version"],
        "infor-beta/.claude-plugin/plugin.json": plugin["version"],
        "pyproject.toml": pyproject["project"]["version"],
    }
    assert len(set(found.values())) == 1, f"the three version files disagree: {found}"
