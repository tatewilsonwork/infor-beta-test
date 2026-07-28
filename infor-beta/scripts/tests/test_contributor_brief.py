"""Drift lock on `CLAUDE.md`'s shared-helpers import block.

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
