"""Tests for the bash helpers in `infor-beta/scripts/`.

Covers `find_template.sh` and `sanitize_name.sh`. Both are pure-shell
infrastructure used by every file-producing skill.

Run with:
    python -m unittest infor-beta/scripts/test_shell_helpers.py

Note (Phase 1): `infor-beta` ships no real templates yet, so the
`find_template.sh` tests use a tmpdir fixture by setting CLAUDE_PLUGIN_ROOT.
The corresponding `test_resolves_all_shipped_templates` regression test from
`infor-workflows` is deferred until templates land (Phase 3+).
"""

import os
import subprocess
import tempfile
import unittest
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parent
FIND_TEMPLATE = SCRIPTS_DIR / "find_template.sh"
SANITIZE_NAME = SCRIPTS_DIR / "sanitize_name.sh"


def _run(script_path: Path, *args: str, env: dict | None = None) -> subprocess.CompletedProcess:
    full_env = os.environ.copy()
    if env:
        full_env.update(env)
    return subprocess.run(
        ["bash", str(script_path), *args],
        capture_output=True,
        text=True,
        check=False,
        env=full_env,
    )


def _normalize_bash_path(p: str) -> str:
    """Convert a git-bash-style /c/Users/... path to native C:/Users/...
    so os.path.isfile resolves it on Windows. No-op on Unix paths."""
    if len(p) >= 3 and p[0] == "/" and p[1].isalpha() and p[2] == "/":
        return f"{p[1].upper()}:{p[2:]}"
    return p


# ─── find_template.sh ────────────────────────────────────────────────────────

class TestFindTemplate(unittest.TestCase):
    def test_resolves_template_under_claude_plugin_root(self):
        """Drop a fake template under a CLAUDE_PLUGIN_ROOT/templates dir and
        verify find_template.sh returns its absolute path."""
        with tempfile.TemporaryDirectory() as tmp:
            plugin_root = Path(tmp)
            (plugin_root / "templates").mkdir()
            fake = plugin_root / "templates" / "INFOR Comps Template.xlsx"
            fake.write_bytes(b"")

            result = _run(
                FIND_TEMPLATE,
                "INFOR Comps Template.xlsx",
                env={"CLAUDE_PLUGIN_ROOT": str(plugin_root)},
            )
            self.assertEqual(result.returncode, 0, msg=result.stderr)
            path = _normalize_bash_path(result.stdout.strip())
            self.assertTrue(
                os.path.isfile(path),
                f"find_template.sh returned a non-existent path: {path!r}",
            )
            self.assertTrue(path.endswith("INFOR Comps Template.xlsx"))

    def test_missing_template_exits_nonzero(self):
        result = _run(FIND_TEMPLATE, "Nonexistent Template.xlsx")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("not found", result.stderr.lower())

    def test_no_args_exits_with_usage(self):
        result = _run(FIND_TEMPLATE)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("usage", result.stderr.lower())

    def test_error_lists_candidate_paths(self):
        """When the template is missing, the error should list every directory
        that was searched — otherwise debugging install issues is impossible."""
        result = _run(FIND_TEMPLATE, "Bogus.xlsx")
        self.assertIn("Searched:", result.stderr)
        self.assertIn("templates", result.stderr)


# ─── sanitize_name.sh ────────────────────────────────────────────────────────

class TestSanitizeName(unittest.TestCase):
    def _sanitize(self, value: str) -> str:
        result = _run(SANITIZE_NAME, value)
        self.assertEqual(
            result.returncode,
            0,
            f"sanitize_name.sh failed on {value!r}: {result.stderr}",
        )
        return result.stdout.strip()

    def test_company_name_with_spaces(self):
        self.assertEqual(
            self._sanitize("Rogers Communications Inc."),
            "Rogers-Communications-Inc",
        )

    def test_capiq_ticker_keeps_case_drops_colon(self):
        # The Nasdaq tier prefix is case-sensitive in CapIQ — must preserve it
        self.assertEqual(self._sanitize("NasdaqGS:MSFT"), "NasdaqGS-MSFT")
        self.assertEqual(self._sanitize("TSX:RY"), "TSX-RY")

    def test_ampersand_removed(self):
        self.assertEqual(self._sanitize("Dye & Durham"), "Dye-Durham")
        self.assertEqual(self._sanitize("AT&T Inc."), "AT-T-Inc")

    def test_leading_and_trailing_whitespace_stripped(self):
        self.assertEqual(self._sanitize("  spaces  "), "spaces")

    def test_consecutive_specials_collapse(self):
        self.assertEqual(self._sanitize("a -- b --- c"), "a-b-c")

    def test_period_and_apostrophe_dropped(self):
        self.assertEqual(self._sanitize("Macy's Inc."), "Macy-s-Inc")

    def test_unicode_dropped_to_hyphen(self):
        # Non-ASCII chars are non-alphanumeric per the [A-Za-z0-9] rule,
        # so they collapse to a hyphen — predictable for filenames
        self.assertEqual(self._sanitize("Café Co"), "Caf-Co")

    def test_empty_input_exits_with_usage(self):
        result = _run(SANITIZE_NAME)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("usage", result.stderr.lower())


if __name__ == "__main__":
    unittest.main(verbosity=2)
