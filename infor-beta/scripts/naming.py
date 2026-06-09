"""Shared filename sanitisation.

Kept dependency-free (stdlib ``re`` only) so the openpyxl-only workbook builders
(`ltm_metrics`, `workbook_aggregator`) and the python-pptx deck assemblers can all
import it without pulling in unrelated dependencies.

This replaces the per-module ``_safe_name`` / ``_safe_file_stem`` copies that had
drifted into four identical definitions.
"""

from __future__ import annotations

import re

_UNSAFE_FILENAME_CHARS = re.compile(r'[/\\:*?"<>|]+')


def safe_filename(value: str, *, default: str = "Output") -> str:
    """Make ``value`` safe to use as a file-name stem.

    Path-unsafe characters become a single ``-`` (a separator — unlike
    ``codename._strip_unsafe``, which *deletes* them), internal whitespace runs
    collapse to one space, and an empty result falls back to ``default``. So
    ``"NasdaqGS:MSFT"`` -> ``"NasdaqGS-MSFT"``. Mirrors ``sanitize_name.sh``.
    """
    safe = _UNSAFE_FILENAME_CHARS.sub("-", value).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe or default
