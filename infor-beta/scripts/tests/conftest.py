"""Pytest config — put `scripts/` on sys.path so tests can `import schemas`.

The repo root holds `pyproject.toml`, the plugin root holds `scripts/`, and the
tests live in `scripts/tests/`. From the tests directory, the package layout is:

    ../schemas/      ← `from schemas import Company, ...`
    ../codename.py   ← `import codename`

so we prepend the parent directory (`scripts/`) to sys.path here. This makes
`pytest infor-beta/scripts/tests/` work from any cwd without depending on
pyproject's `pythonpath` setting.
"""

from __future__ import annotations

import sys
from pathlib import Path

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))
