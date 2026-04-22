"""
Bridge between pytest and scripts/check_invariants.py.

Runs the invariant script as a pytest so the full preflight gate is green
or red under a single command.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "scripts" / "check_invariants.py"


def test_invariants_script_passes():
    result = subprocess.run(
        [sys.executable, str(SCRIPT)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise AssertionError(
            "check_invariants.py reported hard failures:\n"
            + result.stdout
            + "\n---\n"
            + result.stderr
        )
