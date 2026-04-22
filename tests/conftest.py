"""
Shared pytest configuration.

Adds repo root and app/ to sys.path so tests can import both
`app.models.models` style and direct `models.models` style
(the app uses both depending on branch).
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(REPO_ROOT / "app"))
