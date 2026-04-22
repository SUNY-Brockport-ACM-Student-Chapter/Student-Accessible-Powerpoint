#!/usr/bin/env python3
"""
preflight.py — run before opening a PR.

Aggregates:
    1. scripts/doctor.py            (environment)
    2. scripts/check_invariants.py  (static invariants)
    3. pytest -q                    (unit tests, if any)
    4. import-smoke of app modules  (catches syntax errors / missing deps)

Any hard failure -> exit code 1.
Warnings -> exit code 0 but flagged.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = REPO_ROOT / "scripts"


def run_step(title: str, argv: list[str], *, cwd: Path = REPO_ROOT, required: bool = True) -> int:
    bar = "=" * 72
    print(f"\n{bar}\n[preflight] {title}\n{bar}")
    try:
        proc = subprocess.run(argv, cwd=cwd)
    except FileNotFoundError as e:
        print(f"[preflight] could not run {argv[0]}: {e}")
        return 1 if required else 0
    if proc.returncode != 0:
        print(f"[preflight] step '{title}' exited with code {proc.returncode}")
    return proc.returncode


THIRD_PARTY = {
    "streamlit", "google", "google.generativeai", "pptx", "PIL",
    "chromadb", "fastapi", "pydantic", "pytesseract", "requests",
    "docx", "lxml", "uvicorn",
}


def _is_missing_third_party(exc: BaseException) -> bool:
    if not isinstance(exc, ModuleNotFoundError):
        return False
    missing = getattr(exc, "name", "") or ""
    return any(missing == t or missing.startswith(t + ".") for t in THIRD_PARTY)


def import_smoke() -> int:
    """
    Import app modules to catch syntax errors early.
    Missing third-party deps -> WARN (not FAIL), since a fresh workstation
    may not yet have `pip install -r requirements.txt`.
    """
    print("\n" + "=" * 72)
    print("[preflight] import smoke test")
    print("=" * 72)
    mods = [
        "app.models.models",
        "app.pptx_rag_quizzer.rag_core",
        "app.pptx_rag_quizzer.image",
        "app.pptx_rag_quizzer.utils",
    ]
    failed = []
    warned = []
    sys.path.insert(0, str(REPO_ROOT))
    sys.path.insert(0, str(REPO_ROOT / "app"))  # some modules use top-level imports
    for m in mods:
        try:
            __import__(m)
            print(f"  OK   import {m}")
        except SyntaxError as e:
            print(f"  FAIL import {m}: SyntaxError: {e}")
            failed.append(m)
        except ModuleNotFoundError as e:
            if _is_missing_third_party(e):
                print(f"  WARN import {m}: third-party dep missing ({e.name}) - run pip install -r requirements.txt")
                warned.append(m)
            else:
                print(f"  FAIL import {m}: ModuleNotFoundError: {e}")
                failed.append(m)
        except Exception as e:
            print(f"  FAIL import {m}: {type(e).__name__}: {e}")
            failed.append(m)

    if warned:
        print(f"  ({len(warned)} warning(s) - install deps before pushing)")
    return 0 if not failed else 1


def main() -> int:
    failures = 0

    # 1. doctor
    rc = run_step("doctor", [sys.executable, str(SCRIPTS / "doctor.py")], required=True)
    if rc != 0:
        failures += 1

    # 2. invariants
    rc = run_step(
        "check_invariants",
        [sys.executable, str(SCRIPTS / "check_invariants.py")],
        required=True,
    )
    if rc != 0:
        failures += 1

    # 3. pytest (optional - project may have no tests yet, or pytest not installed locally)
    tests_dir = REPO_ROOT / "tests"
    has_tests = tests_dir.exists() and any(tests_dir.rglob("test_*.py"))
    if has_tests:
        try:
            import pytest  # noqa: F401
            rc = run_step("pytest", [sys.executable, "-m", "pytest", "-q"], required=True)
            if rc != 0:
                failures += 1
        except ImportError:
            print("\n[preflight] pytest not installed - skipping (WARN).")
            print("           `pip install pytest` and re-run, or rely on CI.")
    else:
        print("\n[preflight] no tests/ directory with test_*.py - skipping pytest.")
        print("           Add a test with every change. See CONTRIBUTING.md section 5.")

    # 4. import smoke
    rc = import_smoke()
    if rc != 0:
        failures += 1

    print("\n" + "=" * 72)
    if failures == 0:
        print("[preflight] PASS - you are clear to open a PR.")
        print("            Fill the matching template from docs/templates/ in the PR body.")
        return 0
    print(f"[preflight] FAIL - {failures} step(s) failed. Fix before opening a PR.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
