#!/usr/bin/env python3
"""
doctor.py — fast offline environment sanity check.

Runs on a fresh clone or an existing dev setup and reports whether the
environment is ready to develop / run the app. stdlib-only; does not hit
the network.

Exit code 0 = all green. 1 = at least one hard failure.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent


# ---------- reporting helpers ----------

def _supports_color() -> bool:
    if not sys.stdout.isatty():
        return False
    if os.name == "nt":
        return "WT_SESSION" in os.environ or os.environ.get("TERM") not in (None, "")
    return True


def _c(s: str, code: str) -> str:
    if not _supports_color():
        return s
    return f"\033[{code}m{s}\033[0m"


GREEN = lambda s: _c(s, "32")  # noqa: E731
RED = lambda s: _c(s, "31")  # noqa: E731
YELLOW = lambda s: _c(s, "33")  # noqa: E731
DIM = lambda s: _c(s, "2")  # noqa: E731


class Check:
    def __init__(self, name: str, hard: bool = True):
        self.name = name
        self.hard = hard
        self.passed = False
        self.detail = ""

    def ok(self, detail: str = "") -> "Check":
        self.passed = True
        self.detail = detail
        return self

    def fail(self, detail: str) -> "Check":
        self.passed = False
        self.detail = detail
        return self

    def render(self) -> str:
        if self.passed:
            return f"  {GREEN('[OK]')}    {self.name}" + (
                f"  {DIM(self.detail)}" if self.detail else ""
            )
        tag = RED("[FAIL]") if self.hard else YELLOW("[WARN]")
        return f"  {tag}  {self.name}\n         {self.detail}"


# ---------- checks ----------

def check_python_version() -> Check:
    c = Check("Python 3.11.x")
    major, minor = sys.version_info[:2]
    if major == 3 and minor == 11:
        return c.ok(f"running {sys.version.split()[0]}")
    return c.fail(
        f"prod runs 3.11.2; you are on {major}.{minor}. "
        "Create a 3.11 venv to match prod ABI."
    )


def check_in_venv() -> Check:
    c = Check("Running inside a virtualenv", hard=False)
    in_venv = (
        hasattr(sys, "real_prefix")
        or (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix)
    )
    if in_venv:
        return c.ok(f"prefix={sys.prefix}")
    return c.fail("not in a venv; install deps into a venv to match prod.")


def check_required_tools() -> List[Check]:
    out = []
    for tool in ["git", "pip"]:
        c = Check(f"`{tool}` on PATH")
        if shutil.which(tool):
            out.append(c.ok())
        else:
            out.append(c.fail(f"{tool} not found on PATH"))
    return out


def check_repo_layout() -> List[Check]:
    expected_dirs = ["app", "app/pptx_rag_quizzer", "app/models", "app/chroma-api", "docs"]
    expected_files = [
        "requirements.txt",
        "requirements-app.txt",
        ".env.example",
        "app/ppt_notes.py",
        "app/models/models.py",
        "app/pptx_rag_quizzer/rag_core.py",
        "app/pptx_rag_quizzer/utils.py",
        "app/pptx_rag_quizzer/image.py",
        "app/chroma-api/app.py",
    ]
    out = []
    for d in expected_dirs:
        c = Check(f"dir exists: {d}")
        p = REPO_ROOT / d
        out.append(c.ok() if p.is_dir() else c.fail(f"missing: {p}"))
    for f in expected_files:
        c = Check(f"file exists: {f}")
        p = REPO_ROOT / f
        out.append(c.ok() if p.is_file() else c.fail(f"missing: {p}"))
    return out


def check_env_file() -> Check:
    c = Check(".env present with GOOGLE_API_KEY", hard=False)
    env = REPO_ROOT / ".env"
    if not env.is_file():
        return c.fail(
            "no .env file. Copy .env.example -> .env and paste your GOOGLE_API_KEY."
        )
    try:
        txt = env.read_text(encoding="utf-8", errors="replace")
    except Exception as e:
        return c.fail(f"could not read .env: {e}")
    has_key = any(
        line.strip().startswith("GOOGLE_API_KEY=") and "=" in line and line.strip().split("=", 1)[1]
        for line in txt.splitlines()
    )
    if not has_key:
        return c.fail(".env exists but GOOGLE_API_KEY is empty or missing.")
    return c.ok()


def check_not_gitignored_assets() -> List[Check]:
    """Warn if files that must not be committed already are."""
    out = []
    risky = [".env", "chroma", "chroma-db"]
    try:
        res = subprocess.run(
            ["git", "ls-files"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        tracked = set(res.stdout.splitlines())
    except FileNotFoundError:
        return [Check("git installed", hard=False).fail("git not on PATH")]
    for r in risky:
        c = Check(f"'{r}' is not tracked", hard=True)
        bad = [p for p in tracked if p == r or p.startswith(r + "/") or p.startswith(r + os.sep)]
        out.append(c.fail(f"tracked: {bad}") if bad else c.ok())
    return out


def check_disk_space() -> Check:
    c = Check("at least 500 MB free on repo drive", hard=False)
    try:
        usage = shutil.disk_usage(REPO_ROOT)
        free_mb = usage.free // (1024 * 1024)
        if free_mb < 500:
            return c.fail(f"only {free_mb} MB free")
        return c.ok(f"{free_mb} MB free")
    except Exception as e:
        return c.fail(str(e))


def check_git_branch_known() -> Check:
    c = Check("On a known branch", hard=False)
    try:
        res = subprocess.run(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            cwd=REPO_ROOT, capture_output=True, text=True, check=False,
        )
        branch = res.stdout.strip()
        known = {"main", "Aggrement", "Prod-v1", "nextjs-impl", "RAG-integration-branch"}
        if branch in known:
            return c.ok(f"on '{branch}'")
        if branch.startswith(("feat/", "fix/", "chore/", "refactor/", "deps/")):
            return c.ok(f"on '{branch}' (short-lived)")
        return c.fail(
            f"on '{branch}'. See docs/guardrails/BRANCHING.md — use feat/ fix/ chore/ refactor/ deps/."
        )
    except Exception as e:
        return c.fail(str(e))


# ---------- entry point ----------

def main() -> int:
    print(f"doctor.py :: environment check for {REPO_ROOT.name}\n")

    batches: List[Tuple[str, List[Check]]] = [
        ("Interpreter",        [check_python_version(), check_in_venv()]),
        ("Tools",              check_required_tools()),
        ("Repo layout",        check_repo_layout()),
        ("Config",             [check_env_file()]),
        ("Safety",             check_not_gitignored_assets()),
        ("Capacity",           [check_disk_space()]),
        ("Branch",             [check_git_branch_known()]),
    ]

    hard_fails = 0
    for title, checks in batches:
        print(f"[{title}]")
        for c in checks:
            print(c.render())
            if not c.passed and c.hard:
                hard_fails += 1
        print()

    if hard_fails == 0:
        print(GREEN("doctor: all hard checks pass."))
        return 0
    print(RED(f"doctor: {hard_fails} hard check(s) failed."))
    return 1


if __name__ == "__main__":
    sys.exit(main())
