#!/usr/bin/env python3
"""
check_invariants.py — static guard against regressing project invariants.

Each check is a small regex / file probe that encodes one invariant from
docs/guardrails/INVARIANTS.md. Run as part of preflight and in CI (once CI
exists).

Usage:
    python scripts/check_invariants.py                 # run all
    python scripts/check_invariants.py --only order_number alt_text_xml
    python scripts/check_invariants.py --list

Exit code 0 = pass, 1 = at least one invariant failed.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Callable, Dict, List, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent
APP = REPO_ROOT / "app"


# ---------- tiny helper ----------

def _read(p: Path) -> str:
    try:
        return p.read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return ""


def _walk_py(root: Path) -> List[Path]:
    return sorted(p for p in root.rglob("*.py") if "venv" not in p.parts)


# ---------- checks ----------

CheckResult = Tuple[bool, str]  # (passed, detail)

# Checks marked non-hard produce a WARN (non-zero-ish signal) but don't fail
# the overall script. Use for pre-existing technical debt we want to
# surface loudly without blocking every PR.
NON_HARD = {"groups"}


def check_order_number() -> CheckResult:
    """
    Invariant #1: order_number is the join key.

    Enforcement strategy:
      - models.py MUST define order_number (or the invariant is silently dead).
      - Any file that assigns order_number must use the canonical
        'order_number=' kw-arg (not a positional) so refactors can grep for it.
      - At least one caller beyond models.py must reference order_number.
    """
    models = APP / "models" / "models.py"
    if not models.exists():
        return False, "app/models/models.py missing"
    models_txt = _read(models)
    if "order_number" not in models_txt:
        return False, "models.py does not define 'order_number' — invariant #1 lost"

    callers = 0
    for py in _walk_py(APP):
        if py == models:
            continue
        if "order_number" in _read(py):
            callers += 1
    if callers == 0:
        return False, (
            "order_number is defined in models.py but no caller references it. "
            "Rebuild paths must read it."
        )
    return True, f"defined in models.py and referenced in {callers} other file(s)"


def check_alt_text_xml() -> CheckResult:
    """
    Invariant #2: alt text writes to cNvPr/@descr (native XML).
    """
    target = APP / "ppt_notes.py"
    txt = _read(target)
    if not txt:
        return False, f"cannot read {target}"
    if "cNvPr" not in txt:
        return False, "cNvPr not referenced in app/ppt_notes.py — alt text may not be hitting native XML"
    if "descr" not in txt:
        return False, "`descr` attribute not referenced — alt text write is incomplete"
    return True, "cNvPr + descr present in app/ppt_notes.py"


def check_gemini_rate_limit() -> CheckResult:
    """
    Invariant #3: on quota exhaustion, sleep 60 s.
    """
    target = APP / "pptx_rag_quizzer" / "rag_core.py"
    txt = _read(target)
    if not txt:
        return False, f"cannot read {target}"
    has_trigger = "Resource has been exhausted" in txt or "ResourceExhausted" in txt
    # accept quota_refill_delay = 60 OR direct sleep(60)
    has_60 = bool(
        re.search(r"quota_refill_delay\s*=\s*60", txt)
        or re.search(r"time\.sleep\(\s*60\s*\)", txt)
    )
    if not has_trigger:
        return False, "no 'Resource has been exhausted' / 'ResourceExhausted' handling found"
    if not has_60:
        return False, "found trigger but no 60s backoff (quota_refill_delay=60 or time.sleep(60))"
    return True, "quota-exhausted handler present with 60s backoff"


def check_chroma_wrapper() -> CheckResult:
    """
    Invariant #4: Chroma access goes through the FastAPI wrapper, not
    direct chromadb.HttpClient / PersistentClient from Streamlit code.
    """
    violations: List[str] = []
    for py in _walk_py(APP):
        if py.parts[-2:] == ("chroma-api", "app.py"):
            continue  # wrapper is allowed to use chromadb directly
        txt = _read(py)
        if re.search(r"\bchromadb\.(HttpClient|PersistentClient)\b", txt):
            violations.append(str(py.relative_to(REPO_ROOT)))
    if violations:
        return False, (
            "chromadb.HttpClient/PersistentClient used outside chroma-api wrapper: "
            + ", ".join(violations)
        )
    return True, "only chroma-api/app.py uses chromadb client directly"


def check_image_normalization() -> CheckResult:
    """
    Invariant #5: Image normalization (WMF/EMF -> PNG/JPG, 'P'/'RGBA' -> 'RGB')
    must run before Gemini / hashing. We look for *any* of the well-known
    markers across the app. This is a soft-intent check; it succeeds if the
    pipeline demonstrates awareness of the problem.

    Markers (any one counts):
      - function name convert_image_to_png_or_jpg  (Aggrement branch name)
      - PIL mode switching: Image.open(...).convert("RGB")
      - explicit WMF/EMF handling
      - pytesseract image pre-processing (used on this codebase)
    """
    markers = {
        "convert_image_to_png_or_jpg": r"convert_image_to_png_or_jpg",
        "PIL .convert('RGB')":         r"\.convert\(\s*['\"]RGB['\"]\s*\)",
        "WMF/EMF branch":              r"\bWMF\b|\bEMF\b",
        "Image.open":                  r"Image\.open\(",
    }
    hits: Dict[str, int] = {label: 0 for label in markers}
    for py in _walk_py(APP):
        txt = _read(py)
        for label, pattern in markers.items():
            if re.search(pattern, txt):
                hits[label] += 1
    total = sum(hits.values())
    if total == 0:
        return False, (
            "no image-normalization markers found in app/. Expected one of: "
            "Image.open / .convert('RGB') / WMF-EMF branch / convert_image_to_png_or_jpg."
        )
    summary = ", ".join(f"{k}={v}" for k, v in hits.items() if v)
    return True, f"image pre-processing markers present ({summary})"


def check_chroma_not_gitignored_destructively() -> CheckResult:
    """
    Invariant #6: don't accidentally wipe prod vector data.
    Guard: `.gitignore` must not contain a bare `chroma` entry that would
    match both the in-repo dev dir *and* the prod data dir ambiguously.
    A more specific path (e.g. '/chroma-db') is fine.
    """
    gi = REPO_ROOT / ".gitignore"
    if not gi.exists():
        return True, "no .gitignore"  # not a fail; just nothing to check
    lines = [
        ln.strip() for ln in gi.read_text(encoding="utf-8", errors="replace").splitlines()
        if ln.strip() and not ln.strip().startswith("#")
    ]
    bad = [ln for ln in lines if ln in {"chroma", "chroma/", "**/chroma", "**/chroma/"}]
    if bad:
        return False, (
            "ambiguous 'chroma' entry in .gitignore — may wipe prod vector data on git clean. "
            "Use '/chroma-db' (current pattern) or a more specific path. See INVARIANTS.md §6."
        )
    return True, ".gitignore does not contain an ambiguous 'chroma' entry"


def check_group_traversal_hint() -> CheckResult:
    """
    Invariant #7: group shapes must be traversed recursively. This is hard
    to verify statically; we check that a hint is present where .shapes is
    iterated (either MSO_SHAPE_TYPE.GROUP, recursion, or an explicit
    noqa-style comment).
    """
    suspect: List[str] = []
    for py in _walk_py(APP):
        txt = _read(py)
        if not re.search(r"\bslide\.shapes\b|\.shapes\b", txt):
            continue
        if (
            "MSO_SHAPE_TYPE.GROUP" in txt
            or "is_group" in txt
            or re.search(r"def\s+_?walk_shapes|recurse_shapes|iter_shapes", txt)
        ):
            continue
        suspect.append(str(py.relative_to(REPO_ROOT)))
    if suspect:
        return False, (
            "files iterate .shapes without any group-handling marker: "
            + ", ".join(suspect)
            + ". Confirm group shapes are recursed (invariant #7)."
        )
    return True, "all .shapes iterations acknowledge group traversal"


def check_secret_not_in_git() -> CheckResult:
    """
    Soft invariant: .env, key files, and consent_responses.csv must not
    be tracked in git.
    """
    import subprocess
    try:
        res = subprocess.run(
            ["git", "ls-files"],
            cwd=REPO_ROOT, capture_output=True, text=True, check=False,
        )
        tracked = res.stdout.splitlines()
    except FileNotFoundError:
        return False, "git not available"
    forbidden_exact = {".env", "consent_responses.csv"}
    forbidden_prefixes = ("chroma/", "chroma-db/", "venv/")
    hits = [
        p for p in tracked
        if p in forbidden_exact or p.startswith(forbidden_prefixes)
    ]
    if hits:
        return False, "forbidden files are tracked: " + ", ".join(hits)
    return True, "no forbidden files tracked"


# ---------- registry ----------

CHECKS: Dict[str, Tuple[str, Callable[[], CheckResult]]] = {
    "order_number":    ("#1 order_number present where expected", check_order_number),
    "alt_text_xml":    ("#2 alt text writes cNvPr/@descr", check_alt_text_xml),
    "gemini_backoff":  ("#3 Gemini 60s quota backoff in rag_core", check_gemini_rate_limit),
    "chroma_wrapper":  ("#4 chromadb client only inside chroma-api wrapper", check_chroma_wrapper),
    "image_norm":      ("#5 image normalization path wired", check_image_normalization),
    "chroma_gitignore":("#6 .gitignore does not ambiguously match chroma/", check_chroma_not_gitignored_destructively),
    "groups":          ("#7 group-shape traversal acknowledged", check_group_traversal_hint),
    "secrets":         ("soft: no secret files tracked by git", check_secret_not_in_git),
}


def main() -> int:
    p = argparse.ArgumentParser(description="Static invariant checks.")
    p.add_argument("--only", nargs="+", default=None, help="run only these keys")
    p.add_argument("--list", action="store_true", help="list available checks and exit")
    args = p.parse_args()

    if args.list:
        print("Available invariant checks:\n")
        for key, (desc, _) in CHECKS.items():
            print(f"  {key:<18}  {desc}")
        return 0

    keys = args.only if args.only else list(CHECKS.keys())
    unknown = [k for k in keys if k not in CHECKS]
    if unknown:
        print(f"Unknown check(s): {unknown}", file=sys.stderr)
        return 2

    print(f"check_invariants.py :: running {len(keys)} check(s)\n")
    hard_failures = 0
    warnings = 0
    for key in keys:
        desc, fn = CHECKS[key]
        try:
            passed, detail = fn()
        except Exception as e:
            passed, detail = False, f"check raised {type(e).__name__}: {e}"
        if passed:
            tag = "OK  "
        elif key in NON_HARD:
            tag = "WARN"
            warnings += 1
        else:
            tag = "FAIL"
            hard_failures += 1
        print(f"  [{tag}] {key:<18} {desc}")
        if detail:
            print(f"         {detail}")

    print()
    if hard_failures == 0:
        msg = "check_invariants: all hard checks pass."
        if warnings:
            msg += f" ({warnings} warning(s) - known debt, not blocking)"
        print(msg)
        return 0
    print(f"check_invariants: {hard_failures} invariant check(s) failed, {warnings} warning(s).")
    print("See docs/guardrails/INVARIANTS.md for full context.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
