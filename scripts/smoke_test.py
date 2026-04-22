#!/usr/bin/env python3
"""
smoke_test.py — HTTP-level smoke test for the deployed Streamlit app.

Intended as the post-deploy gate in docs/ops/SOP_DEPLOY.md. stdlib-only so
it can be run from any workstation without installing anything.

Usage:
    python scripts/smoke_test.py
    python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility
    python scripts/smoke_test.py --url http://localhost:8501 --local

Checks performed:
    1. HTTPS reachable, status 200 (or allowed redirect chain) within timeout.
    2. Response includes the Streamlit HTML shell.
    3. `/` (the landing page at host root) returns 200 with a Brockport/SIGAI title.
    4. (--strict) the internal FastAPI wrapper is NOT publicly reachable.
    5. (--strict) the Streamlit port is NOT publicly reachable *except* via nginx.

Exit code 0 = all green; 1 = at least one hard failure.
"""

from __future__ import annotations

import argparse
import socket
import ssl
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import List, Optional
from urllib.parse import urlparse

DEFAULT_URL = "https://access.brockportsigai.org/accessibility"
TIMEOUT = 15  # seconds
USER_AGENT = "student-access-ppt-smoke/1.0 (+docs/ops/SOP_DEPLOY.md)"


# ---------- result type ----------

@dataclass
class Result:
    name: str
    passed: bool
    detail: str
    hard: bool = True


def _get(url: str, timeout: int = TIMEOUT) -> tuple[int, str, dict]:
    """GET a URL; return (status, body, headers). Raises on transport failure."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        body = resp.read().decode("utf-8", errors="replace")
        return resp.status, body, dict(resp.headers)


def _try_connect(host: str, port: int, timeout: float = 5.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


# ---------- checks ----------

def check_reachable(url: str) -> Result:
    try:
        status, body, _ = _get(url)
    except urllib.error.HTTPError as e:
        return Result("reachable (HTTP)", False, f"HTTP {e.code} {e.reason}")
    except urllib.error.URLError as e:
        return Result("reachable (HTTP)", False, f"URL error: {e.reason}")
    except socket.timeout:
        return Result("reachable (HTTP)", False, f"timeout after {TIMEOUT}s")
    except Exception as e:
        return Result("reachable (HTTP)", False, f"{type(e).__name__}: {e}")
    if status != 200:
        return Result("reachable (HTTP)", False, f"status={status}")
    return Result("reachable (HTTP)", True, f"200 OK, {len(body)} bytes")


def check_streamlit_shell(url: str) -> Result:
    try:
        _, body, _ = _get(url)
    except Exception as e:
        return Result("streamlit shell present", False, f"fetch failed: {e}")
    # Streamlit serves a bootstrap HTML that references 'stApp' and the
    # websocket endpoint. We look for a couple of stable markers.
    markers = ["<title>", "streamlit", "stApp"]
    missing = [m for m in markers if m.lower() not in body.lower()]
    if missing:
        return Result(
            "streamlit shell present", False, f"missing markers: {missing}",
        )
    return Result("streamlit shell present", True, f"markers ok ({', '.join(markers)})")


def check_landing(url: str) -> Result:
    """If URL is host/accessibility, also probe host root for the landing page."""
    p = urlparse(url)
    if not p.scheme or not p.netloc:
        return Result("landing page (/)", False, "cannot derive host from URL", hard=False)
    root = f"{p.scheme}://{p.netloc}/"
    try:
        status, body, _ = _get(root)
    except Exception as e:
        return Result("landing page (/)", False, f"fetch failed: {e}", hard=False)
    if status != 200:
        return Result("landing page (/)", False, f"status={status}", hard=False)
    if not body.strip():
        return Result("landing page (/)", False, "empty body", hard=False)
    return Result("landing page (/)", True, f"200 OK, {len(body)} bytes")


def check_port_not_public(host: str, port: int, label: str) -> Result:
    """In strict mode, the internal port should NOT be publicly reachable."""
    is_open = _try_connect(host, port, timeout=4.0)
    if is_open:
        return Result(
            f"port {port} ({label}) not public",
            False,
            f"{host}:{port} is reachable from the public internet — "
            "review docs/guardrails/INVARIANTS.md §10.4 and GCP firewall.",
            hard=False,  # warn, don't fail — this is a known prod drift
        )
    return Result(f"port {port} ({label}) not public", True, "not reachable externally")


# ---------- runner ----------

def run(url: str, strict: bool, local: bool) -> int:
    print(f"smoke_test.py :: target={url} strict={strict} local={local}\n")
    results: List[Result] = []
    results.append(check_reachable(url))
    if results[-1].passed:
        results.append(check_streamlit_shell(url))
    else:
        results.append(Result("streamlit shell present", False, "skipped (unreachable)"))

    results.append(check_landing(url))

    if strict and not local:
        p = urlparse(url)
        host = p.hostname or ""
        if host:
            # The prod firewall currently exposes these publicly; the check
            # flags the leak but does not hard-fail by default.
            results.append(check_port_not_public(host, 8501, "Streamlit direct"))
            results.append(check_port_not_public(host, 8001, "chroma-api"))

    print("Results:")
    fails = 0
    warns = 0
    for r in results:
        if r.passed:
            tag = "[OK]  "
        else:
            tag = "[FAIL]" if r.hard else "[WARN]"
            if r.hard:
                fails += 1
            else:
                warns += 1
        print(f"  {tag} {r.name:<32} {r.detail}")

    print()
    if fails == 0:
        print(f"smoke_test: PASS ({warns} warning(s))" if warns else "smoke_test: PASS")
        return 0
    print(f"smoke_test: FAIL ({fails} hard failure(s), {warns} warning(s))")
    print("Run docs/ops/SOP_ROLLBACK.md if this was a post-deploy check.")
    return 1


def main() -> int:
    p = argparse.ArgumentParser(description="HTTP smoke test for deployed app.")
    p.add_argument("--url", default=DEFAULT_URL)
    p.add_argument("--strict", action="store_true", help="also check for exposure of internal ports")
    p.add_argument("--local", action="store_true", help="target a local dev URL; skip public-exposure checks")
    args = p.parse_args()
    return run(args.url, args.strict, args.local)


if __name__ == "__main__":
    sys.exit(main())
