# scripts/

Operational helpers for the Student-Accessible-Powerpoint project.
All scripts are **stdlib-only** so they work on any developer / agent machine
and on the production VM without extra installs.

| Script | What it does | When to run |
|---|---|---|
| [`doctor.py`](doctor.py) | Offline environment check: Python version, venv, repo layout, `.env`, git state, disk space. | Right after `git clone`. Whenever "it doesn't work on my machine". |
| [`check_invariants.py`](check_invariants.py) | Static regex-level guards for the invariants in [`../docs/guardrails/INVARIANTS.md`](../docs/guardrails/INVARIANTS.md). | Before every PR. In CI once CI exists. |
| [`preflight.py`](preflight.py) | Aggregates `doctor` + `check_invariants` + `pytest` + import smoke. | Before opening a PR. |
| [`smoke_test.py`](smoke_test.py) | HTTP smoke test against a deployed URL (defaults to prod). | Post-deploy gate; see [`../docs/ops/SOP_DEPLOY.md`](../docs/ops/SOP_DEPLOY.md). |

## Typical flows

**Fresh clone:**

```bash
python scripts/doctor.py
```

**Before opening a PR:**

```bash
python scripts/preflight.py
```

**After a prod deploy:**

```bash
python scripts/smoke_test.py --strict
```

**Debug a single invariant:**

```bash
python scripts/check_invariants.py --list
python scripts/check_invariants.py --only order_number alt_text_xml
```

## Adding a new script

- Keep it stdlib-only unless there's a compelling reason.
- Print a clear final line with `PASS` / `FAIL`.
- Return exit code 0 on pass, 1 on hard failure, 2 on usage error.
- Add it to this README and to [`preflight.py`](preflight.py) if it should gate PRs.
