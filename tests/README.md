# tests/

Test suite for Student-Accessible-Powerpoint.

## Running

```bash
python -m pytest -q
```

Preflight (`scripts/preflight.py`) runs this automatically if any `test_*.py`
files exist.

## Policy

- Every PR adds at least one test for the code it changes, or explicitly
  justifies "N/A" in the PR description. See
  [`../CONTRIBUTING.md`](../CONTRIBUTING.md) section 5.
- Prefer fast, offline, stdlib + pytest tests. No network calls to Gemini
  or production ChromaDB from the test suite.
- When you need to exercise an invariant from
  [`../docs/guardrails/INVARIANTS.md`](../docs/guardrails/INVARIANTS.md),
  mirror the check in `scripts/check_invariants.py` if possible so CI can
  enforce it without pytest.

## Layout

```
tests/
  conftest.py          # shared fixtures (add as needed)
  test_invariants.py   # re-runs scripts/check_invariants.py as a pytest
  test_models.py       # Pydantic model smoke tests
  ...
```

New test files must follow `test_*.py`.
