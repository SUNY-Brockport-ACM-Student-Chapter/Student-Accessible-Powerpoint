# Change Checklist

Before you mark a PR "ready for review", walk through this list. Paste it into the PR description and check items off. If an item is N/A, write "N/A — reason" instead of deleting the line.

---

## A. Intent

- [ ] I picked the right template ([`FEATURE`](../templates/FEATURE.md) / [`UPDATE`](../templates/UPDATE.md) / [`BUG`](../templates/BUG.md) / [`REFACTOR`](../templates/REFACTOR.md)) and filled it out in the PR description.
- [ ] I identified which branch this targets ([`BRANCHING.md`](BRANCHING.md)) and why.
- [ ] If this is a behavior change, I noted the expected user-visible difference.

## B. Invariants

- [ ] I read [`INVARIANTS.md`](INVARIANTS.md) and confirmed none are violated.
- [ ] If the change touches parsing, rebuild, or XML writing, I verified #1 (`order_number`) and #2 (`cNvPr/@descr`) still hold.
- [ ] If the change touches Gemini calls, I preserved the 60 s rate-limit sleep (invariant #3).
- [ ] If the change touches ChromaDB, access still goes through the FastAPI wrapper (invariant #4).
- [ ] If the change touches images, PIL `P` / WMF / EMF normalization still happens before Gemini (invariant #5).
- [ ] If the change touches the Aggrement branch, the consent gate still runs before upload (invariant #8).

## C. Local validation

- [ ] `python scripts/doctor.py` passes.
- [ ] `python scripts/check_invariants.py` passes.
- [ ] `python -m pytest -q` passes (or: N/A — no tests apply, with justification).
- [ ] `python scripts/preflight.py` passes.
- [ ] I ran the three-process stack locally and uploaded a `.pptx` end-to-end at least once.

## D. Dependencies & config

- [ ] If I added a Python dep: it's pinned in `requirements.txt` **and** `requirements-app.txt` (if the Streamlit UI uses it).
- [ ] If I added an env var: it's listed (with a placeholder) in `.env.example` and documented in [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md).
- [ ] If I added a new port or public route: it's noted in [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) and in a `deploy` label on the PR.

## E. Accessibility (non-negotiable)

- [ ] My change does not reduce WCAG 2.1 AA conformance ([`../PROJECT_OVERVIEW.md`](../PROJECT_OVERVIEW.md) ADA section).
- [ ] If my change affects alt-text generation, I spot-checked the output deck in PowerPoint's Accessibility Checker on at least one slide.
- [ ] If my change affects reading order, groups, or tables, I verified with a screen reader (NVDA / VoiceOver) or documented why I couldn't.

## F. Production safety

- [ ] If this needs a deploy, I linked to [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md) and identified a rollback SHA.
- [ ] If this affects the Gemini model, secret format, or startup order, I flagged it with `breaking-change` on the PR.
- [ ] I did **not** commit anything in `.env`, API keys, or generated `.pptx` files.
- [ ] I did **not** modify `chroma/` or anything else that would be affected by `git clean -fdx` on prod.

## G. Docs

- [ ] If behavior changed, I updated [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) (agent-facing) and/or [`../PROJECT_OVERVIEW.md`](../PROJECT_OVERVIEW.md) (human-facing).
- [ ] If an operational procedure changed, I updated the matching SOP under `docs/ops/`.
- [ ] If I introduced a new guardrail, I added it to [`INVARIANTS.md`](INVARIANTS.md) with a `Check` command.

## H. Tests

- [ ] I added at least one test for the code I changed, OR explicitly stated why a test is impractical.
- [ ] I updated `scripts/smoke_test.py` if I added or changed a public URL, endpoint, or health indicator.

---

## If any box is unchecked

Either check it or replace it with `N/A — <reason>`. Unchecked boxes in an open PR = request-changes.
