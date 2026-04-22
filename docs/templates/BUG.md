# Bug Report / Fix Template

> **Purpose.** Agent-optimized template for reporting *and* fixing a defect. If you are reporting, fill §§1–4. If you are fixing, fill §§5–10 too. Copy this file into the issue/PR and delete this block.

---

## 1. TL;DR

One sentence: what's wrong.

## 2. Environment

- [ ] Local dev (Windows / macOS / Linux)
- [ ] Production (`access.brockportsigai.org/accessibility`)
- [ ] Older prod instance (`35.196.195.118`)
- [ ] Other (describe)

- Branch / commit observed on: …
- Python version, OS, browser if UI bug: …

## 3. Repro

Exact steps. Include the input artifact (`.pptx` filename / source) if relevant.

1. 
2. 
3. 

**Expected:** …
**Actual:** …

## 4. Evidence

- Screenshot / recording: …
- Stack trace (from `journalctl -u student-access-ppt` if prod, or local terminal): paste in a code block
- `consent_responses.csv` / `chroma/` effects (if relevant):

```text
<paste traceback or log excerpt here>
```

---

*(Everything below is for the fix PR.)*

## 5. Root cause

A paragraph: *why* this happens. Not the fix yet. Tie it back to the code path — which module, which function, which invariant broke.

- Invariant violated (if any, from [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md)): …
- Regression introduced in commit: `<sha>` (use `git log -p -S <symbol>`)
- Branches affected: `main` / `Aggrement` / `Prod-v1` / …

## 6. Fix

Summary of the change. Link the diff. If the fix is only on one branch but the bug exists on another, describe the cherry-pick plan ([`../guardrails/BRANCHING.md`](../guardrails/BRANCHING.md)).

## 7. Test that would have caught this

Every bug fix adds a regression test. Exceptions must be justified.

- [ ] Unit test added in `tests/…`
- [ ] Invariant check added in `scripts/check_invariants.py`
- [ ] Smoke test updated in `scripts/smoke_test.py`
- [ ] N/A — justification: …

## 8. Risk

- **Blast radius of the fix:** local | prod | both
- **Could the fix introduce a new regression?** In which module?
- **Rollback plan:** see [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md).

## 9. Deployment

- [ ] Needs production deploy — follow [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md)
- [ ] Urgency (hotfix now / next regular deploy / low-priority):
- [ ] Incident already open ([`../ops/INCIDENT_LOG.md`](../ops/INCIDENT_LOG.md)): yes (link) / no

## 10. Post-merge

- [ ] DEPLOY_LOG.md entry added.
- [ ] INCIDENT_LOG.md closed (if an incident was open).
- [ ] [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) footgun section updated if the bug exposed a new trap.
