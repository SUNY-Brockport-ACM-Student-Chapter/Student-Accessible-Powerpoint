# Refactor Template

> **Purpose.** For structural changes that do **not** alter user-visible behavior: renaming, module splits, type hints, extracting helpers, moving config into git, promoting a branch, etc. Copy into the issue/PR and delete this block.

A change that intentionally alters behavior is a [`FEATURE`](FEATURE.md) or [`BUG`](BUG.md), not a refactor. Be strict about this.

---

## 1. TL;DR

One sentence: what shape changes, and why behavior stays identical.

## 2. Motivation

Why now? Pick at least one:

- [ ] Closes a tech-debt item (link to [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md) §10 entry)
- [ ] Unblocks a planned feature — name it
- [ ] Reduces footgun risk (link to [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) footgun)
- [ ] Enforces an invariant in code instead of docs
- [ ] Other (describe)

## 3. Scope

### In scope (structural changes)
- …

### Strictly out of scope
- Any behavior change. If reviewers detect one, this PR gets rejected as "refactor + feature bundled".

## 4. Branch

Per [`../guardrails/BRANCHING.md`](../guardrails/BRANCHING.md):

- Target: …
- Cross-branch plan (almost always `main` first, then cherry-pick to `Aggrement`):

## 5. Before/after

- **Before** (paste current layout / signature / structure, or link to it):

```
<current>
```

- **After** (paste new):

```
<new>
```

## 6. Invariant safety

This is the whole point of the refactor template. For each, explain why it still holds after the refactor.

- [ ] #1 `order_number`
- [ ] #2 `cNvPr/@descr` XML path
- [ ] #3 Gemini 60 s sleep
- [ ] #4 Chroma via FastAPI wrapper
- [ ] #5 Image normalization
- [ ] #6 `chroma/` safety
- [ ] #7 Recursive group traversal
- [ ] #8 Consent gate (Aggrement)

If a refactor makes an invariant *impossible to violate* (e.g. encodes it as a type), update [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md) to note the code enforcement and change the Check to reference the test.

## 7. Equivalence evidence

How do we know behavior didn't change?

- [ ] Existing tests still pass (paste `pytest` output).
- [ ] New behavior-parity tests (e.g. same input deck → byte-identical or diff-only-in-whitespace output):
- [ ] Manual round-trip: same deck before/after produced same alt text set.
- [ ] `scripts/smoke_test.py` passes against the refactored build.

## 8. Risk

- **Blast radius:** small | medium | large
- **Most likely regression surface:** …
- **Rollback:** revert PR; if production-deployed, see [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md).

## 9. Deployment

- [ ] No prod deploy (pure code refactor on `main`).
- [ ] Prod deploy required — follow [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md) and treat it like a regular deploy.

## 10. Docs updates

- [ ] [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) — repository map if modules moved.
- [ ] [`../PROJECT_OVERVIEW.md`](../PROJECT_OVERVIEW.md) — architecture diagram if it changes.
- [ ] [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md) — note if the refactor closes a tech-debt item (§10) or encodes an invariant in code.
- [ ] [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) — if deployment shape changes (systemd unit rename, nginx route rename, directory moves).

## 11. Follow-ups

Refactors often expose more debt. List follow-ups here as separate issues so they don't sneak into this PR.

- …
