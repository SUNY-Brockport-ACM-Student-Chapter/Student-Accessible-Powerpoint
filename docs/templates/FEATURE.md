# Feature Request / Implementation Template

> **Purpose.** This template is agent-optimized. A future AI agent should be able to read a filled-out copy and know exactly what to build, how to scope it, and what *not* to touch. Copy this file into the issue or PR description. Delete this block.

---

## 1. Summary (≤ 3 sentences)

*What does the user / stakeholder gain when this ships? No solution hints here — just the outcome.*

Example: "Users can upload `.docx` files alongside `.pptx` and receive the same accessibility treatment in the same UI flow."

## 2. Motivation

- **Who asked for this?** (course owner, researcher, IRB, user report)
- **Why now?**
- **What evidence do we have that this is needed?** (issue link, class transcript, WCAG audit, etc.)

## 3. Scope

### In scope
- *One bullet per capability.*

### Out of scope
- *Explicit exclusions. Use this to prevent scope creep.*

## 4. User-visible contract

Describe the change from the user's point of view. A checklist is fine.

- [ ] UI: what new controls / screens appear and where?
- [ ] Input: what formats / sizes are accepted?
- [ ] Output: what artifact is produced and downloadable?
- [ ] Copy: any IRB-relevant language? (consent screens, disclaimers)

## 5. Technical design

### 5.1 Branch

Which branch does this target and why? See [`../guardrails/BRANCHING.md`](../guardrails/BRANCHING.md).

- `main` | `Aggrement` | `Prod-v1` | `nextjs-impl`
- Cherry-pick plan (if cross-branch):

### 5.2 Files I expect to touch

Based on [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) repository map:

- `app/ppt_notes.py` — why
- `app/pptx_rag_quizzer/<mod>.py` — why
- `app/models/models.py` — why (new fields? keep invariants)
- `requirements.txt` / `requirements-app.txt` — new deps?
- `docs/…` — what needs to be updated

### 5.3 Data model changes

- New Pydantic fields: …
- New ChromaDB collection / schema: …
- Migration required: yes / no

### 5.4 External calls

- New Gemini model / prompt? → respects invariant #3 (60 s backoff)? [y/n]
- New ChromaDB endpoints? → added to the FastAPI wrapper (invariant #4)? [y/n]
- New secrets? → added to `.env.example` + [`../ops/SOP_SECRETS.md`](../ops/SOP_SECRETS.md)? [y/n]

### 5.5 Invariants audit

From [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md), tick each that this change touches and state how it preserves the invariant.

- [ ] #1 `order_number`
- [ ] #2 `cNvPr/@descr`
- [ ] #3 Gemini 60 s sleep
- [ ] #4 Chroma via FastAPI wrapper
- [ ] #5 Image normalization
- [ ] #6 `chroma/` data safety
- [ ] #7 Recursive group traversal
- [ ] #8 Consent gate (Aggrement only)

## 6. Accessibility (WCAG 2.1 AA)

- **Does this affect alt text quality?** How?
- **Does this affect reading order?** How?
- **Does this affect contrast / color reliance?** How?
- **Evidence of compliance:** (PowerPoint Accessibility Checker screenshot, screen-reader recording, etc.)

## 7. Risks & rollback

- **Blast radius:** local | prod | both
- **Worst-case failure:** …
- **Rollback plan:** see [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md); rollback SHA will be recorded at deploy time.

## 8. Testing plan

- [ ] Unit test(s): …
- [ ] Manual QA: upload an exemplar deck, verify … 
- [ ] Smoke test update needed (`scripts/smoke_test.py`): yes / no
- [ ] Accessibility check: PowerPoint Accessibility Checker on the output deck

## 9. Docs plan

- [ ] Update [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) (agent-facing facts)
- [ ] Update [`../PROJECT_OVERVIEW.md`](../PROJECT_OVERVIEW.md) (human narrative)
- [ ] Update [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) (if infra changes)
- [ ] Add/update invariant in [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md)?

## 10. Deployment plan

- [ ] Local only — no deploy needed
- [ ] Prod deploy required — follow [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md)
- [ ] Backfill / migration step needed: describe
- [ ] Secret rotation piggy-backed: see [`../ops/SOP_SECRETS.md`](../ops/SOP_SECRETS.md)

## 11. Open questions

*Anything you want the reviewer (or the next agent) to answer before coding starts.*

---

## Definition of done

- [ ] Code merged on the target branch.
- [ ] All items in §5.5 (invariants) and §6 (accessibility) verified.
- [ ] Pre-merge checklist complete ([`../guardrails/CHANGE_CHECKLIST.md`](../guardrails/CHANGE_CHECKLIST.md)).
- [ ] Docs updated.
- [ ] Deployed (if applicable) with a line in [`../ops/DEPLOY_LOG.md`](../ops/DEPLOY_LOG.md).
- [ ] Smoke test green post-deploy.
