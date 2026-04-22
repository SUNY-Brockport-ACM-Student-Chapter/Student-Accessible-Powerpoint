<!--
Thanks for contributing to Student-Accessible-Powerpoint.

Before you open this PR:
  1. You MUST have read AGENTS.md (agents) or CONTRIBUTING.md (humans).
  2. Replace the TL;DR below and fill out the change template that matches.
  3. Fill out the pre-merge checklist at the bottom.

Pick ONE template below and DELETE the others + this HTML comment.

  - New capability            -> docs/templates/FEATURE.md
  - Dependency / version bump -> docs/templates/UPDATE.md
  - Bug fix                   -> docs/templates/BUG.md
  - Structural / no behavior  -> docs/templates/REFACTOR.md
-->

## TL;DR

<!-- One sentence: what does this PR do and why. -->

## Change type

- [ ] Feature (copy structure from [`docs/templates/FEATURE.md`](../docs/templates/FEATURE.md))
- [ ] Update / dependency bump (copy structure from [`docs/templates/UPDATE.md`](../docs/templates/UPDATE.md))
- [ ] Bug fix (copy structure from [`docs/templates/BUG.md`](../docs/templates/BUG.md))
- [ ] Refactor (copy structure from [`docs/templates/REFACTOR.md`](../docs/templates/REFACTOR.md))
- [ ] Docs only

## Target branch

- [ ] `main` (canonical)
- [ ] `Aggrement` (production)
- [ ] `Prod-v1` (Docker)
- [ ] Other: …

See [`docs/guardrails/BRANCHING.md`](../docs/guardrails/BRANCHING.md) for the policy.

---

## Change template (paste the filled-out template here)

<!-- Replace this block with your chosen template, fully filled out. -->

---

## Pre-merge checklist

Full form in [`docs/guardrails/CHANGE_CHECKLIST.md`](../docs/guardrails/CHANGE_CHECKLIST.md). Tick each or write `N/A - reason`.

**Intent**
- [ ] Correct change template pasted above
- [ ] Correct target branch selected

**Invariants** ([`docs/guardrails/INVARIANTS.md`](../docs/guardrails/INVARIANTS.md))
- [ ] No invariant violated
- [ ] #1 `order_number` preserved (if parsing/rebuild touched)
- [ ] #2 `cNvPr/@descr` preserved (if alt-text touched)
- [ ] #3 60s Gemini backoff preserved (if Gemini touched)
- [ ] #4 Chroma wrapper preserved (if Chroma touched)
- [ ] #5 Image normalization preserved (if images touched)
- [ ] #8 Consent gate preserved (if on `Aggrement`)

**Local validation**
- [ ] `python scripts/doctor.py` passes
- [ ] `python scripts/check_invariants.py` passes
- [ ] `python scripts/preflight.py` passes
- [ ] `python -m pytest -q` passes (or N/A)
- [ ] Manual three-process run + `.pptx` round-trip succeeded (or N/A)

**Accessibility (WCAG 2.1 AA)**
- [ ] No regression in alt-text quality or reading order
- [ ] PowerPoint Accessibility Checker spot-check on output deck (or N/A)

**Dependencies & config**
- [ ] New Python deps added to `requirements.txt` and (if UI) `requirements-app.txt`
- [ ] New env vars added to `.env.example` and [`docs/AGENT_CONTEXT.md`](../docs/AGENT_CONTEXT.md)
- [ ] Secrets NOT committed

**Production safety**
- [ ] If prod deploy needed: plan follows [`docs/ops/SOP_DEPLOY.md`](../docs/ops/SOP_DEPLOY.md)
- [ ] Rollback SHA identified (recorded at deploy time)
- [ ] Breaking change labelled, if applicable

**Docs**
- [ ] [`docs/AGENT_CONTEXT.md`](../docs/AGENT_CONTEXT.md) updated (if behavior changed)
- [ ] [`docs/PROJECT_OVERVIEW.md`](../docs/PROJECT_OVERVIEW.md) updated (if narrative changed)
- [ ] [`docs/PRODUCTION_ENVIRONMENT.md`](../docs/PRODUCTION_ENVIRONMENT.md) updated (if infra changed)
- [ ] New invariant added to [`docs/guardrails/INVARIANTS.md`](../docs/guardrails/INVARIANTS.md) (if applicable)

**Tests**
- [ ] At least one test added (or explicit N/A with justification)
- [ ] [`scripts/smoke_test.py`](../scripts/smoke_test.py) updated if a URL/endpoint/health indicator changed
