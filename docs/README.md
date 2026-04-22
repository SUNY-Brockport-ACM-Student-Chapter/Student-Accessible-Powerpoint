# Documentation Index

Start here.

| File | Audience | Purpose |
|---|---|---|
| [`PROJECT_OVERVIEW.md`](PROJECT_OVERVIEW.md) | Humans | Narrative overview of the project, architecture, branch evolution, ADA/WCAG implementation. |
| [`AGENT_CONTEXT.md`](AGENT_CONTEXT.md) | AI agents | Dense rapid-ingestion technical reference: data model, control flow, RAG pipeline, branch comparison, footguns. |
| [`PRODUCTION_ENVIRONMENT.md`](PRODUCTION_ENVIRONMENT.md) | Ops | Live GCP deployment map, differences from the repo, operator playbook. |
| [`../AGENTS.md`](../AGENTS.md) | AI agents | Top-level entry point and hard invariants. |
| [`../CONTRIBUTING.md`](../CONTRIBUTING.md) | Humans | Developer onboarding, setup, change flow. |

## `ops/` — Standard Operating Procedures

| File | When to use |
|---|---|
| [`ops/SOP_DEPLOY.md`](ops/SOP_DEPLOY.md) | Shipping a change to production. |
| [`ops/SOP_ROLLBACK.md`](ops/SOP_ROLLBACK.md) | Backing out a bad deploy. |
| [`ops/SOP_INCIDENT.md`](ops/SOP_INCIDENT.md) | Prod is broken. |
| [`ops/SOP_SECRETS.md`](ops/SOP_SECRETS.md) | Rotating or adding a secret. |
| [`ops/DEPLOY_LOG.md`](ops/DEPLOY_LOG.md) | Append-only log of prod deploys. |
| [`ops/INCIDENT_LOG.md`](ops/INCIDENT_LOG.md) | Append-only post-incident notes. |
| [`ops/SECRET_ROTATION_LOG.md`](ops/SECRET_ROTATION_LOG.md) | Append-only rotation log. |

## `guardrails/` — Architectural protection

| File | Purpose |
|---|---|
| [`guardrails/INVARIANTS.md`](guardrails/INVARIANTS.md) | Hard invariants that must not be violated, with Check commands. |
| [`guardrails/BRANCHING.md`](guardrails/BRANCHING.md) | Branch roles, cherry-pick flow, protected-branch rules. |
| [`guardrails/CHANGE_CHECKLIST.md`](guardrails/CHANGE_CHECKLIST.md) | Pre-merge checklist to paste into every PR. |

## `templates/` — Change request forms

| File | Use when... |
|---|---|
| [`templates/FEATURE.md`](templates/FEATURE.md) | Proposing / building a new capability. |
| [`templates/UPDATE.md`](templates/UPDATE.md) | Bumping a dependency or changing config. |
| [`templates/BUG.md`](templates/BUG.md) | Reporting / fixing a defect. |
| [`templates/REFACTOR.md`](templates/REFACTOR.md) | Structural change with no behavior change. |

---

## Where everything lives

```
repo/
  AGENTS.md                 - agent entry point (invariants + routing)
  CONTRIBUTING.md           - human onboarding
  docs/
    README.md               - this file
    PROJECT_OVERVIEW.md     - human narrative
    AGENT_CONTEXT.md        - agent-dense tech reference
    PRODUCTION_ENVIRONMENT.md
    ops/                    - SOPs + logs
    guardrails/             - invariants, branching, checklist
    templates/              - change request templates
  scripts/                  - doctor, preflight, check_invariants, smoke_test
  tests/                    - pytest suite (minimal)
  .github/                  - PR & issue templates, CI workflow, CODEOWNERS
```
