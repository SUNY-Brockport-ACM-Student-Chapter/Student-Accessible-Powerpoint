# Branching Policy

The repo has several long-lived branches that each mean something different. Cross-merging them blind has broken things. This file is the policy.

---

## 1. The branches

| Branch | Role | Owns... |
|---|---|---|
| `main` | Canonical codebase. The "source of truth" for the project's design. | Core app, models, RAG core, Streamlit UI |
| `Aggrement` | **What production runs.** Adds the IRB consent gate and modular split (`pptx.py` + `word.py`). | Consent wall, Word support |
| `Prod-v1` | Dockerized deployment model. *Not actually deployed* — see [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §6. | `docker-compose.yml`, Caddy config |
| `nextjs-impl` | Experimental Next.js/React rewrite of the frontend. Not shipped. | `frontend/` Next.js app |
| `RAG-integration-branch` | Legacy RAG prototype. Archive. | Historical reference only |

---

## 2. Where your change goes

Decide before you branch.

| Change type | Land on | Then forward to |
|---|---|---|
| Bug fix in core RAG, parsing, UI, models | `main` | Cherry-pick or merge into `Aggrement` for prod |
| New feature in core pipeline | `main` | Cherry-pick into `Aggrement` after design review |
| Consent-gate / IRB / research-only logic | `Aggrement` | Nowhere — stays on `Aggrement` |
| Docker / Caddy / compose changes | `Prod-v1` | Nowhere — `Prod-v1` is currently dormant |
| Next.js frontend experiments | `nextjs-impl` | Nowhere until the rewrite is the plan of record |

**Never merge `Aggrement` → `main`** directly. It carries IRB-specific code that doesn't belong in the canonical branch. Use `git cherry-pick` for the core bits that do.

**Never merge `Prod-v1` → `main`** directly. It carries deployment artifacts that don't belong in the app tree.

---

## 3. Branch naming for short-lived work

```
feat/<slug>        new capability
fix/<slug>         bug fix
chore/<slug>       tooling / docs / non-code
refactor/<slug>    structural change, no behavior change
deps/<slug>        dependency bump
```

Examples: `feat/word-consent-wall`, `fix/order-number-groups`, `refactor/move-chroma-outside-repo`.

---

## 4. The cherry-pick flow (main → Aggrement)

When a fix on `main` needs to reach production:

```bash
git checkout Aggrement
git pull
git cherry-pick <sha-on-main>      # or --no-commit for a range
# resolve any conflicts (the modular split in Aggrement often causes them)
git push origin Aggrement
```

Then open a PR against `Aggrement` — even cherry-picks get a PR so the deploy log has a record. After merge, follow [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md).

---

## 5. The migration decision (when do we collapse branches?)

Long-term, having `main` ≠ production is a liability. Two acceptable end-states:

- **Option A (recommended):** promote `Aggrement` as the default branch, demote `main` to archival. Rename and update all docs.
- **Option B:** make `main` the source of truth, carry the consent gate via a feature flag. Retire `Aggrement`.

Either migration is a `REFACTOR.md`, not a one-liner. Do not start it without buy-in from the course / research owner.

---

## 6. What the production VM sees

Production runs `Aggrement`. The VM's `git status` historically shows uncommitted edits (see [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §3.4). Agent contract: **you do not commit on the VM.** All commits happen via PR; the VM only runs `git pull`. If you see drift, capture it with [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md) §1.

---

## 7. Protected-branch rules (to be configured)

On GitHub, `main` and `Aggrement` should have:

- Require PR before merging.
- Require at least 1 approving review.
- Require status checks (once CI exists): `preflight`, `tests`, `check_invariants`.
- Disallow force-push.
- Disallow deletion.

These are not currently enforced at the repo level. Configure them.
