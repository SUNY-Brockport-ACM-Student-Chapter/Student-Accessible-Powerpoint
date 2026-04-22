# AGENTS.md — Entry Point for AI Agents

> **You are an AI agent working on this repository.** Read this file first. It points you at everything else and lists the small number of rules that must never be broken.

---

## 1. What this project is (30 seconds)

A Streamlit + RAG pipeline that takes a `.pptx` upload, generates ADA-compliant alt text (native `cNvPr/@descr` XML) and AI-enhanced speaker notes using Google Gemini + ChromaDB, and returns a downloadable accessible deck. Python 3.11. See [`docs/PROJECT_OVERVIEW.md`](docs/PROJECT_OVERVIEW.md) for the human narrative.

**Live production:** `https://access.brockportsigai.org/accessibility` (branch `Aggrement`, bare-metal systemd on GCP). See [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md).

---

## 2. Your reading order (always, in this order)

1. **This file** — invariants + routing
2. [`docs/AGENT_CONTEXT.md`](docs/AGENT_CONTEXT.md) — dense technical reference (data model, control flow, branches, foot-guns)
3. [`docs/guardrails/INVARIANTS.md`](docs/guardrails/INVARIANTS.md) — things that will silently break if violated
4. The template that matches your change type (§4 below)
5. [`docs/PROJECT_OVERVIEW.md`](docs/PROJECT_OVERVIEW.md) or [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md) — only if the task needs them

Do **not** skim — a lot of this codebase relies on non-obvious invariants (see §3).

---

## 3. Hard invariants (violation = breakage)

Full list with rationale in [`docs/guardrails/INVARIANTS.md`](docs/guardrails/INVARIANTS.md). Short form:

1. **`order_number` is the universal key** between parsed items and `.pptx` shapes. If you touch parsing, rebuilding, or any loop over `slide.shapes`/`slide.items`, preserve `order_number` exactly. Changing shape order during rebuild = alt text assigned to the wrong shape.
2. **Alt text goes on `cNvPr/@descr`** (native XML attribute), with `shape.alternative_text` as the fallback write. Screen readers use the XML attribute; `python-pptx`'s property alone is not enough.
3. **Never block the Streamlit event loop** with Gemini calls on the main thread without a spinner/progress UI — users stare at 0 % for minutes otherwise.
4. **Gemini rate-limit: sleep 60 s** on `ResourceExhausted` / 429. Do not reduce this; production will hit daily quota during a demo otherwise.
5. **ChromaDB access is via HTTP to `chroma-api`**, not direct `chromadb.Client()` from the Streamlit process. Keep the FastAPI wrapper in the path.
6. **Image normalization**: WMF/EMF + PIL "P" mode must be converted to PNG/RGB before hashing or sending to Gemini, or you get opaque base64 errors.
7. **No writes to `chroma/` inside the repo tree in production** — the prod server stores vector data at `./chroma/`. A `git clean -fdx` wipes it. Do not add it to `.gitignore` changes without reading [`docs/guardrails/INVARIANTS.md`](docs/guardrails/INVARIANTS.md) §6.
8. **Branch discipline**: `main` = clean codebase of record; `Aggrement` = what production runs; `Prod-v1` = Dockerized variant; `nextjs-impl` = experimental. Never merge these laterally without a migration plan — see [`docs/guardrails/BRANCHING.md`](docs/guardrails/BRANCHING.md).
9. **Consent gate is IRB-mandated** (Aggrement branch). Do not bypass, shorten, or remove the consent screen without an IRB amendment.
10. **Secrets never in git.** `.env` is gitignored; production's `.env` is a separate artifact.

---

## 4. Choose a template before you start coding

| Change type | Template | Typical scope |
|---|---|---|
| New capability | [`docs/templates/FEATURE.md`](docs/templates/FEATURE.md) | new UI screen, new file type, new model |
| Dependency bump / version change | [`docs/templates/UPDATE.md`](docs/templates/UPDATE.md) | `requirements*.txt`, Python version, Gemini model |
| Defect fix | [`docs/templates/BUG.md`](docs/templates/BUG.md) | incorrect alt text, crash, UX regression |
| Structural / no behavior change | [`docs/templates/REFACTOR.md`](docs/templates/REFACTOR.md) | module split, renames, typing |

Copy the template into the PR description (or the issue). Fill it out *before* writing code. The template exists so the reviewer and the next agent can reconstruct your intent in 60 seconds.

---

## 5. Pre-flight, tests, and merge gate

Before pushing:

```bash
python scripts/doctor.py       # env + layout sanity (offline, fast)
python scripts/preflight.py    # run this before opening a PR
python -m pytest -q            # unit tests
python scripts/check_invariants.py   # static invariant checks
```

After deploying to production (`docs/ops/SOP_DEPLOY.md`):

```bash
python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility
```

Full checklist: [`docs/guardrails/CHANGE_CHECKLIST.md`](docs/guardrails/CHANGE_CHECKLIST.md).

---

## 6. Operations — you are also the ops team

| Task | SOP |
|---|---|
| Ship a change to prod | [`docs/ops/SOP_DEPLOY.md`](docs/ops/SOP_DEPLOY.md) |
| Back out a bad deploy | [`docs/ops/SOP_ROLLBACK.md`](docs/ops/SOP_ROLLBACK.md) |
| Prod is down / broken | [`docs/ops/SOP_INCIDENT.md`](docs/ops/SOP_INCIDENT.md) |
| Rotate the Gemini key | [`docs/ops/SOP_SECRETS.md`](docs/ops/SOP_SECRETS.md) |

Production access requires `gcloud auth` credentials with `compute.instances.setMetadata` on `instance-20250905-023343-pub` (zone `us-central1-c`). Only three instances are SSH-accessible to current operators — see [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md) §2.

---

## 7. Things you are *not* allowed to do unprompted

- Commit without the user asking.
- Push force to any shared branch.
- `git clean -fdx` or `rm -rf chroma/` on the production VM.
- Edit files directly on the production VM (edit in git, deploy via SOP).
- Change the Gemini model without a dedicated [`UPDATE.md`](docs/templates/UPDATE.md) request.
- Add a new public nginx route without updating [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md) + firewall review.
- Remove or weaken the consent gate (IRB).
- Widen the GCP firewall.
- Expose `.env` in logs, PRs, or screenshots.

---

## 8. When the docs contradict the code

The running code is always authoritative for *behavior*; the docs are authoritative for *intent*. If you find a contradiction:

1. Read the git history on the contradicting file.
2. Open an issue with the [`docs/templates/BUG.md`](docs/templates/BUG.md) template labeled `docs-drift`.
3. Fix the smaller of (code, docs) to match the other. Prefer fixing docs unless there's clear bug evidence.

---

## 9. Quick links

- Human onboarding: [`CONTRIBUTING.md`](CONTRIBUTING.md)
- Dense tech reference: [`docs/AGENT_CONTEXT.md`](docs/AGENT_CONTEXT.md)
- Live environment: [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md)
- Invariants: [`docs/guardrails/INVARIANTS.md`](docs/guardrails/INVARIANTS.md)
- Branching: [`docs/guardrails/BRANCHING.md`](docs/guardrails/BRANCHING.md)
- Pre-merge checklist: [`docs/guardrails/CHANGE_CHECKLIST.md`](docs/guardrails/CHANGE_CHECKLIST.md)
- Ops SOPs: [`docs/ops/`](docs/ops/)
- Templates: [`docs/templates/`](docs/templates/)
- Scripts: [`scripts/`](scripts/)
