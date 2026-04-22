# Contributing to Student-Accessible-Powerpoint

Welcome. This document is the human onboarding path. AI agents should start at [`AGENTS.md`](AGENTS.md) instead (and then this file afterwards).

---

## 1. One-page orientation

- **What it is**: A Streamlit web app that makes PowerPoint decks ADA-compliant (WCAG 2.1 AA) by generating alt text and speaker notes via Google Gemini + a RAG pipeline over ChromaDB.
- **What lives where**: Full map in [`docs/PROJECT_OVERVIEW.md`](docs/PROJECT_OVERVIEW.md). Dense agent-oriented reference in [`docs/AGENT_CONTEXT.md`](docs/AGENT_CONTEXT.md).
- **Where it runs**: Production is on a single GCP VM under systemd. See [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md).
- **How changes flow**: local dev → PR against `main` → manual deploy to prod (we do **not** deploy `main` directly; see [Branching](docs/guardrails/BRANCHING.md)).

---

## 2. Local setup (10 minutes)

Prerequisites: Python 3.11.x, git, a Google Gemini API key.

```bash
git clone https://github.com/SUNY-Brockport-ACM-Student-Chapter/Student-Accessible-Powerpoint.git
cd Student-Accessible-Powerpoint
python -m venv venv

# Windows (PowerShell)
venv\Scripts\Activate.ps1
# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
cp .env.example .env        # then paste your GOOGLE_API_KEY
python scripts/doctor.py    # sanity check your environment
```

Run the full stack (three processes):

```bash
# Terminal 1 — ChromaDB (vector store)
venv/bin/chroma run --path ./chroma

# Terminal 2 — FastAPI wrapper
python app/chroma-api/app.py

# Terminal 3 — Streamlit UI
streamlit run app/ppt_notes.py
```

Open http://localhost:8501 . Upload any `.pptx`. If it crashes, check the three terminals — the Gemini key is the usual culprit.

> **Shortcut**: `python start_app.py` starts Chroma + Streamlit together (dev only; production uses systemd).

---

## 3. Making a change

1. **Pick a change type** and read the matching template:
   - New capability → [`docs/templates/FEATURE.md`](docs/templates/FEATURE.md)
   - Version bump / dep update → [`docs/templates/UPDATE.md`](docs/templates/UPDATE.md)
   - Bug fix → [`docs/templates/BUG.md`](docs/templates/BUG.md)
   - Restructure without behavior change → [`docs/templates/REFACTOR.md`](docs/templates/REFACTOR.md)

2. **Branch** from `main`:

   ```bash
   git checkout main && git pull
   git checkout -b feat/<short-slug>     # or fix/, chore/, refactor/
   ```

3. **Read the invariants** in [`docs/guardrails/INVARIANTS.md`](docs/guardrails/INVARIANTS.md) *before* touching parsing, rebuild, or Gemini calls.

4. **Code**. Favor small PRs. No new dependency without adding it to both `requirements.txt` and `requirements-app.txt` if the Streamlit UI uses it.

5. **Run the pre-merge gate**:

   ```bash
   python scripts/doctor.py
   python scripts/check_invariants.py
   python -m pytest -q
   python scripts/preflight.py
   ```

6. **Open a PR** against `main`. The PR template will prompt you to paste the filled-out change template.

7. **Deploy** only after merge. Follow [`docs/ops/SOP_DEPLOY.md`](docs/ops/SOP_DEPLOY.md). The deploy is manual and gated on the smoke test.

---

## 4. Style & conventions

- **Python**: 3.11. Type hints on new public functions. `snake_case`. Docstrings on any function that crosses a module boundary.
- **No lateral commits to `Aggrement` or `Prod-v1`** without an accompanying cherry-pick plan — those are deployment branches, not feature branches. Details: [`docs/guardrails/BRANCHING.md`](docs/guardrails/BRANCHING.md).
- **No comments that narrate the code.** Comments explain *why*, not *what*.
- **Env vars**: new variables go in `.env.example` with a placeholder and in [`docs/AGENT_CONTEXT.md`](docs/AGENT_CONTEXT.md) "Stable facts" section.
- **Pydantic models** in `app/models/models.py` are load-bearing — treat them like a schema. Add fields, do not rename.

---

## 5. Testing

The codebase has historically had no automated tests. We are adding them incrementally; see `tests/` and [`scripts/check_invariants.py`](scripts/check_invariants.py). **Every new PR should add at least one test for the change it introduces.** If the change is literally untestable, say so in the PR description.

Minimum bar:

- Unit test for any pure function you add or change.
- Smoke test update (`scripts/smoke_test.py`) if you add a public URL or endpoint.

---

## 6. Reporting a bug or requesting a feature

- **Bug**: use [`docs/templates/BUG.md`](docs/templates/BUG.md) or the GitHub "Bug" issue form.
- **Feature**: use [`docs/templates/FEATURE.md`](docs/templates/FEATURE.md) or the "Feature" issue form.

Templates are agent-optimized: future AI contributors will read the filled-out template as their primary source of truth for the task.

---

## 7. Getting help

- Open an issue with as much of the relevant template filled out as you can.
- Operator questions (prod, credentials, GCP): see [`docs/PRODUCTION_ENVIRONMENT.md`](docs/PRODUCTION_ENVIRONMENT.md) §7.
- Accessibility questions: see [`docs/PROJECT_OVERVIEW.md`](docs/PROJECT_OVERVIEW.md) § "ADA / WCAG implementation".
