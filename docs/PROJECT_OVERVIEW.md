# Student-Accessible-Powerpoint — Project Overview

A human-readable guide to the project's purpose, architecture, branch-by-branch evolution, and ADA Title II / WCAG accessibility implementation.

> **Companion files:**
> - [`AGENT_CONTEXT.md`](./AGENT_CONTEXT.md) — dense technical reference for AI agents.
> - [`PRODUCTION_ENVIRONMENT.md`](./PRODUCTION_ENVIRONMENT.md) — live GCP deployment map.
> - [`README.md`](./README.md) — index of all docs.
> - [`../CONTRIBUTING.md`](../CONTRIBUTING.md) — how to set up, change, and ship the code.

---

## 1. What This Project Does

Student-Accessible-Powerpoint (originally piloted at the SUNY Brockport ACM Student Chapter) is a tool that transforms standard `.pptx` presentations into **ADA Title II / WCAG 2.1 AA compliant** documents by:

1. **Extracting** every text run, image, chart, table, and background fill from a deck.
2. **Indexing** the extracted content into a **ChromaDB** vector collection so each image can be described with awareness of its surrounding slide and of the entire presentation.
3. **Generating alt text** for every image using **Google Gemini** (multi-stage RAG pipeline: OCR → enhanced description → Lambda-Index context retrieval → final description).
4. **Generating comprehensive speaker notes** for each slide that summarize the slide in Markdown and describe all visual content — so a student using a screen reader receives the same information as a sighted student.
5. **Writing the results back** into native PowerPoint XML fields (`cNvPr/@descr` for alt text, `notesSlide` for notes) so the output is a standard `.pptx` that Office, Google Slides, and screen readers all understand without any custom tooling.

The end user uploads a PPTX in a Streamlit UI, (optionally) reviews AI-generated descriptions, and downloads an accessibility-enhanced file.

---

## 2. Architecture at a Glance

```
┌─────────────────────────┐      ┌──────────────────────────┐      ┌──────────────────────┐
│  Streamlit UI           │──────▶│  Accessibility Pipeline  │──────▶│ Google Gemini API    │
│  (app/ppt_notes.py)     │      │  (pptx_rag_quizzer/*)    │      │ gemini-2.0-flash-lite│
│  - upload / batch review│      │  - parse_powerpoint      │      └──────────────────────┘
│  - download             │      │  - ImageProcessor (RAG)  │                  ▲
└─────────────────────────┘      │  - rebuild_presentation  │                  │
           │                     └──────────────────────────┘                  │ image + prompt
           │                                │                                  │
           │                                ▼                                  │
           │                     ┌──────────────────────────┐                  │
           │                     │ ChromaDB HTTP Wrapper    │◀─────────────────┘
           │                     │ (FastAPI, port 8001)     │
           │                     │  app/chroma-api/app.py   │
           │                     └──────────────────────────┘
           │                                │
           │                                ▼
           │                     ┌──────────────────────────┐
           │                     │ ChromaDB Server (8000)   │
           │                     │ persistent volume        │
           │                     └──────────────────────────┘
```

### Three processes, one deployment
- **`web`** — Streamlit app (`app/ppt_notes.py`), port 8501.
- **`chroma-api`** — Thin FastAPI wrapper over ChromaDB, port 8001. It exists so the web app never talks to ChromaDB's native client directly, making the app deployable without the heavy `chromadb` Python dependency.
- **`chroma`** — ChromaDB server, port 8000, persistent volume.

The Caddy reverse proxy (production branches) terminates TLS in front of `web`.

### Major modules

| Module | Responsibility |
|---|---|
| `app/ppt_notes.py` | Streamlit UI; orchestrates the 4-stage user flow (upload → describe → final processing → download). |
| `app/pptx_rag_quizzer/utils.py` | `parse_powerpoint()` — extracts text + images from a `.pptx`; `convert_image_to_png_or_jpg()` — normalizes exotic formats (WMF/EMF/SVG) via ImageMagick. |
| `app/pptx_rag_quizzer/rag_core.py` | `RAGCore` — talks to Chroma via HTTP; `prompt_gemini()` and `prompt_gemini_with_image()` with quota/retry handling. |
| `app/pptx_rag_quizzer/image.py` | `Image` processor — 4-stage RAG image-description pipeline with caching, Lambda-Index ranking, and chat-history continuity. |
| `app/models/models.py` | Pydantic data models: `Presentation`, `Slide`, `SlideItem`, `Image`, `Text`. |
| `app/chroma-api/app.py` | FastAPI service exposing REST endpoints for collections, documents, and queries. |
| `start_app.py` | Developer helper that starts the ChromaDB API and Streamlit app together. |

---

## 3. Branch-by-Branch Evolution

The repository's history reads as an incremental hardening of a single product. Each branch represents a milestone.

### 🟢 `main` — current stable (Apache 2.0)
- Core pipeline: parse → build RAG text collection → per-batch image description (Gemini) → rebuild collection with image descriptions → write alt text + generate RAG-enhanced notes → save output PPTX.
- Two-path UI: full review workflow **or** `Quick Generate & Download (Skip Review)` for speed.
- HTTP-only Chroma access (no native `chromadb` dependency on the web process).
- `requirements-app.txt` introduced as a slimmer install target for the Streamlit process.
- `2691d92` added native PPTX alt-text attribute (`cNvPr/@descr`) with a `alternative_text` fallback — this is the canonical commit for ADA-grade alt-text plumbing.
- `898c490` added the `Quick Generate & Download` option.

### 🟠 `RAG-integration-branch` — the original Dockerized release
- First "production-ish" deployment, built on a single `Dockerfile` and a `docker-compose.yml` exposing port 8501.
- `DEPLOYMENT.md` documents the Debian-bookworm VM setup (Docker CE + Docker Compose plugin).
- This is the heritage branch — everything on `main` descends from it.

### 🟣 `Prod-v1` — multi-service production deployment
Adds:
- **Split Dockerfiles**: `Dockerfile.api` (FastAPI Chroma wrapper) and `Dockerfile.web` (Streamlit with Tesseract + libjpeg).
- **`docker-compose.yml` with four services**: `chroma`, `chroma-api`, `web`, `caddy`.
- **Caddy reverse proxy** with ACME/Let's Encrypt for `access.brockportsigai.org`.
- Introduces inter-container networking (`CHROMA_API_URL=http://chroma-api:8001`, `CHROMA_SERVER_HOST=chroma`).

This is the blueprint for running the system on a cloud VM with TLS, persistent Chroma storage, and independent scaling of the three tiers.

### 🟡 `Aggrement` — ADA/IRB consent gate + feature expansion (most feature-rich)
The most ambitious branch. Adds:

- **IRB/Research Consent Wall** (`c14117b`) — a mandatory Streamlit gate that collects informed-consent responses before the user can upload a file. This is required because the tool is used as an educational-research instrument at SUNY Brockport. Consent decisions (and email, when volunteered) are appended to a local `consent_responses.csv`.
- **Under-18 block stage** — hard-stops the workflow for minors.
- **AI-content warning** — Streamlit banner reminds the instructor that AI output may contain errors.
- **Modular pipeline split** — `utils.py` is broken into:
  - `utils.py` — OCR + image-format conversion only.
  - `pptx.py` — `parse_powerpoint()` + `generate_accessible_notes()` + `rebuild_presentation_with_accessible_features()`.
  - `word.py` — **new: DOCX parsing & alt-text rebuilder** (`parse_word_document`, `rebuild_word_document_with_accessible_features`). Word support uses `python-docx` and walks paragraphs, tables, and nested cells, applying alt text to `wp:docPr/@descr`.
- **Expanded Pydantic models** — `WordDocument`, `WordSection`, `WordText`, `WordImage` alongside the original presentation models; `Image.image_bytes` is now base64-serialized via `@field_serializer` so the model is JSON-round-trippable.
- **Major accessibility & performance improvements** (`6ba2cb8`):
  - Reuse a single `RAGCore()` instance across all slides (~70 % speedup, per `CATCH_UP.md`).
  - Token budget raised 200 → 400 for non-truncated notes.
  - Post-processing strips AI "Okay, here are…" preambles.
  - Native `cNvPr/@descr` alt-text path with `alternative_text` fallback.
  - Recursive shape processing (groups, diagrams, charts, background fills).
  - Order-number tracking fixed so text shapes also increment the index used to match images back to their parsed counterparts.
  - Graceful WMF/EMF failure: `convert_image_to_png_or_jpg()` returns `(None, None)` and the pipeline skips the image instead of crashing.
  - PIL palette-with-transparency warning fixed by explicit `P → RGBA → RGB on white background` conversion.
- **Streamlit config** — `app/.streamlit/config.toml` sets `maxUploadSize = 500 MB` and a `baseUrlPath = "accessibility"` for path-based hosting behind Caddy.
- **`CATCH_UP.md`** — detailed chronological change log of every fix/feature in the branch (the most thorough single document in the repo).

### 🔵 `nextjs-impl` — experimental TS/React port
- Full migration to **Next.js 16** + **React 19** + **TypeScript** + **Tailwind CSS 4**.
- Uses `adm-zip` + `xml2js` to parse PPTX directly from the `ppt/slides/*.xml` entries (no `python-pptx`).
- A `GeminiService` class calls `gemini-2.0-flash:generateContent` via `axios`.
- A `ChromaService` class mirrors the Python HTTP client.
- The Python implementation is preserved under `backups/python-legacy/` for reference.
- Status: early — the `/api/process` route returns slide metadata only; the rebuild/write-back pathway is not yet ported.

### Branch evolution, summarized

| Milestone | Branch |
|---|---|
| First RAG-enhanced accessibility pipeline + Docker | `RAG-integration-branch` |
| Alt-text & Options merged to main | `main` (PRs #1, #2) |
| Multi-service Docker + Caddy TLS | `Prod-v1` |
| IRB consent, DOCX support, perf + robustness, CATCH_UP log | `Aggrement` |
| TypeScript / Next.js rewrite (experimental) | `nextjs-impl` |
| Apache 2.0 license added | `main` (commit `1a31150`) |

---

## 4. ADA Title II / WCAG Compliance — How It's Implemented

**ADA Title II** (updated 2024 regulations) requires state and local government entities — including public universities — to make digital content conform to **WCAG 2.1 Level AA** by April 2026/2027. This project targets the WCAG success criteria most commonly violated by slide decks.

### 4.1 Non-text content (WCAG 1.1.1 — Level A)
**Every image gets a real alt-text description, not a placeholder.**

- Parsing: `parse_powerpoint()` extracts every image blob — regular pictures, diagrams, charts, grouped shapes (recursively), and background fills — and normalizes it to PNG/JPG via ImageMagick (`convert_image_to_png_or_jpg`). WMF/EMF failures are logged and skipped rather than crashing the whole deck.
- Description generation (`pptx_rag_quizzer/image.py` — 4 stages):
  1. **OCR** via Tesseract on any text baked into the image.
  2. **Enhanced description** — Gemini receives the image + OCR text + the rest of the slide's text context, producing a 1–3 sentence visual-primary description.
  3. **Lambda-Index context retrieval** — key terms from the enhanced description query Chroma across the whole presentation, then results are re-ranked by term overlap × image-metadata relevance.
  4. **Final description** — Gemini refines the description using the retrieved cross-slide context and a rolling chat history so descriptions remain consistent across similar images in the deck.
- Writing: `update_images_with_alt_text()` sets the native PPTX XML attribute:
  ```python
  shape._element._nvXxPr.cNvPr.attrib["descr"] = alt_text
  ```
  With `shape.alternative_text = alt_text` as a fallback. This is the attribute every major screen reader (JAWS, NVDA, VoiceOver, Narrator) reads, so descriptions survive export to PDF, Google Slides, and Keynote.

### 4.2 Programmatically determined info & relationships (WCAG 1.3.1 — Level A)
- **Speaker notes** (the `notesSlide` part of the PPTX) are regenerated for every slide with a Markdown-formatted summary that includes:
  - Slide title (`## Slide N: …`)
  - Key concepts as bullet points
  - Visual-element descriptions (so a blind student reading the notes knows what was on screen)
- Charts and tables receive their own descriptive alt text (`Chart on slide N — {title}`, `Table on slide N — {title}`).
- Reading order is preserved by sorting shapes top-to-bottom, left-to-right (`sorted(shapes, key=lambda x: (x.top, x.left))`) before extraction.

### 4.3 Language of parts & consistent voice (WCAG 3.1.x)
- A rolling `chat_history` (max 10 entries) is passed into Gemini on each image so descriptions stay consistent across a 50-slide deck (e.g., the same diagram style is described with the same vocabulary each time).

### 4.4 Robust, machine-readable output (WCAG 4.1.x)
- All metadata is written into standard OOXML elements (`cNvPr/@descr`, `notesSlide`, `descr` on `wp:docPr` for DOCX on the `Aggrement` branch). No custom namespaces, no proprietary extensions.
- `Aggrement/app/.streamlit/config.toml` raises `maxUploadSize` to 500 MB so long instructor decks are not silently truncated.

### 4.5 Reasonable-use safeguards
- The `Aggrement` branch adds an **AI-content warning banner** prompting the instructor to verify AI output before distribution — important because WCAG conformance requires *accurate* text alternatives, and an uncorrected hallucination could itself be a violation.
- The 4-stage `describe_images` batch UI exists specifically so a human can review and edit every description before it is written to the file.
- The IRB consent gate ensures the research-deployment context complies with human-subjects rules at SUNY Brockport.

### 4.6 What the project does *not* yet cover (known gaps)
- **Color contrast** (WCAG 1.4.3) is not inspected or corrected.
- **Live captions / audio descriptions** are not generated for embedded media.
- **Reading-order reordering** only affects extraction, not the on-slide tab order shown in the Accessibility Inspector.
- **Hyperlink descriptive text** (WCAG 2.4.4) is not rewritten.

These are reasonable roadmap items and are called out in the RAG README under "Future Enhancements".

---

## 5. Running the Project

### Local development (on `main`)
```bash
pip install -r requirements.txt
# .env must contain GOOGLE_API_KEY
python start_app.py      # starts Chroma API + Streamlit together
```

### Production (on `Prod-v1`)
```bash
docker compose up --build -d     # starts chroma, chroma-api, web, caddy
# Visit https://access.brockportsigai.org
```

### Research deployment with consent gate (on `Aggrement`)
Same as `Prod-v1`, but the Streamlit app first requires the user to complete the consent radio-form; responses are appended to `consent_responses.csv` on the container volume.

### Environment variables

| Variable | Purpose | Default |
|---|---|---|
| `GOOGLE_API_KEY` | Gemini auth | **required** |
| `CHROMA_API_URL` | URL to the FastAPI wrapper | `http://localhost:8001` |
| `CHROMA_SERVER_HOST` | Where `chroma-api` finds the ChromaDB server | `localhost` |
| `CHROMA_SERVER_HTTP_PORT` | ChromaDB port | `8000` |
| `API_HOST` / `API_PORT` | FastAPI wrapper bind | `0.0.0.0` / `8001` |
| `ACME_EMAIL` | Caddy TLS contact (Prod-v1) | — |

---

## 6. Repository Layout (current `main`)

```
Student-Accessible-Powerpoint/
├── app/
│   ├── ppt_notes.py                  # Streamlit UI & orchestrator
│   ├── chroma-api/
│   │   ├── app.py                    # FastAPI wrapper over ChromaDB
│   │   └── .env.example
│   ├── models/
│   │   └── models.py                 # Pydantic models
│   └── pptx_rag_quizzer/
│       ├── utils.py                  # parse_powerpoint, OCR, image convert
│       ├── rag_core.py               # Gemini + Chroma HTTP client
│       └── image.py                  # 4-stage RAG image description
├── docs/
│   ├── PROJECT_OVERVIEW.md           # this file
│   └── AGENT_CONTEXT.md              # rapid-ingestion AI-agent context
├── start_app.py                      # dev launcher
├── requirements.txt                  # full (web + api + chroma)
├── requirements-app.txt              # Streamlit-only dependencies
├── RAG_INTEGRATION_README.md         # narrative RAG feature doc
├── LICENSE                           # Apache 2.0
├── .env.example
├── .gitattributes
└── .gitignore
```

---

## 7. Further Reading
- [`../RAG_INTEGRATION_README.md`](../RAG_INTEGRATION_README.md) — the original feature-centric README for the RAG workflow.
- [`./AGENT_CONTEXT.md`](./AGENT_CONTEXT.md) — the structured reference for AI coding agents.
- Branch `Aggrement`'s `CATCH_UP.md` — a detailed chronological change log of every fix/optimization in the most advanced branch.
- U.S. DOJ ADA Title II rule (2024), 28 CFR Part 35 — the legal basis for WCAG 2.1 AA conformance this project targets.
