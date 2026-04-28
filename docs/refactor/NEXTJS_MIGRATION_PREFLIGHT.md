# Next.js Migration — Preflight Discovery

> **Status:** Draft. Questions only, no answers yet. Awaiting human review before any code moves.
> **Branch:** `refactor/nextjs-migration`
> **Scope:** Port the Streamlit + FastAPI + ChromaDB stack described in `docs/PROJECT_OVERVIEW.md`, `docs/AGENT_CONTEXT.md`, and `docs/PRODUCTION_ENVIRONMENT.md` onto the Next.js 16 App Router scaffold that already exists on the `nextjs-impl` branch.
> **Hard constraints carried over:** every invariant in `docs/guardrails/INVARIANTS.md`. Nothing below may be answered in a way that violates them without an explicit exception in that file.

This is a **discovery document**, not a design. Every section below is a question or a set of questions that need to be answered before any .tsx or .py file on this branch is touched. Each question is phrased so that the answer forces a concrete architectural decision.

---

## 0. Reading order

Answer the sections in this order. Later sections assume earlier ones are settled.

1. [Scope & End-State](#1-scope--end-state) — what "done" means.
2. [Stage → Component Mapping](#2-stage--component-mapping) — UI architecture.
3. [Backend Surface Mapping](#3-backend-surface-mapping) — where Python logic goes.
4. [PPTX Parsing & Rebuild Strategy](#4-pptx-parsing--rebuild-strategy) — the riskiest port.
5. [RAG Pipeline Port](#5-rag-pipeline-port) — Gemini + Chroma over the wire.
6. [State, Session, and Persistence](#6-state-session-and-persistence) — the biggest Streamlit-shaped hole.
7. [Long-Running Work & File Uploads](#7-long-running-work--file-uploads) — timeout math.
8. [Invariant Translation Matrix](#8-invariant-translation-matrix) — line-by-line.
9. [Consent Gate (IRB)](#9-consent-gate-irb) — the non-negotiable.
10. [Auth & Identity](#10-auth--identity) — new problem Streamlit never solved.
11. [Production Deployment](#11-production-deployment) — replacing systemd bare-metal.
12. [Coexistence with the Chroma Service](#12-coexistence-with-the-chroma-service) — keep or replace?
13. [Existing `nextjs-impl` Scaffold Audit](#13-existing-nextjs-impl-scaffold-audit) — what to keep, what to burn.
14. [Cutover, Rollback, Data Migration](#14-cutover-rollback-data-migration) — flip-day plan.
15. [Testing, Accessibility QA, Observability](#15-testing-accessibility-qa-observability) — how we'll know.
16. [Framework Integration](#16-framework-integration) — fitting this into the ops framework we just built.

---

## 1. Scope & End-State

The answers here become the non-negotiable assumptions for every later section.

1.1. Is the Next.js app the **replacement** for the Streamlit deployment at `https://access.brockportsigai.org/accessibility`, or does it coexist on a new subpath (e.g. `/v2`) during a parallel-run window? If coexist, for how long and who owns the cutover signal?

1.2. Which branch is the source of behavior truth for v2 parity — `main`, `Aggrement` (currently live, carries the IRB consent gate), or `Prod-v1` (documented Docker stack, not actually running)? If the answer is `Aggrement`, the IRB consent flow (invariant #8) is in scope on day 1.

1.3. What is the **hard feature freeze** for v1 Streamlit during the migration? If Streamlit keeps accepting changes, which changes are in scope to re-port and which are frozen?

1.4. Are we targeting a Node-only runtime, or is there room to keep `app/chroma-api/` (FastAPI) running in place and only port the UI + orchestration? This one question forks the entire plan.

1.5. What are the product-level KPIs the migration must not regress (upload-to-download latency, max deck size, Gemini cost per deck, concurrent users)? Which of these does v1 even measure today?

1.6. What is the planned sunset policy for the Streamlit install on the prod VM — hard stop, or read-only archive?

---

## 2. Stage → Component Mapping

The current UI lives in `app/ppt_notes.py :: main()` and is driven by four values of `st.session_state.processing_stage`: `upload` → `describe_images` → `final_processing` → `download`. Every transition is a server-rendered re-execution of the whole script. Next.js has no equivalent; decisions here determine whether the app is server-first or client-first.

2.1. Which of the four stages should be a **React Server Component** (static shell, no interactivity) vs. a **Client Component** (stateful, event handlers)? Candidate split to confirm or reject:

| Stage | Candidate RSC boundary | Candidate Client boundary |
|---|---|---|
| `upload` | Page shell, copy, branding, stepper skeleton | File picker, drag-drop handler, progress UI |
| `describe_images` | None — fully interactive | Entire batch-review UI (`<textarea>` edits, next/prev buttons, per-image preview) |
| `final_processing` | Status polling shell | "Generate" button + progress indicator |
| `download` | Result summary (server-rendered from job record) | "Download" click handler |

2.1.a. Does every stage need to survive a full page refresh (i.e. stage stored server-side, keyed by a URL param or cookie) or is a refresh-wipes-progress UX acceptable for v1?

2.1.b. The current "Quick Generate & Download" button bypasses review entirely. Does that become a separate `/api/quick-process` route or a toggled code path inside `/api/process`?

2.2. Should the stage state live in the **URL** (e.g. `/process/[jobId]/review`), in a **server-side session**, in **React Context** inside one client component, or in a client-state library (`zustand`, Jotai, Redux Toolkit)? Rank the trade-offs with respect to:
- Back button behavior (Streamlit currently offers a "← Back" button; URL-driven state matches browser history for free).
- Shareable / resumable job links (supporting accessibility coordinators sharing WIP jobs).
- Refresh survivability.

2.3. The existing `nextjs-impl/src/app/page.tsx` is a single 300-line client component that uses `useState<Stage>` and a hard-coded review mockup. Is that prototype the **starting point** we iterate, or do we **discard it** and build from the stage map in 2.1? (If we keep it, item 2.1's server-first split is dead.)

2.4. Does the stepper in the current nextjs-impl prototype stay as the v2 IA, or do we move to a **segmented route** (`/upload`, `/review/[jobId]`, `/result/[jobId]`) so each stage gets its own URL? Segmented routes are the App Router idiom; keeping a single `/` page is not.

2.5. Is **Server-Sent Events** or a WebSocket needed for live progress during batched image description (current code does one `st.write(...)` per image), or is polling the job status from a client component acceptable for v1? Streamlit's implicit streaming is the hardest Streamlit-ism to reproduce.

2.6. Should review edits to alt-text be saved via a **Server Action** (`'use server'` function called from the client) or a **Route Handler** POST? We need to pick one pattern per write path for the whole app and write it down; today's prototype uses only route handlers.

2.7. Which shared UI primitives in the Streamlit app have no direct mapping and therefore require a design decision before building: `st.progress`, `st.spinner`, `st.image` (auto-resize), `st.download_button` (blob streaming), `st.error` (toast vs. inline)?

2.8. What is the accessibility baseline for the chrome itself (stepper, buttons, file picker, progress)? We are shipping an accessibility tool; WCAG 2.1 AA on the UI is table stakes — who signs off, and against what checklist?

---

## 3. Backend Surface Mapping

Today there are two Python entry points the UI talks to:

- `RAGCore` (in-process) — parse, Gemini calls, collection build, context retrieval.
- `chroma-api` (FastAPI on `:8001`) — REST wrapper around ChromaDB.

Each Python function needs a Next.js home. The matrix below is what needs to be filled in by the answer to this section:

| Current Python entry | Callers | Proposed Next.js home | Runtime |
|---|---|---|---|
| `parse_powerpoint(file_object, file_name)` in `app/pptx_rag_quizzer/utils.py` | upload stage | Route Handler `POST /api/process` or Server Action | **node** (not edge) |
| `RAGCore.create_collection(Presentation)` | upload stage | ? | ? |
| `ImageProcessor.describe_image(bytes, ext, slide_num, collection_id)` | describe_images stage | ? | ? |
| `RAGCore.remove_collection(id)` + re-create | final_processing | ? | ? |
| `process_powerpoint_with_rag_enhanced(...)` | final_processing | ? | ? |
| `ExtractText_LLM(image_base64, ...)` in `app/ppt_notes.py` | image description path | merges with `ImageProcessor`? | ? |
| `generate_enhanced_notes_with_context(...)` | final_processing | ? | ? |

3.1. Each row above: **Route Handler, Server Action, or background worker**? Server Actions are ergonomic but have a ~1 MB payload cap and a 30–60 s Vercel/Node-timeout window that WILL be exceeded on a 50-slide deck with image OCR + Gemini describe + Chroma rebuild.

3.2. Every row that calls Gemini or Chroma must run on the **Node runtime** (not Edge) because:
- `google-generativeai`'s equivalent JS SDK (`@google/generative-ai`) needs Node APIs.
- `adm-zip` in the existing scaffold is Node-only.
- `axios` + HTTP to the Chroma wrapper is fine either way, but see 12.x on whether that wrapper survives.

Confirm that no part of this app is a candidate for the Edge runtime. If any part is, identify it now.

3.3. Do we keep the multi-module layout (`lib/pptx-utils`, `lib/gemini`, `lib/chroma`) the existing scaffold started, or do we introduce a domain layer (`lib/domain/presentation.ts`, `lib/domain/alt-text.ts`) that mirrors `app/models/models.py`? The Pydantic invariant (`order_number`, `SlideItem` hierarchy) has to be honored by whatever TS types we pick — see section 4.4.

3.4. Is there a "BFF" boundary we want to enforce — e.g. all Gemini and Chroma keys stay strictly inside route handlers / server actions and are never shipped to the client? Today Streamlit effectively enforces this by running on the server; in Next.js this has to be a written rule.

3.5. What is the error contract between handlers and the UI? The current code leaks Python exceptions into `st.error(...)`; we need a typed error shape (`{ code, message, retryable, jobId? }`) and a question about whether to adopt `neverthrow`, `zod`-validated responses, or plain typed unions.

3.6. Do we want **streaming responses** (React Server Components with `Suspense` + `loading.tsx` boundaries) for the review stage, where image descriptions trickle in as Gemini finishes each one? This is the most natural Next.js rewrite of the batched `st.write(...)` loop but requires commitment to the App Router streaming model.

---

## 4. PPTX Parsing & Rebuild Strategy

This is the riskiest port. `python-pptx` has no faithful Node equivalent, and the current scaffold's `lib/pptx-utils.ts` reimplements *some* of it via `adm-zip` + `xml2js` — well enough to read text and image relationships, but **not** well enough to write back `cNvPr/@descr`, handle group shapes, or manage speaker notes at XML fidelity (line 146 of that file is literally a `console.log` placeholder for alt-text writeback).

4.1. Do we **keep** a Python process for parsing + rebuild (e.g. a FastAPI service co-located with the Next.js app or the existing chroma-api), or **fully port** to TypeScript?

4.1.a. If port: which library? The candidates are:
- `officegen` — write-only, does not read existing decks.
- `pptxgenjs` — write-only, generative, not round-trip.
- `adm-zip` + `xml2js` + hand-rolled XML (current scaffold). Requires us to own the rebuild code path including group-shape recursion (invariant #7) and `cNvPr/@descr` write (invariant #2).
- `docx4js` / `node-pptx` — review state and maintenance risk.
- Shelling out to LibreOffice headless — ports deck fidelity but explodes the runtime.

4.1.b. If keep Python: do we invoke it via (a) a sidecar FastAPI service the Next.js route handlers call, (b) `child_process.spawn('python', ...)` per request, or (c) a managed worker queue (BullMQ + a Python consumer)?

4.2. Who owns **re-implementing invariant #1** (`order_number` as the universal join key)? In Python this is set in `parse_powerpoint` by incrementing `order_number` in the shape loop. The scaffold's `PptxProcessor` currently uses `id: 'text-${slideNumber}-full'` / `id: 'img-${slideNumber}-${relId}'` with no `order_number` field. This will silently misalign alt text during rebuild. What type contract replaces `app/models/models.py :: SlideItem.order_number` in TypeScript, and where is it enforced (zod schema? branded type?)?

4.3. Who owns **re-implementing invariant #2** (`cNvPr/@descr` write)? What's the plan for:
- Locating `<p:nvPicPr><p:cNvPr/>` under each picture in `slideN.xml`.
- Setting or replacing the `descr=` attribute.
- Preserving the rest of the XML (namespaces, xml:space, comments).
- Also setting python-pptx's fallback `alternative_text` equivalent (the `a:blip` / descrElement pair — or just `cNvPr/@descr` in the TS port).
- Writing the .pptx back as a valid OOXML zip (central directory, compression, content types).

4.4. TypeScript type fidelity for `Presentation / Slide / SlideItem / Image / Text`: generate from a zod schema? Handwrite? Share via a generated package? Whichever answer, it must include `order_number: number` as a required, non-defaulted field (invariant #1).

4.5. Image normalization (invariant #5) — WMF/EMF → PNG, paletted → RGB. The current scaffold has none of this. What TS equivalent of `Pillow.convert('RGB')` and the WMF branch do we use (sharp + ImageMagick shell? `@resvg/resvg-js`? `imagemagick-native`?) and does it run on the same Node process or a separate service?

4.6. Group-shape recursion (invariant #7) — `xml2js` flattens `<p:grpSp>` differently from python-pptx's shape tree. Define the recursion in TS, including test fixtures covering grouped images and nested groups.

4.7. Chart and table alt-text (currently written as `shape.alternative_text` in Python) — are these in scope for v1 or explicitly deferred?

4.8. What's the max deck size we promise to handle? 50 slides × ~5 images × ~2 MB images = ~500 MB in memory on the server per upload. Does that fit in a single Node process, or do we need to stream / chunk from upload → disk → worker?

---

## 5. RAG Pipeline Port

Today `RAGCore` (in `app/pptx_rag_quizzer/rag_core.py`) orchestrates four things:

- Create a ChromaDB collection, add per-slide combined text.
- Query the collection by slide number or by query text.
- Call Gemini (text + image-with-prompt) with a 60 s `ResourceExhausted` backoff.
- Cache LLM model handle in a module-level global (`_llm_model_cache`).

The `ImageProcessor` wraps a 4-stage pipeline (OCR → enhanced description → lambda-index context → final description) with a 1-hour in-process TTL cache.

5.1. Which Gemini client do we pick? `@google/generative-ai` is the obvious choice, but it has a different streaming API, different error shapes, and does **not** raise the same `Resource has been exhausted` string that invariant #3's check grep'd for. Who writes the TS analog of the 60 s backoff + retry, and where does the check_invariants guard point?

5.2. Where does the Gemini API key live at runtime?
- `process.env.GOOGLE_API_KEY` in route handlers (server only) — same shape as Streamlit.
- GCP workload identity / service account — tidier, needs Vercel/Cloud Run integration.
- Secret Manager — requires infra we don't yet have (see SOP_SECRETS §6).

5.3. Do we preserve the per-request `ImageProcessor.context_cache` (TTL 3600 s) across requests in Next.js? Options:
- In-memory per-node: cheap, wrong under multi-node deploys.
- Redis/Upstash: correct, new infra.
- Drop the cache for v1: accept extra Gemini cost.

5.4. The current `describe_image` runs Stage 1 OCR (currently disabled — `utils.py :: ExtractText_OCR` returns a placeholder). Is OCR in scope for v2 day 1? If yes, native Node (`tesseract.js`) vs. a Python sidecar. If no, document the regression explicitly.

5.5. Lambda Index (`get_context_with_lambda_index` in `ImageProcessor`) — does the logic port straight as a TS function, or is it a candidate to replace with a vector-DB-native re-ranker? List the minimum behavioral invariants before rewriting.

5.6. The `get_random_slide_context`, `get_random_slide_with_image`, `get_context_from_slide_number` reads all return a weird shape where `documents` can be a list of characters and must be `.join()`ed — a bug waiting to bite. Do we preserve the shape for wire compatibility or fix it during the port (and risk contract drift with the Python FastAPI wrapper if both stay alive)?

5.7. Prompt fidelity: the Gemini prompts in `ppt_notes.py :: ExtractText_LLM`, `rag_core.py :: prompt_gemini*`, and `image.py` embed specific phrasing ("under 125 characters for alt text", "Image Information:", etc.). Do we copy them verbatim, version them (`prompts/v1.ts`), or re-write? Accessibility output quality is sensitive to prompt drift; lock the answer before the port.

5.8. Cost accounting: is there a per-user or per-deck cap to enforce in the TS version that Streamlit never had (rate limit, quota, API-key-per-tenant)?

---

## 6. State, Session, and Persistence

Streamlit gives us `st.session_state` for free — a per-user server-side dict keyed by the WebSocket. Next.js gives us nothing equivalent. Every state bullet below needs a concrete home.

| State today in `st.session_state` | What it holds | Proposed Next.js home |
|---|---|---|
| `processing_stage` | stage enum | URL segment / query? DB? cookie? |
| `presentation_model` | full Pydantic `Presentation` incl. `bytes` images | ? |
| `rag_core` | in-process singleton | not serializable, must be rebuilt per request |
| `image_processor` | in-process, holds `context_cache` | see 5.3 |
| `collection_id` | Chroma collection name | DB row or cookie |
| `current_batch`, `batch_size` | pagination | URL param |
| `uploaded_file_name`, `file_bytes` | 50+ MB binary | object store (GCS? local disk? tmpfs?) |
| `output_path` | server filesystem path to accessible.pptx | object store URL |

6.1. Do we introduce a **database** (Postgres? SQLite on the VM? Firestore?) and a `Job` / `Presentation` / `ImageDescription` schema, or do we lean on the existing Chroma collection + an object store and avoid adding a DB? Rank the two paths by net-new-infra count, failure modes, and migration toil.

6.2. Where do **uploaded `.pptx` blobs** live between stages? Options:
- Node-local `/tmp` keyed by `jobId` — breaks under multi-node, survives single-VM.
- GCS bucket — requires bucket + credentials.
- Base64-in-DB — bad, but noting it because it's what Streamlit's `st.session_state.file_bytes` basically does.

6.3. Where does the **rebuilt accessible deck** live for download? Same options as 6.2, but adds the question: does the download URL expire, require auth, or include a CSRF token? Streamlit's `st.download_button` bypasses all of this by serving bytes directly; we have to replace that pattern.

6.4. Do we version the job schema from day 1 (`jobs.schema_version`) so we can evolve it without breaking in-flight jobs during future deploys?

6.5. Does the UI support **resuming an in-progress job** from a new browser or tab? If yes, auth is required (section 10). If no, the job ID can be unguessable-random and URL-shared.

6.6. What's the retention policy for jobs (minutes? days?) and by extension for uploaded decks and Chroma collections? This matters for both Gemini cost and IRB data handling on the Aggrement branch (see §9).

---

## 7. Long-Running Work & File Uploads

Streamlit's `st.spinner(...)` happily blocks for minutes. Next.js Route Handlers (Node runtime, Vercel) cap at 60 s default (configurable to ~5 min on Vercel Pro, effectively unbounded on self-hosted). Answer these before picking a hosting target.

7.1. What is the worst-case end-to-end time for a full run on a 50-slide deck with 30 images, given current Gemini latency + 60 s backoff + Chroma add + rebuild? Measure on the live Streamlit and put a number here.

7.2. Given that number, do we:
- Run synchronously in a Route Handler and raise the Node timeout (self-host only, fragile).
- Split into **async job** — Route Handler accepts upload, enqueues, returns `jobId`; a worker processes; client polls `/api/jobs/[id]`. New infra: queue.
- Use **Server Actions + streaming** (`useFormState`/`useOptimistic`) with progressive RSC rendering.

7.3. If async: which queue? Options:
- BullMQ + Redis (adds Redis).
- Postgres-backed queue (adds Postgres; might already be on from 6.1).
- Google Cloud Tasks / Pub/Sub (adds GCP infra).
- Inline `setImmediate` + in-memory map (dev-only, do not ship).

7.4. File upload ceiling: Next.js `formData.get('file')` loads the whole file into memory. For 50 MB decks this is fine, for 500 MB it is not. Do we switch to chunked / resumable upload (tus, UploadThing, GCS signed URLs) up-front, or gate deck size at 50 MB and defer?

7.5. How do we surface progress? Option matrix:
- Polling `/api/jobs/[id]` every 2 s — simplest, works everywhere.
- Server-Sent Events stream — better UX, slightly more infra.
- Websocket — overkill for one direction.
- React Server Components streaming with `Suspense` — ergonomic but ties us to RSC, may not fit a queued worker model.

7.6. Cancel / abort semantics — Streamlit has none; what do we promise?

---

## 8. Invariant Translation Matrix

Each row in `docs/guardrails/INVARIANTS.md` is a constraint the Next.js port must honor. Answer one question per row.

| # | Invariant | Next.js-era question |
|---|---|---|
| 1 | `order_number` is the join key | Which TS type owns it? Where is it asserted (zod? branded type)? How does `scripts/check_invariants.py --only order_number` become a TS lint / test? |
| 2 | Alt text writes to `cNvPr/@descr` | Does our OOXML writer do this faithfully? What automated test proves it on every PR (see 15.x)? Is the existing `check_alt_text_xml` regex check still valid, or do we rewrite it to scan `src/lib/pptx/**` instead of `app/`? |
| 3 | Gemini 60 s backoff on `ResourceExhausted` | What is the TS equivalent error detection? What's the unit test that injects the 429 and asserts the delay? |
| 4 | Chroma access goes through the FastAPI wrapper | Does the wrapper stay (see §12)? If not, what replaces invariant #4 — a single `lib/chroma.ts` module that everything imports, forbidding direct `chromadb` client usage (which has no Node equivalent anyway)? |
| 5 | Image normalization before hashing or Gemini | Which Node image lib? Which tests cover WMF, EMF, PNG-with-palette, RGBA? |
| 6 | `chroma/` vector data is not to be cleaned | Does v2 relocate the data directory out of the repo (long-planned; §10.5 in INVARIANTS)? If yes, this is the time. |
| 7 | Group-shape recursion | TS analog of the `MSO_SHAPE_TYPE.GROUP` walk. See 4.6. |
| 8 | Consent gate (IRB) | See §9. |
| 9 | Branches are a pipeline, not a landscape | Does the Next.js code live on `nextjs-impl` (existing), a new `refactor/nextjs-migration` (this branch), or does it eventually become `main`? Either way, update `BRANCHING.md`. |
| 10 | Tech-debt gaps (secrets, firewall, no CI, drift) | Which gaps does v2 close by construction (e.g., ditching systemd closes 10.8), and which do we carry forward? |
| 11 | Configuration of record | Nginx, systemd, and shell scripts must land in repo. Does v2 replace them with a `Dockerfile` + compose / Cloud Run YAML that we commit? |

---

## 9. Consent Gate (IRB)

Invariant #8 is non-negotiable on the `Aggrement` branch. The Streamlit gate is a blocking full-page form that writes to `consent_responses.csv` on the VM before any upload UI renders.

9.1. Does v2 ship with the IRB consent gate from day 1? If yes, where does it live:
- Middleware at `src/middleware.ts` redirecting to `/consent` until a `consent_ok=...` cookie is set.
- A server component wrapper at `src/app/(consented)/layout.tsx`.
- A client-side modal (bad — bypassable).

9.2. Where is the consent record persisted?
- Flat CSV (matches current) on the host — brittle in multi-node.
- DB table `consent_events` (timestamp, user-pseudonym, deck hash).
- Append-only file in a GCS bucket.

9.3. What's the data contract with the IRB? We need the exact list of fields to store and the data-retention horizon in writing before we swap storage mechanisms.

9.4. Do we version the consent wording? If consent language changes, do previously-consented users re-consent?

9.5. Consent flow under auth (section 10) — is consent per-account (one-time) or per-session (every visit)?

---

## 10. Auth & Identity

Streamlit has no auth — anyone with the URL can use it. The GCP firewall also exposes ports 8001 and 8501 publicly (invariant #10.4). A Next.js migration is the natural point to add auth.

10.1. Do we gate the app with auth in v2? If yes:
- SUNY SSO (Brockport LDAP / CAS / Microsoft Entra)?
- Google Workspace login via NextAuth / Auth.js?
- Passwordless email link?
- None (matches current)?

10.2. If no auth, what's the abuse story for the publicly reachable Gemini-cost-bearing endpoints? (Today, nothing; this has not yet bitten because of obscurity.)

10.3. If we add auth, how does it interact with IRB consent (§9)? Does a logged-in user skip consent once they've accepted once, or does every session re-prompt?

10.4. Session storage: Auth.js JWT, database session, or iron-session cookies? Impact on 6.x state decisions.

10.5. Does the download URL for the accessible deck require a signed session, or is a short-lived unguessable URL sufficient?

---

## 11. Production Deployment

The live environment is bare-metal Python on systemd, behind nginx + certbot, on a single `e2-medium` Debian VM. The documented-but-unused Docker stack is in `Prod-v1`. Neither of these is a Next.js host. Answer before picking a hosting plan.

11.1. What is the **target runtime** for the Next.js app?
- **Vercel** — easy, great DX, pricing unclear at our scale, loss of control over timeouts and file-system persistence, hard to self-host Chroma alongside.
- **Self-host Node on the existing VM** under a new systemd unit — minimal new infra, keeps Chroma co-located, but fights Next.js's strengths and inherits all the drift in §11 of INVARIANTS.
- **Docker on the existing VM** — finally uses the Docker work sunk into `Prod-v1`, pairs naturally with a compose file that also runs the Chroma service.
- **Google Cloud Run** — managed, scales, pairs with Secret Manager and GCS, requires a move off `chroma/` on local disk.
- **GKE** — overkill for the current traffic.

11.2. If Vercel: where does Chroma live? Vercel has no persistent file system; Chroma either moves to a managed DB (Pinecone, Weaviate Cloud, Chroma Cloud, pgvector on Supabase) or to a separate container that Vercel calls. Which?

11.3. If self-host / Docker-on-VM: do we also finally add a second VM + load balancer, or accept the single-VM SPOF (invariant #10.7)?

11.4. TLS termination — nginx + certbot stays, or we move to Caddy (as `Prod-v1` intended), or a managed LB terminates TLS?

11.5. Process model:
- Single Next.js server (`next start`) — simple.
- Next.js standalone output + reverse proxy — smaller container.
- Next.js on a platform (Vercel / Cloud Run) — no process model to manage.

11.6. Environment variables: today it's a plaintext `.env` on the VM. Target:
- `.env.production` file + systemd `EnvironmentFile=` (minor upgrade, same risk).
- GCP Secret Manager + workload identity (closes invariant §10.1).
- `.env` managed by the hosting platform (Vercel / Cloud Run).

11.7. Logging & observability: today, `journalctl -u` on the VM. Target stack? (Cloud Logging, Datadog, Grafana + Loki, none for v1.)

11.8. What replaces the existing `start_scripts/` shell scripts (currently on-VM-only, per invariant §10.3)? They must be committed to this repo before we cut over.

11.9. Domain: stays `access.brockportsigai.org`. Does `/accessibility` path stay, or does the App Router get a cleaner `/`? nginx reverse-proxy config is the gatekeeper regardless; that config must land in repo (invariant #11).

---

## 12. Coexistence with the Chroma Service

The Python FastAPI wrapper at `app/chroma-api/app.py` is invariant #4's reason for existing (decouples Streamlit from the `chromadb` Python package). The Next.js scaffold's `src/lib/chroma.ts` already talks to it via HTTP.

12.1. Do we **keep** the FastAPI wrapper as-is and have Next.js speak HTTP to it? Pros: invariant #4 survives; zero Python-side change. Cons: we drag the Python runtime forward forever.

12.2. Do we **port the wrapper to TypeScript** (a `src/app/api/chroma/**/route.ts` layer over the Chroma server directly or via `chromadb`'s JS client, if we trust it)? Pros: one language. Cons: we own one more thing to maintain; the JS Chroma client is younger than the Python one.

12.3. Do we **replace Chroma** entirely for v2 (pgvector on Postgres, or Supabase vector, or Pinecone)? This also answers invariant #6 (the `chroma/` data dir problem) by construction.

12.4. Migration: can an existing Chroma collection on prod be read by the new stack on day 1 of cutover, or do we require re-indexing from uploaded decks? If the former, coordinate with §14. If the latter, accept the one-time Gemini cost of re-embedding.

12.5. Embedding function parity: what model does ChromaDB use today (default `all-MiniLM-L6-v2` via SentenceTransformers?), and what does the TS port use? If they differ, old vectors and new queries live in different spaces — incoherent. This is the single biggest footgun in this section.

12.6. If the wrapper survives, do we extend invariant #4 to say "TS code must also go through the wrapper, never via direct `chromadb` JS client"?

---

## 13. Existing `nextjs-impl` Scaffold Audit

The `nextjs-impl` branch already has a Next.js 16 + React 19 + Tailwind 4 + App Router scaffold with one route handler, one page, three lib modules, and the Python app moved under `backups/python-legacy/`. Before we write more on this branch, inventory what we keep.

13.1. `package.json` — keep Next 16.2.3 / React 19.2.4, or pin to the latest LTS minor at migration start? What's our policy on `caret` vs. exact versions (the scaffold uses caret)?

13.2. `src/app/layout.tsx` and global CSS — keep the current "dark + purple gradient" visual direction or redesign against a Brockport / accessibility-brand-compliant palette? (The tool is for accessibility; low contrast on a dark gradient is self-defeating.)

13.3. `src/app/page.tsx` — single client component with hard-coded mock review cards and a `setTimeout(() => setStage('review'), 2000)` fake transition. **Keep as reference only** and rebuild per §2? **Keep and iterate**? **Delete**?

13.4. `src/app/api/process/route.ts` — returns a hand-assembled `presentation` object derived from `PptxProcessor`. It does **no** Gemini, **no** Chroma, **no** alt-text write, **no** rebuild. Is this the spine we extend or a stub we throw away?

13.5. `src/lib/pptx-utils.ts` — the critical file per §4. Line 146's `updateAltText` is a `console.log` stub. Do we extend this file or start over with a deliberate API (`parse`, `rebuild`, `writeAltText`)?

13.6. `src/lib/gemini.ts` — has **no retry**, **no backoff**, **no image-normalization**. Would ship a direct regression of invariants #3 and #5. When do we fix it (before using it, obviously — but is that part of §5 or a prereq)?

13.7. `src/lib/chroma.ts` — `axios.post(...)` wrapper around the existing FastAPI service. Stays or evolves per §12 answer.

13.8. `backups/python-legacy/` — is that the **archive of record**, or do we keep the Python code live at `app/` during the transition and only move it to `backups/` at cutover?

13.9. `AGENTS.md` on that branch says "This is NOT the Next.js you know … Read the relevant guide in `node_modules/next/dist/docs/` before writing any code." Do we formalize that as a cross-branch rule on `refactor/nextjs-migration` too?

13.10. `CLAUDE.md` on that branch — inspect and decide whether it becomes part of our `AGENTS.md` lineage or gets deleted.

13.11. Does this branch start from `nextjs-impl` (cherry-pick or merge) or from `main` / `feature/agent-ops-framework` and re-bootstrap Next.js? Answer determines whether we inherit the scaffold's early mistakes.

---

## 14. Cutover, Rollback, Data Migration

14.1. Cutover mechanics: DNS flip of `access.brockportsigai.org`, nginx location swap (`/accessibility` → Next.js), dual-run with a kill switch, or a hard cut at a specific time?

14.2. What's the acceptance criterion for "v2 is ready"? List every user-visible feature from §4 of `docs/PROJECT_OVERVIEW.md` and mark green/red before flip.

14.3. Data to migrate (or not):
- Chroma collections on the VM — probably discardable (§12.4).
- `consent_responses.csv` — must migrate or be archived per IRB (§9.2).
- Uploaded decks in `/tmp` — discardable.

14.4. Rollback procedure — during v2's first 72 hours, what brings us back to Streamlit? Existing `docs/ops/SOP_ROLLBACK.md` assumes a git-based Python rollback on the VM. Rewrite for the Next.js host choice from §11.

14.5. Communication plan: who gets notified of the cutover, and on what channel?

14.6. Do we keep Streamlit **frozen and running** on port 8501 behind a feature-flagged nginx block for 30 days post-cutover, so we can re-expose it if v2 fails in production ways we didn't catch in staging?

---

## 15. Testing, Accessibility QA, Observability

15.1. What are the test layers?
- **Unit** (Vitest / Jest): zod schema, order_number invariant, alt-text XML writer round-trip, Gemini backoff, image normalization branches.
- **Integration** (Playwright + MSW or real Chroma-dev): upload → describe → download.
- **Fixture-based PPTX round-trip**: one real deck per invariant. Which decks do we commit as fixtures?
- **Accessibility QA**: do we run PowerPoint Accessibility Checker on the output deck in CI? There is no headless CLI — what's the plan?

15.2. Do we port `scripts/check_invariants.py` to TypeScript, or keep the Python script running in CI against the TS codebase? The regex-based checks mostly still apply but need path updates.

15.3. Do we port `scripts/doctor.py`, `scripts/preflight.py`, `scripts/smoke_test.py` to equivalent Node scripts, or do we keep Python as the "ops language" and call `python` from npm scripts?

15.4. What's the CI matrix? Today GitHub Actions runs `preflight.yml` on push/PR. For a Next.js app we also need:
- `npm run lint`
- `npm run build` (catches RSC / Client boundary errors)
- `tsc --noEmit`
- `playwright test`

15.5. Observability — what logs, traces, and metrics do we emit from day 1? At minimum: Gemini call latency, Gemini retries, Chroma add/query latency, request duration per route, job success rate, job durations.

15.6. Error reporting — Sentry, Google Error Reporting, or `console.error` + Cloud Logging?

15.7. Load testing — what's the target concurrency and who runs the test before cutover?

---

## 16. Framework Integration

This migration will outlive several agents. It must fit the ops framework already built.

16.1. Does this work ship as a single giant PR, or as a sequence of PRs each using `docs/templates/REFACTOR.md` (and some `FEATURE.md`) against this branch?

16.2. Which existing framework files need rewrites after the migration lands?
- `docs/PROJECT_OVERVIEW.md` — the architecture diagram.
- `docs/AGENT_CONTEXT.md` — file layout & conventions.
- `docs/PRODUCTION_ENVIRONMENT.md` — new runtime & topology.
- `docs/guardrails/INVARIANTS.md` — paths in every §Where; see §8 above.
- `docs/guardrails/BRANCHING.md` — `nextjs-impl` reaches end-of-life.
- `docs/ops/SOP_DEPLOY.md` — bare-metal → new deploy.
- `docs/ops/SOP_ROLLBACK.md` — see §14.4.
- `docs/ops/SOP_SECRETS.md` — key rotation steps for the new host.
- `scripts/*` and `.github/workflows/preflight.yml` — see §15.

16.3. Do we add new framework documents that this migration requires?
- `docs/guardrails/NEXTJS_RULES.md` — RSC vs Client rules, Server Action payload caps, when to choose Route Handler vs Action, edge runtime forbidden.
- `docs/ops/SOP_QUEUE.md` — if we adopt BullMQ / Cloud Tasks in §7.
- `docs/refactor/NEXTJS_MIGRATION_DESIGN.md` — the successor to this document, post-review.

16.4. When do we archive this file? When section 14 ships, this document becomes historical and should be moved to `docs/refactor/history/` with its answers inline.

16.5. Who owns the refactor? A single reviewer for all PRs on this branch? `CODEOWNERS` update?

---

## 17. Open bets to make explicit

The following are judgment calls that, once made, become implicit everywhere downstream. Make them on purpose.

17.1. **Language monoculture vs. polyglot.** TypeScript-only or TypeScript + a Python sidecar (parsing / OCR / maybe Chroma wrapper). The polyglot path costs more ops but preserves python-pptx fidelity (see 4.1).

17.2. **Runtime monoculture vs. platform.** Self-host vs. Vercel vs. Cloud Run. Affects §7 timeouts, §11 secrets, §12 Chroma, §15 observability.

17.3. **State: cookies vs. database.** Adding Postgres is an infra escalation; not adding it pushes complexity into the URL and cookies.

17.4. **RSC-first vs. Client-first.** Existing scaffold is client-first; §2 recommends RSC-first. Commit to one.

17.5. **Auth now vs. never.** §10 — adding auth now closes §10.4 of INVARIANTS forever; not adding it preserves the current UX and obscurity-based security.

17.6. **Re-index vs. carry vectors forward.** §12.4 — deletes prod Chroma data and re-embeds vs. wire-compatible migration.

---

## 18. Halt

This document ends here. Do not start writing Next.js code until every section has an owner and an answer. The next document on this branch is `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`, which is produced by answering the questions above.
