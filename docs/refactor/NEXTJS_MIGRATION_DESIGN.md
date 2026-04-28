# Next.js Migration — Design

> **Status:** Finalized technical design, derived from explicit architectural decisions made against [`NEXTJS_MIGRATION_PREFLIGHT.md`](./NEXTJS_MIGRATION_PREFLIGHT.md). This document is the **binding blueprint** for the refactor on branch `refactor/nextjs-migration`.
>
> **Institutional / operational items that require external sign-off** are marked **[PENDING MANUAL CONFIRMATION]**. These do not block technical work; they block cutover.
>
> **Invariants of record:** [`docs/guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md). Every design choice below is justified against those invariants; none are relaxed.

---

## 0. Executive Summary

We are splitting the monolithic Python + Streamlit app into a two-tier system:

- **Tier 1 — Next.js 16 App Router on Vercel.** Owns all UI, routing, auth, consent enforcement, job state, DB access, and file upload surface. Stateless compute.
- **Tier 2 — Python FastAPI sidecar + ChromaDB on the existing GCP VM via Docker Compose.** Owns PPTX parsing, OOXML rebuild (including the `cNvPr/@descr` write), Gemini calls, and Chroma access. All legacy Python invariants live here unchanged.

Supabase provides **auth, Postgres, and blob storage**. Prisma is the ORM. Jobs run asynchronously: Next.js inserts a `Job` row, fires one HTTP call to the Python sidecar to start work, and the client polls a Next.js route handler that reads job status from Postgres. No background workers and no queue in Next.js — the Python service is the worker.

Caddy (from `Prod-v1`) continues to terminate TLS on the VM, but only in front of the FastAPI endpoint, not a UI. The Streamlit `web` service and all `systemd` units are retired at cutover.

### One-diagram view

```
          ┌─────────────────────────────┐
  user ──▶│   Vercel: Next.js (App Rtr) │◀──── Supabase Auth (JWT)
          │   RSC + Client + Route Hdrs │
          └───────┬──────────────┬──────┘
                  │              │
         Prisma   │              │  HTTPS (mTLS/shared-secret)
                  ▼              ▼
        ┌────────────────┐  ┌──────────────────────────┐
        │ Supabase:      │  │ GCP VM (Docker Compose)  │
        │  - Postgres    │  │  caddy  ──▶  chroma-api  │
        │  - Storage     │  │              (FastAPI,   │
        │  - Auth        │  │               parse,     │
        └────────────────┘  │               rebuild,   │
                            │               Gemini)    │
                            │              └──▶ chroma │
                            │                  (vector)│
                            └──────────────────────────┘
```

No component of this diagram violates any invariant; §16 of this document maps each invariant to its owner.

---

## 1. Scope & End-State (resolves preflight §1)

| # | Preflight question | Decision |
|---|---|---|
| 1.1 | Replacement or coexist? | **Replacement.** Vercel deployment will own the public domain after cutover. |
| 1.2 | Which branch is behavior truth? | **`Aggrement`** (the live one), including the IRB consent flow. |
| 1.3 | Feature freeze on Streamlit? | **[PENDING MANUAL CONFIRMATION]** — owner and exact freeze date. |
| 1.4 | Node-only or keep FastAPI? | **Keep FastAPI** on the VM; Next.js never touches PPTX or Chroma directly. |
| 1.5 | KPIs? | Minimum: upload-to-download wall time, Gemini cost/deck, p95 page TTI, 5xx rate. Baseline measurements against live Streamlit **[PENDING]**. |
| 1.6 | Streamlit sunset policy? | Freeze Streamlit post-cutover for 30 days (read-only on a non-public port), then uninstall the systemd units. See §14. |

**[PENDING MANUAL CONFIRMATION]:** hard feature-freeze date on `Aggrement`; DNS cutover date; whether Streamlit stays reachable on a private port during the 30-day window.

---

## 2. UI & Routing (resolves preflight §2)

### 2.1 Route map (App Router)

All routes under `src/app/` in the Next.js repo.

| Path | Kind | Stage | Notes |
|---|---|---|---|
| `/` | RSC | Landing | Marketing shell, static. |
| `/auth/sign-in`, `/auth/callback` | Client + RSC | — | Supabase Auth UI. |
| `/consent` | RSC + Server Action | Consent gate | Rendered when `profile.consent_accepted_at IS NULL`. Server Action writes the flag. |
| `/upload` | RSC shell + Client dropzone | Stage 1 | `POST /api/uploads` streams the file to Supabase Storage, creates a `Job` row, returns `jobId`. |
| `/process/[jobId]` | RSC | Stage 2 | Reads `Job` from Postgres. Streams status via React `<Suspense>` + client polling component. Redirects to `/review/[jobId]` when `status='awaiting_review'`. |
| `/review/[jobId]` | RSC shell + Client review grid | Stage 3 | Paginated per-image review with optimistic updates via Server Actions. |
| `/download/[jobId]` | RSC | Stage 4 | Generates a Supabase Storage signed URL for the rebuilt deck, renders a download button. |
| `/api/uploads` | Route Handler, Node runtime | — | Streams `multipart/form-data` → Supabase Storage, inserts `Job`, POSTs to Python. |
| `/api/jobs/[id]` | Route Handler, Node runtime | — | Polling endpoint. Reads Postgres only. |
| `/api/jobs/[id]/descriptions` | Route Handler, Node runtime | — | Bulk read of per-image descriptions for the review UI. |
| `/api/webhooks/processor` | Route Handler, Node runtime | — | Called by the Python sidecar to push phase transitions. Protected by a shared secret (§11.4). |

The App Router's segmented routes map 1:1 to the Streamlit `processing_stage` enum. Back-button and reload behavior fall out naturally because stage lives in the URL.

### 2.2 RSC vs Client split (the architectural bet)

| Surface | RSC boundary | Client boundary |
|---|---|---|
| Landing, auth shells, consent, upload, process, download | Entire page shell, all data fetches | A single `<Dropzone />` on `/upload`, a single `<StatusPoller jobId=... />` on `/process/[jobId]`, a single `<DownloadButton />` on `/download/[jobId]`. |
| `/review/[jobId]` | Page shell, initial batch of descriptions (server-fetched from Postgres) | `<ReviewGrid>`, `<ImageCard>`, `<AltTextEditor>` — interactive editors with optimistic UI. Server Actions write back to Postgres. |

**Rules that govern the split** (all enforced in code review):

- RSC is the default. `'use client'` is only added to a component that needs local state, refs, or browser event handlers.
- No Gemini key, no Supabase service-role key, and no FastAPI shared secret ever appears in a client component or in any RSC that returns props to a client component.
- Route Handlers only ever run on the **Node runtime**. No Edge runtime anywhere (Gemini SDK needs Node; `adm-zip`-style libraries are not in scope but the policy stays).
- Server Actions are used for write paths where the response body is trivial (alt-text edits, consent acceptance). Route Handlers are used for reads that the client polls and for the webhook from Python.

### 2.3 Progress UX

- Client polls `GET /api/jobs/[id]` every **2 s** with exponential back-off to 10 s after 2 min.
- Optional upgrade: replace polling with Supabase Postgres Changes (WebSocket-driven row subscriptions) once the base path is stable. **Post-v1.**

### 2.4 Existing `nextjs-impl` scaffold disposition

The current prototype at `nextjs-impl` is **reference only**. This design deliberately does **not** build on:
- its single-page client-only `page.tsx`,
- its stubbed `src/lib/pptx-utils.ts` (violates invariants #1, #2, #5, #7 by omission),
- its `src/lib/gemini.ts` (no 60 s backoff — violates invariant #3).

We do carry forward:
- `package.json`'s Next.js 16.2.3 / React 19.2.4 pinning.
- Tailwind 4 + the `lucide-react` / `framer-motion` UI direction, provided the final palette clears WCAG AA contrast (§15.5).
- The `backups/python-legacy/` archival pattern — once `refactor/nextjs-migration` lands, the Streamlit code is moved there.

---

## 3. Backend Surface (resolves preflight §3)

### 3.1 Python FastAPI sidecar — expanded responsibilities

The sidecar currently at `app/chroma-api/app.py` is renamed in intent to **"processing service"** and gains job-orchestration endpoints. The existing Chroma CRUD endpoints stay untouched (invariant #4). No Python logic moves to TypeScript.

New endpoints (to be implemented on this branch when the code-freeze unlocks):

| Verb + Path | Purpose | Auth | Body |
|---|---|---|---|
| `POST /jobs/:id/start` | Begin processing for a job. Called by Next.js after an upload finishes. | Shared secret header | `{ storage_object: string, presentation_name: string }` |
| `POST /jobs/:id/commit` | Client finished review. Rebuild the deck and upload to Supabase Storage. | Shared secret header | `{ descriptions: [{ order_number, alt_text }] }` |
| `GET  /jobs/:id/status` | Mirror of Postgres status. Used only by health checks and debug. | Shared secret header | — |
| `POST /jobs/:id/cancel` | Abort and cleanup. | Shared secret header | — |

All new endpoints emit a webhook POST to `${NEXT_PUBLIC_APP_URL}/api/webhooks/processor` on each phase transition (`parsed`, `described`, `awaiting_review`, `rebuilding`, `ready`, `error`) so Postgres stays authoritative for status.

The existing endpoints (`/collections`, `/collections/*/add`, `/collections/*/query`, `/health`) are preserved verbatim. Invariant #4 holds: Next.js talks to Chroma only through this service.

### 3.2 Python module boundaries — preserved

All existing modules stay in place and keep their invariants:

| Module | Preserved responsibility | Invariant owned |
|---|---|---|
| `app/pptx_rag_quizzer/utils.py :: parse_powerpoint` | Extract text/images; set `order_number` | #1, #7 |
| `app/ppt_notes.py :: process_powerpoint_with_rag_enhanced` (or its `Aggrement`-branch analog in `app/pptx_rag_quizzer/pptx.py`) | OOXML rebuild; `cNvPr/@descr` write | #2 |
| `app/pptx_rag_quizzer/rag_core.py :: RAGCore` | Gemini calls with 60 s backoff; Chroma add/query via HTTP wrapper | #3, #4 |
| `app/pptx_rag_quizzer/image.py :: ImageProcessor` | 4-stage describe pipeline; image normalization | #5 |
| `app/chroma-api/app.py` | FastAPI; wraps Chroma | #4 |

A small module is added for the new surface:

- `app/processing_service/jobs.py` — implements the `/jobs/*` endpoints. Its only job is orchestration: pull the blob from Supabase Storage via a signed URL, call the existing modules, push results back. It writes **nothing** OOXML-adjacent itself.

### 3.3 Error contract (wire)

Both the Next.js route handlers and the Python endpoints return a uniform shape:

```json
{ "ok": true, "data": { ... } }
{ "ok": false, "error": { "code": "GEMINI_QUOTA", "message": "...", "retryable": true } }
```

Error codes are enumerated (`UPLOAD_TOO_LARGE`, `CONSENT_REQUIRED`, `GEMINI_QUOTA`, `GEMINI_FAILED`, `PARSE_FAILED`, `REBUILD_FAILED`, `CHROMA_UNAVAILABLE`, `JOB_NOT_FOUND`, `UNAUTHORIZED`). The Python side maps existing `Resource has been exhausted` strings to `GEMINI_QUOTA` while still honoring the 60 s backoff internally (invariant #3).

### 3.4 Runtime constraints — written down

- All Next.js Route Handlers: **Node runtime only**. Never Edge.
- All Server Actions that touch Supabase service-role, the FastAPI shared secret, or signed-URL generation: **Node runtime only**.
- No TypeScript code imports anything that talks to Chroma or parses `.pptx`. Enforced by a new invariant check (§16.4).

---

## 4. PPTX Parsing & Rebuild (resolves preflight §4)

**Decision:** PPTX parsing, OOXML tree traversal, and rebuild stay in Python. TypeScript never opens a `.pptx`.

This collapses all of preflight §4 into a short statement:

- `parse_powerpoint` unchanged.
- The `cNvPr/@descr` writer (invariant #2) unchanged.
- `convert_image_to_png_or_jpg` and WMF/EMF handling (invariant #5) unchanged.
- Group-shape recursion (invariant #7) unchanged.
- `python-pptx` stays the parser of record.

The Next.js side's only responsibility for a deck is to move its bytes to Supabase Storage and hand the Python side a reference.

> Consequence: `nextjs-impl/src/lib/pptx-utils.ts` is deleted on this branch and not replaced. Any future PR that introduces TS-side OOXML code must pass through a guardrail PR that relaxes this decision explicitly.

---

## 5. RAG Pipeline (resolves preflight §5)

**Decision:** Gemini and Chroma stay in Python. Invariants #3, #4, #5 are preserved without translation. No JavaScript Gemini or Chroma client is introduced.

| Preflight concern | Resolution |
|---|---|
| Gemini client choice | Stays `google.generativeai` (Python). |
| 60 s backoff | Stays in `rag_core.py`. |
| `GOOGLE_API_KEY` location | Docker Compose env var on the VM (injected from `.env` on the host; upgrade path to GCP Secret Manager is §11.6). **Never** present in the Vercel environment. |
| `ImageProcessor.context_cache` (TTL 3600s) | Stays per-process in the Python container. Same behavior as today. Post-v1 upgrade to Redis is possible; not required. |
| OCR (`ExtractText_OCR`) | Already a placeholder on `main`. Stays a placeholder. Explicit regression accepted. |
| Lambda Index | Unchanged. |
| Prompt wording | Frozen exactly as on `Aggrement` branch at cutover. **[PENDING MANUAL CONFIRMATION]** for the owner of prompt changes going forward. |
| Per-user Gemini cost cap | Enforced in Next.js Route Handler before `/jobs/:id/start` is called: per-user rate-limit row in Postgres. Concrete numbers **[PENDING]**. |

---

## 6. State, Auth, Database (resolves preflight §6 and §10)

### 6.1 Supabase project layout

Supabase gives us three primitives used here:

- **Auth** — email + magic link (optionally SUNY SSO later). Issues a JWT that both Next.js and (optionally) the Python sidecar can verify. For v1 only Next.js verifies; Python trusts the shared-secret header.
- **Postgres** — app database. Accessed from Next.js via Prisma only. The Python sidecar does **not** read or write Postgres directly; it talks to Next.js webhooks instead. This keeps schema ownership in one place.
- **Storage** — two buckets:
  - `pptx-uploads` (private) — raw uploads. Lifecycle rule deletes after N days (see §6.4).
  - `pptx-outputs` (private) — rebuilt accessible decks. Same lifecycle.

### 6.2 Prisma schema (authoritative)

Filename `prisma/schema.prisma`. Generated types live under `src/lib/db`. No migration is applied by this document; the schema here is the agreement.

```prisma
generator client {
  provider = "prisma-client-js"
}

datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")        // Supabase pooled connection
  directUrl = env("DIRECT_URL")         // Supabase direct (migrations)
}

/// One row per authenticated user. Mirrors Supabase auth.users.id.
model Profile {
  id                     String   @id            // == auth.users.id (uuid)
  email                  String   @unique
  consentAcceptedAt      DateTime?               // invariant #8 anchor
  consentVersion         String?                 // e.g. "v1-2026-04"
  createdAt              DateTime @default(now())
  jobs                   Job[]
  consentEvents          ConsentEvent[]
}

/// One row per deck upload.
model Job {
  id                     String      @id @default(uuid())
  profileId              String
  profile                Profile     @relation(fields: [profileId], references: [id])

  uploadedFilename       String
  uploadObjectPath       String      // key in 'pptx-uploads' bucket
  outputObjectPath       String?     // key in 'pptx-outputs' bucket, null until ready

  status                 JobStatus   @default(queued)
  phase                  String?     // human-readable: "parsing" / "describing" / "rebuilding"
  progressCurrent        Int         @default(0)
  progressTotal          Int         @default(0)
  errorCode              String?
  errorMessage           String?

  collectionId           String?     // Chroma collection id owned by the Python side
  schemaVersion          Int         @default(1)

  createdAt              DateTime    @default(now())
  startedAt              DateTime?
  awaitingReviewAt       DateTime?
  committedAt            DateTime?
  readyAt                DateTime?

  descriptions           SlideDescription[]

  @@index([profileId, createdAt])
}

enum JobStatus {
  queued
  parsing
  describing
  awaiting_review
  rebuilding
  ready
  error
  cancelled
}

/// One row per reviewable image in a deck. `orderNumber` is the invariant-#1 join key.
model SlideDescription {
  id                     String   @id @default(uuid())
  jobId                  String
  job                    Job      @relation(fields: [jobId], references: [id], onDelete: Cascade)

  slideNumber            Int
  orderNumber            Int      // invariant #1 — must match python-pptx shape order
  itemType               String   // "image" | "text" (text is read-only, images are editable)
  aiDescription          String?  // Gemini-generated
  finalAltText           String?  // user-edited; this is what gets written to cNvPr/@descr

  createdAt              DateTime @default(now())
  updatedAt              DateTime @updatedAt

  @@unique([jobId, slideNumber, orderNumber])
  @@index([jobId])
}

/// Append-only IRB audit trail. Duplicates some Profile data on purpose.
model ConsentEvent {
  id                     String   @id @default(uuid())
  profileId              String
  profile                Profile  @relation(fields: [profileId], references: [id])
  acceptedAt             DateTime @default(now())
  consentVersion         String
  ipHash                 String                     // SHA-256(ip + pepper)
  userAgent              String?
}
```

### 6.3 Session & state mapping

Every Streamlit `st.session_state` key from preflight §6 has a new home:

| `st.session_state` key | New home |
|---|---|
| `processing_stage` | URL segment (`/upload`, `/process/[id]`, `/review/[id]`, `/download/[id]`). |
| `presentation_model` | Not held in Next.js at all. Lives as rows in `SlideDescription` + the object in `pptx-uploads`. |
| `rag_core`, `image_processor` | Per-request in the Python container. Unchanged. |
| `collection_id` | `Job.collectionId`. |
| `current_batch`, `batch_size` | Query-string params on `/review/[id]?page=N`. |
| `uploaded_file_name`, `file_bytes` | `Job.uploadedFilename`, `Job.uploadObjectPath`. Never in Node memory past the stream. |
| `output_path` | `Job.outputObjectPath` (→ signed URL at download time). |

### 6.4 Retention — **[PENDING MANUAL CONFIRMATION]**

Defaults proposed for approval:

- Uploads bucket: **30 days**, then hard delete.
- Outputs bucket: **30 days**, then hard delete.
- Chroma collections on the VM: **30 days**, swept by a cron inside the chroma-api container.
- `Job` and `SlideDescription` rows: retained indefinitely (no PII); linked blobs are gone after 30 days.
- `ConsentEvent` rows: retained per IRB policy — **[PENDING]** exact horizon.

---

## 7. Long-Running Work (resolves preflight §7)

### 7.1 Async job model

```
Client ──POST /api/uploads (multipart)──▶ Next.js Route Handler
                                            │
                                            ├─ stream body → Supabase Storage (resumable)
                                            ├─ INSERT Job(status=queued)
                                            └─ POST  $PY_SERVICE_URL/jobs/{id}/start
                                                │
                                                └─ 202 Accepted, returns { jobId }
Client ──redirect /process/[jobId]─▶

Python sidecar (background asyncio task per request):
    parse → describe (Gemini) → chroma.add → webhook status=awaiting_review

Client polls GET /api/jobs/[id] every 2–10s → redirect /review/[jobId] when ready.

User reviews, clicks "Confirm & Export":
Client ──POST /api/jobs/[id]/commit──▶ Next.js Route Handler
                                        │
                                        ├─ UPDATE SlideDescription.finalAltText
                                        └─ POST $PY_SERVICE_URL/jobs/{id}/commit
                                                │
                                                └─ rebuild deck → upload to Supabase
                                                └─ webhook status=ready

Client polls until ready → redirect /download/[jobId] → signed URL.
```

Key consequences:

- **No Vercel timeout risk.** Route Handlers never wait for Gemini or rebuild. The slowest Route Handler is `/api/uploads`, bounded by upload bandwidth.
- **No Next.js-side worker / queue.** The Python service is the worker. Concurrency is bounded by the container's thread/process pool, tuned in Compose.
- **Single source of truth for status.** Postgres. The Python service pushes transitions via webhook; Next.js clients never read Python directly.

### 7.2 Upload ceiling

Hard limit: **50 MB** per `.pptx` in v1 (matches existing Streamlit behavior). Enforced at two layers:
- Next.js Route Handler rejects `Content-Length > 50 MB` before streaming.
- Supabase Storage bucket policy sets the same limit.

Above-50MB support is explicitly out of scope for v1; resumable/tus-style uploads are a post-v1 upgrade.

### 7.3 Cancellation

`POST /api/jobs/[id]/cancel` sets `status='cancelled'` in Postgres and POSTs the Python sidecar. The Python task checks `status` at each phase boundary and exits cleanly. Mid-Gemini-call cancellation is best-effort (we don't kill requests in-flight; we drop the result).

---

## 8. Invariant Translation Matrix (resolves preflight §8)

| # | Invariant | Owner in new design | Enforcement |
|---|---|---|---|
| 1 | `order_number` is the join key | Python `parse_powerpoint` sets it. Postgres `SlideDescription.orderNumber` mirrors it. Unique constraint `(jobId, slideNumber, orderNumber)`. | Python-side unchanged (`scripts/check_invariants.py --only order_number`). TS side: Prisma schema guarantees it (DB constraint). |
| 2 | Alt text writes `cNvPr/@descr` | Python rebuild only. TypeScript never writes OOXML. | Unchanged Python invariant check. |
| 3 | Gemini 60 s backoff | Python `rag_core.py`. | Unchanged Python invariant check. |
| 4 | Chroma via FastAPI wrapper | Strengthened: **no** Chroma client of any language outside the FastAPI container. | Invariant check extended (§16.4) to scan Next.js TS for `chromadb` or a JS vector-DB client and fail. |
| 5 | Image normalization | Python `ImageProcessor`. Unchanged. | Unchanged Python invariant check. |
| 6 | `chroma/` data is not to be cleaned | Now a **named Docker volume** (`chroma_data`), mounted at `/chroma/chroma` inside the Chroma container. Lives outside the repo tree. Closes the `git clean` footgun. | `.gitignore` scan unchanged; new compose-level invariant added to `INVARIANTS.md` (§16.1). |
| 7 | Group-shape recursion | Python. Unchanged. | Unchanged Python invariant check. |
| 8 | Consent gate (IRB) | Moved from Streamlit first-page + CSV to `Profile.consentAcceptedAt` + `ConsentEvent` + Next.js middleware (§9). Stronger than before: cannot bypass by direct URL. | New middleware-level assertion; new test in §15.3. |
| 9 | Branches are a pipeline | `refactor/nextjs-migration` → PR → merge into `Aggrement`'s successor (`main` after cutover). No lateral merges with `Prod-v1`. Docker Compose assets cherry-picked, not merged. | Branching doc updated (§16.3). |
| 10 | Tech-debt gaps (§10.x) | Closed by construction: 10.1 (plaintext secrets → Vercel env + Compose env from a secured `.env`; further upgrade to Secret Manager is noted in §11.6), 10.4 (public ports closed; Caddy fronts only the FastAPI), 10.6 (Vercel deploys + GitHub Actions = implicit CI), 10.7 (Vercel scales the UI; backend SPOF remains — noted), 10.8 (compose + Caddyfile + Dockerfiles all in repo). | Update to `INVARIANTS.md §10` at cutover. |
| 11 | Configuration of record | **Satisfied on day 1.** Docker Compose, Dockerfile.api, Caddyfile, Prisma schema, middleware, and every env var template (`.env.example`) are committed on this branch. Changes on the VM must land here first. | New invariant-level test: `docker compose config` must parse, `Caddyfile fmt` clean. |

---

## 9. Consent Gate (resolves preflight §9)

### 9.1 Flow

1. A user signs in via Supabase Auth.
2. `middleware.ts` inspects every non-auth, non-asset request. If `profile.consentAcceptedAt IS NULL`, it rewrites to `/consent`.
3. `/consent` is an RSC that fetches the current consent text from the repo (`docs/irb/consent-v1.md`, authoritative, version-pinned in code), plus the user's `consentVersion` if present.
4. On submit, a Server Action:
   a. inserts a `ConsentEvent` row with IP hash + UA + version,
   b. updates `Profile.consentAcceptedAt` and `consentVersion`,
   c. redirects to `/upload`.
5. The consent gate **cannot** be bypassed by deep-linking to `/upload`, `/process/*`, `/review/*`, `/download/*`, `/api/uploads`, or `/api/jobs/*` — middleware runs first.

### 9.2 IRB specifics — **[PENDING MANUAL CONFIRMATION]**

- Exact wording of `docs/irb/consent-v1.md`.
- Exact list of fields to store on `ConsentEvent` (IP policy, UA, pseudonym, research cohort).
- Retention horizon.
- Whether consent-version bumps require re-consent (default: yes).
- Whether minors-data policy applies.

The technical substrate (Profile flag + append-only event log + middleware) is frozen regardless of the above.

### 9.3 Data migration

The legacy `consent_responses.csv` on the prod VM is imported as historical `ConsentEvent` rows, or archived off-VM if IRB prefers a clean break. **[PENDING]** — the choice is a policy call, not an architectural one.

---

## 10. Auth & Identity (resolves preflight §10)

- **Provider:** Supabase Auth.
- **v1 sign-in:** email + magic link.
- **v1.1+ sign-in:** SUNY SSO via Supabase's external OIDC provider. Target post-cutover.
- **Session:** Supabase cookie session consumed server-side by `@supabase/ssr`. No JWT in localStorage.
- **Server access:** Route Handlers and Server Actions get the user via `supabaseServerClient(cookies())`. Unauthed → 401.
- **Middleware:** enforces auth + consent on all non-auth routes.
- **Service-role key:** stored as `SUPABASE_SERVICE_ROLE_KEY` in Vercel env only. Never in a client bundle. Used only from Route Handlers that need to bypass RLS (admin/debug endpoints) — none in v1.

Public-internet exposure of Gemini-bearing endpoints (§10.4 of INVARIANTS) is now closed by construction: `/api/jobs/*` requires a session; `/jobs/*` on the Python side requires the shared secret; the raw Chroma port 8000 is **not** published outside the Docker network (§11.3).

---

## 11. Production Deployment (resolves preflight §11)

### 11.1 Split hosting

| Tier | Host | Runtime | Code location |
|---|---|---|---|
| Next.js UI | **Vercel** | Node 22 (Vercel's current default) | `nextjs/` in this repo (see §13.1). |
| Python FastAPI + Chroma | **GCP VM** (`instance-20250905-023343-pub`), Docker Compose | Python 3.11-slim + Chroma image | `docker-compose.yml`, `Dockerfile.api`, `Caddyfile` — cherry-picked from `Prod-v1`. |

### 11.2 Compose stack (post-strip)

The `Prod-v1` compose file has **four** services: `chroma`, `chroma-api`, `web` (Streamlit), `caddy`. We keep three and drop `web`:

```
services:
  chroma:
    # unchanged from Prod-v1: ghcr.io/chroma-core/chroma:latest
    # volume mount changes from ./chroma-db (bind) to a named volume `chroma_data`
    # closes invariant #6's git-clean footgun.
  chroma-api:
    # build: Dockerfile.api, unchanged
    # ENV adds: PROCESSOR_SHARED_SECRET, WEBHOOK_URL (=Vercel app URL)
  caddy:
    # Caddyfile reverse_proxy changes from web:8501 to chroma-api:8001
    # new host: api.brockportsigai.org (or a subdomain TBD). Public port 443 only.
```

**Deleted:** the `web` service, both `ports: "8501:8501"` and `ports: "8000:8000"` mappings to the public side (8000 stays container-internal; 8001 is reached only through Caddy/443).

### 11.3 Network posture

- Public ingress on the VM: TCP/443 only (Caddy). TCP/80 redirects to 443.
- GCP firewall rules `allow-8501` and `allow-8001` are **deleted** at cutover (invariant §10.4 closes).
- Streamlit systemd units (`streamlit-app`, `chroma-api`, `chroma-db`) are **stopped and disabled** at cutover. Rollback re-enables them if needed (§14.4).
- `chroma` service is reachable only inside the `sap_net` Docker network.

### 11.4 Secrets

| Secret | Lives in | Injected via |
|---|---|---|
| `DATABASE_URL`, `DIRECT_URL` | Vercel env | Vercel dashboard (encrypted) |
| `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY` | Vercel env (public-safe) | Vercel dashboard |
| `SUPABASE_SERVICE_ROLE_KEY` | Vercel env (server-only) | Vercel dashboard |
| `PROCESSOR_SHARED_SECRET` | Vercel env AND VM `.env` | both sides check a shared HMAC header |
| `PROCESSOR_BASE_URL` (e.g. `https://api.brockportsigai.org`) | Vercel env | Vercel dashboard |
| `GOOGLE_API_KEY` | VM `.env` only | Docker Compose `env_file` |
| `WEBHOOK_URL` (Vercel app URL), `WEBHOOK_SIGNING_SECRET` | VM `.env` only | Docker Compose `env_file` |
| `SUPABASE_STORAGE_DOWNLOAD_URL`s | generated per request | Next.js route handlers |

The VM `.env` file is `chmod 600` and owned by a dedicated user. This closes the plaintext-`.env`-in-repo-root risk without adding Secret Manager. Upgrade path to GCP Secret Manager is documented as post-v1 and does not block cutover.

### 11.5 TLS

- `api.brockportsigai.org` (VM) — Caddy + Let's Encrypt ACME (from the existing Caddyfile).
- `access.brockportsigai.org` (Vercel) — Vercel-managed cert.
- DNS: `access` points at Vercel's edge; `api` points at the VM's external IP. **[PENDING]** DNS flip date.

### 11.6 Observability

- Next.js: Vercel Logs + `@vercel/otel` → (target) Google Cloud Logging. Pino for structured logs. Sentry for errors (post-v1 if needed).
- Python: `docker logs` + Google Cloud Ops Agent already installed on the VM.
- Both sides emit a `jobId` in every log line for correlation. Required by SOP.

### 11.7 Start / stop scripts retired

The existing `start_scripts/` directory on the VM (Streamlit + systemd) is archived into `backups/python-legacy/start_scripts/` on this branch and deleted from the VM post-cutover. Replacement is `docker compose up -d` + a one-line `systemd` unit that runs that command at boot (`docker-compose@sap.service`) — committed to `deploy/systemd/` on this branch.

---

## 12. Chroma Coexistence (resolves preflight §12)

- **Keep** the FastAPI wrapper exactly as-is. Invariant #4 holds.
- **Keep** the `chromadb` Python client as the only client of Chroma in the system.
- **Do not** introduce a JS Chroma client.
- **Do not** replace Chroma with pgvector / Pinecone / Supabase Vector in v1. Noted as a plausible post-v1 consolidation (the Supabase-already-there lever) but deliberately not in scope.
- **Embedding-function parity is a non-issue** because Chroma continues to use its built-in default embedder inside the same container.
- **Data migration:** existing collections on the prod VM are **discarded** at cutover. New uploads re-index from zero. Accepted cost because:
  - Collections are cheap to rebuild on upload.
  - Prod's in-repo `chroma/` directory is messy (invariant #10.5 debt) and migrating it while also moving to a named volume doubles risk.
  - The system has never relied on long-lived collections across decks.

---

## 13. Repository Layout & `nextjs-impl` Audit (resolves preflight §13)

### 13.1 Proposed repo layout on this branch

```
/
├── app/                       # Python — unchanged (source of invariants #1–#7)
│   ├── chroma-api/
│   ├── models/
│   ├── pptx_rag_quizzer/
│   ├── ppt_notes.py
│   └── processing_service/    # NEW — /jobs/* endpoints (same FastAPI app)
├── deploy/                    # NEW — configuration of record (invariant #11)
│   ├── docker-compose.yml     # cherry-picked + stripped from Prod-v1
│   ├── Dockerfile.api
│   ├── Caddyfile
│   ├── env/
│   │   └── .env.example
│   └── systemd/
│       └── docker-compose@sap.service
├── docs/                      # existing framework
├── nextjs/                    # NEW — the Vercel app
│   ├── package.json
│   ├── next.config.ts
│   ├── tsconfig.json
│   ├── prisma/
│   │   └── schema.prisma
│   ├── src/
│   │   ├── app/               # App Router routes per §2.1
│   │   ├── lib/
│   │   │   ├── db.ts          # Prisma singleton
│   │   │   ├── supabase.ts    # server + client helpers
│   │   │   ├── processor.ts   # HTTP client for the Python sidecar
│   │   │   └── storage.ts     # Supabase Storage helpers
│   │   ├── middleware.ts      # auth + consent gate (invariant #8)
│   │   └── components/        # UI primitives
│   └── .env.example
├── scripts/                   # Python — extended (§16.4)
├── tests/                     # Python — extended with contract tests
└── backups/
    └── python-legacy/         # Streamlit + start_scripts at cutover
```

The existing Python modules are **not moved, renamed, or rewritten** by this design. The only Python change on this branch is the new `app/processing_service/` module.

### 13.2 Disposition of `nextjs-impl` branch files

| File on `nextjs-impl` | Disposition on `refactor/nextjs-migration` |
|---|---|
| `package.json` | Carry forward; re-pin only if security advisories force it. |
| `next.config.ts`, `tsconfig.json`, `eslint.config.mjs`, `postcss.config.mjs` | Carry forward. |
| `src/app/layout.tsx`, `globals.css` | Carry forward; re-theme when final palette clears WCAG AA (§15.5). |
| `src/app/page.tsx` | **Delete.** Replace with RSC-first landing per §2.1. |
| `src/app/api/process/route.ts` | **Delete.** Replaced by `/api/uploads` + `/api/jobs/*` + `/api/webhooks/processor`. |
| `src/lib/pptx-utils.ts` | **Delete.** Python owns PPTX. |
| `src/lib/gemini.ts` | **Delete.** Python owns Gemini. |
| `src/lib/chroma.ts` | **Delete.** Python owns Chroma. |
| `AGENTS.md` ("this is NOT the Next.js you know") | Carry forward as `nextjs/AGENTS.md`. Augment with RSC/Client + Edge-forbidden rules. |
| `CLAUDE.md` | Delete — superseded by repo-root `AGENTS.md`. |
| `public/*` | Carry forward; swap marketing assets later. |
| `backups/python-legacy/` | Carry forward pattern; at cutover, Streamlit goes here. |

---

## 14. Cutover, Rollback, Data Migration (resolves preflight §14)

### 14.1 Phases

- **Phase 0 — parity staging.** Vercel `preview` deployments + a staging VM (or staging Compose project on the same VM, distinct ports). Runs through a fixed deck-fixture corpus (§15.2).
- **Phase 1 — private beta.** `access.brockportsigai.org/v2` routed to Vercel via a nginx split on the VM (or a Vercel rewrite on a non-production subdomain). Invited accessibility coordinators only.
- **Phase 2 — dual run.** `access.brockportsigai.org` swings to Vercel. Streamlit remains reachable on a private port (`8501` bound to `127.0.0.1`, SSH-tunnel only) for 30 days.
- **Phase 3 — sunset.** Streamlit systemd units uninstalled. `backups/python-legacy/` retains the code.

### 14.2 Cutover steps (sequence)

1. Freeze `Aggrement` (**[PENDING]** date).
2. Merge `refactor/nextjs-migration` → `Aggrement`-successor branch (new `main`).
3. Deploy Compose stack to VM: `docker compose -f deploy/docker-compose.yml up -d --build`.
4. Delete GCP firewall rules `allow-8501`, `allow-8001`.
5. Bind Streamlit systemd units to `127.0.0.1` (not disabled yet).
6. Flip DNS: `access.brockportsigai.org` A-record → Vercel.
7. Smoke test: `scripts/smoke_test.py --url https://access.brockportsigai.org --strict` (existing script).
8. Announce cutover.

### 14.3 Rollback (T < 24 h)

- **Full rollback:** revert DNS, re-add firewall rules, re-enable Streamlit systemd units. Max 10 minutes.
- **Partial rollback:** keep Vercel up, disable `/api/uploads` with a feature flag (`NEXT_PUBLIC_READ_ONLY=true`), route users to the private Streamlit via a secondary domain. Preserves in-flight review work in Postgres.

`docs/ops/SOP_ROLLBACK.md` gets a new section for these two paths — **[PENDING]** authoring (tracked as a post-design PR).

### 14.4 Data migration

| Data | Action |
|---|---|
| Existing Chroma collections on the VM | Discard. Re-index from new uploads. |
| Streamlit-era uploaded decks on the VM's `/tmp` | Discard. |
| `consent_responses.csv` | Import into `ConsentEvent` OR archive per IRB — **[PENDING]**. |
| Any in-flight Streamlit sessions at cutover | Discarded. Announced in the cutover notice. |

---

## 15. Testing, A11y QA, Observability (resolves preflight §15)

### 15.1 Test layers

| Layer | Tool | Scope |
|---|---|---|
| Python unit | `pytest` | Existing + new `processing_service/` orchestration tests. |
| Python contract | `pytest` + `fastapi.testclient` | `/jobs/*` request/response schemas. |
| TS unit | `vitest` | Prisma model invariants, middleware, Route Handler logic (mocking Python and Supabase). |
| TS integration | `playwright` | End-to-end `/upload → /review → /download` against a staging stack. |
| PPTX round-trip | `pytest` on Python side | One fixture per invariant: group shapes, WMF image, RGBA image, large deck, empty notes. |
| Accessibility (deck output) | Manual via PowerPoint Accessibility Checker on every fixture deck — rotating owner **[PENDING]**. |
| Accessibility (UI) | `axe-core` via Playwright on every page. CI fails on a11y violations above a baseline. |

### 15.2 Deck fixture corpus

Committed under `tests/fixtures/pptx/`. Each fixture targets one invariant:

- `groups.pptx` — nested group shapes (invariant #7).
- `wmf.pptx` — WMF/EMF image (invariant #5).
- `palette.pptx` — paletted PNG (invariant #5).
- `large.pptx` — 50 slides × 5 images (performance ceiling §7.2).
- `order.pptx` — intentionally shuffled shapes to catch `order_number` drift (invariant #1).

### 15.3 CI (GitHub Actions)

Extend `.github/workflows/preflight.yml` already shipped on `feature/agent-ops-framework`:

- **Python job** (existing): `doctor`, `check_invariants`, `preflight`, pytest.
- **TypeScript job** (new): `pnpm -C nextjs lint`, `pnpm -C nextjs build`, `tsc -p nextjs --noEmit`, `pnpm -C nextjs test`, Prisma schema validation.
- **Compose job** (new): `docker compose -f deploy/docker-compose.yml config` (lint only), Caddyfile format check.
- **Invariant expansion job** (new): extended `check_invariants.py` (§16.4).

### 15.4 Observability

- Every request logs `jobId`, `userId`, `route`, `duration_ms`, `status`.
- Python side logs `gemini_retry_count`, `gemini_total_delay_ms` per job.
- Postgres view `v_job_summary` aggregates per-day counts of each terminal state.

### 15.5 UI a11y baseline

- All text on the UI ≥ WCAG AA contrast (the current `nextjs-impl` dark purple gradient is retained **only if** it clears; otherwise re-theme is a blocker before v1).
- Keyboard-navigable upload, review, and download flows.
- Every image in the review grid has an `aria-label` sourced from the current description.
- Form controls have associated `<label>`s.
- The review editor is a plain `<textarea>`, not a contenteditable.
- Screen-reader test on at least one of NVDA / VoiceOver per release.

---

## 16. Framework Integration (resolves preflight §16)

### 16.1 Documents that must be updated after this design lands (same branch)

- `docs/PROJECT_OVERVIEW.md` — add v2 architecture section.
- `docs/AGENT_CONTEXT.md` — add Next.js + Supabase entries; mark Streamlit as legacy.
- `docs/PRODUCTION_ENVIRONMENT.md` — rewrite for split hosting + Compose stack.
- `docs/guardrails/INVARIANTS.md` — update §2, §4, §6, §8, §10, §11 per §8 above. Add invariant #12 (see §16.3).
- `docs/guardrails/BRANCHING.md` — retire `nextjs-impl`; formalize `refactor/nextjs-migration` → `main` flow; forbid lateral merges with `Prod-v1` beyond the compose cherry-pick.
- `docs/ops/SOP_DEPLOY.md` — split into SOP_DEPLOY_VERCEL and SOP_DEPLOY_BACKEND.
- `docs/ops/SOP_ROLLBACK.md` — §14.3 above.
- `docs/ops/SOP_SECRETS.md` — Vercel + VM `.env` flows.
- `scripts/` and `.github/workflows/preflight.yml` — §15.3.

### 16.2 New framework documents

- `docs/guardrails/NEXTJS_RULES.md` — RSC vs Client, Server Action vs Route Handler, Edge forbidden, secret-scope rules.
- `docs/guardrails/PYTHON_SERVICE_BOUNDARY.md` — explicit list of what the Python sidecar owns and what Next.js must never duplicate.
- `docs/ops/SOP_DEPLOY_VERCEL.md`, `docs/ops/SOP_DEPLOY_BACKEND.md` — replace the single SOP_DEPLOY with two.
- `docs/refactor/NEXTJS_MIGRATION_DESIGN.md` — this file.

### 16.3 New / modified invariants

- **Invariant #12 (new) — Service boundary.** TypeScript in this repo must not open a `.pptx`, must not call Gemini, must not talk to Chroma, and must not bypass the Python sidecar for any job-processing action. Enforced by a new `check_invariants.py` check (§16.4).
- **Invariant #2 (updated) — Python-only.** The cNvPr/@descr write is allowed only in Python files. Any TypeScript that references `cNvPr` or `descr=` triggers a hard fail.
- **Invariant #4 (strengthened).** Extends the existing "Chroma via FastAPI wrapper" rule to explicitly forbid any JavaScript Chroma client.
- **Invariant #6 (closed).** Data lives in Docker volume `chroma_data`. `./chroma/` under the repo is deleted post-cutover. Invariant text updated accordingly.
- **Invariant #8 (evolved).** Consent gate is now Profile flag + middleware + event log, not a Streamlit page + CSV.
- **Invariant #11 (satisfied).** Deploy configuration lives in `deploy/`.

### 16.4 `check_invariants.py` extensions

Additions to the existing checks (no change to passing checks on `feature/agent-ops-framework`):

| Key | What it checks |
|---|---|
| `ts_no_pptx` | No `.ts`/`.tsx` under `nextjs/` opens a .pptx (no `adm-zip`, no `jszip`, no `xml2js`, no `officegen`, no `pptxgenjs`). |
| `ts_no_gemini` | No `.ts`/`.tsx` imports `@google/generative-ai` or posts to `generativelanguage.googleapis.com`. |
| `ts_no_chroma` | No `.ts`/`.tsx` imports `chromadb` or matches a JS vector-DB client allowlist. |
| `ts_no_ooxml_strings` | No `.ts`/`.tsx` contains `cNvPr` or `a:blip`. |
| `compose_strips_streamlit` | `deploy/docker-compose.yml` has no service named `web` after cutover PR lands. |
| `no_public_internal_ports` | `deploy/docker-compose.yml` publishes only `80`/`443` on host. |
| `middleware_guards_consent` | `nextjs/src/middleware.ts` contains the consent check (regex probe). |

Each check added in this section is documented in `INVARIANTS.md` with rationale + check command.

### 16.5 Ownership & archival

- Refactor reviewer: **[PENDING MANUAL CONFIRMATION]** — single CODEOWNER for `nextjs/`, `deploy/`, `prisma/`.
- This document becomes historical and is moved to `docs/refactor/history/NEXTJS_MIGRATION_DESIGN.md` at cutover. Post-cutover changes do not edit this file; they land via regular `REFACTOR.md` PRs.

---

## 17. Resolved Strategic Bets (resolves preflight §17)

| Bet | Resolution | Consequence |
|---|---|---|
| Language monoculture vs polyglot | **Polyglot.** TypeScript UI + Python backend. | Long-term ops cost accepted; python-pptx fidelity preserved. |
| Runtime monoculture vs platform | **Split.** Vercel + self-hosted Docker. | Two deploy pipelines; trade-offs acknowledged in §11, §14. |
| State: cookies vs database | **Database (Supabase Postgres + Prisma).** | New infra accepted; enables resumable jobs, auth-gated downloads, IRB-grade consent tracking. |
| RSC-first vs Client-first | **RSC-first.** Client only for interactivity pockets. | Existing `nextjs-impl` page.tsx discarded. |
| Auth now vs never | **Now (Supabase Auth).** | Invariant #10.4 closes; per-user cost caps become enforceable. |
| Re-index vs migrate vectors | **Re-index.** | Zero-migration path; invariant #6's data-dir debt is paid off at the same time. |

---

## 18. Open items (**[PENDING MANUAL CONFIRMATION]**) — summary

These do not block any code work inside `nextjs/` or `deploy/`. They block cutover or IRB sign-off.

1. Hard feature-freeze date on `Aggrement`.
2. DNS cutover date for `access.brockportsigai.org`.
3. IRB retention policy and exact `ConsentEvent` field list.
4. Whether `consent_responses.csv` is imported or archived.
5. Subdomain for the Python API (`api.brockportsigai.org` assumed).
6. Per-user Gemini cost cap numbers (requests/day, bytes/day).
7. Owner of prompt changes going forward (prompt freeze anchor).
8. UI palette a11y sign-off (keep `nextjs-impl` theme vs. re-theme).
9. Baseline KPI measurements from live Streamlit (upload→download wall time, Gemini cost/deck, p95 TTI).
10. CODEOWNER for `nextjs/`, `deploy/`, `prisma/`.
11. Whether SUNY SSO replaces magic-link in v1.1 or later.
12. Whether Streamlit stays reachable on `127.0.0.1:8501` for 30 days post-cutover or is uninstalled immediately.

---

## 19. Halt

This document is the design. No `.tsx` or `.py` files were modified producing it. The next step is:

1. Human review of this file.
2. Per-decision confirmation of the **[PENDING]** items.
3. Then, and only then, scaffold `nextjs/`, `deploy/`, and the new `app/processing_service/` module in follow-up PRs, each using `docs/templates/REFACTOR.md` and referencing this document by section.
