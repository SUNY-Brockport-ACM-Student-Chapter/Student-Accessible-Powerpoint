# Invariants

Things in this repo will silently break if you violate them. None of these are enforced by the compiler. Most of them are enforced weakly (or not at all) by tests. Read this file before touching parsing, rebuild, Gemini calls, ChromaDB wiring, or `.pptx` XML.

Each invariant has:
- **Rule** — what must stay true.
- **Why** — what goes wrong if it doesn't.
- **Where** — the files it lives in.
- **Check** — a command or `grep` to validate.

---

## #1 — `order_number` is the universal join key

**Rule.** Every `SlideItem` produced by `parse_powerpoint` / `parse_word` has an `order_number` that maps 1:1 to a shape/element's position in the source document. Rebuild loops must rely on `order_number`, **not** list index, **not** shape iteration order.

**Why.** `python-pptx`'s `slide.shapes` iteration order is *not* guaranteed stable across edits. If you assign alt text by index, you will place it on the wrong shape. This has happened and it is undetectable without visual QA.

**Where.** `app/pptx_rag_quizzer/utils.py` (parse), `app/ppt_notes.py :: process_powerpoint_with_rag_enhanced` (rebuild), `app/models/models.py :: SlideItem.order_number`.

**Check.**

```bash
rg -n "order_number" app/
python scripts/check_invariants.py --only order_number
```

---

## #2 — Alt text goes on `cNvPr/@descr`

**Rule.** When writing alt text to an image shape, set the native XML attribute `cNvPr/@descr` via lxml. Also set `shape.alternative_text` as a fallback for python-pptx tooling.

**Why.** Screen readers (NVDA, JAWS, VoiceOver) read `cNvPr/@descr`. `python-pptx`'s `alternative_text` property sometimes writes to a different location depending on shape type, and is silently ignored by some clients.

**Where.** Varies by branch:

- `main`, `RAG-integration-branch`: `app/ppt_notes.py` (rebuild loop lives here).
- `Aggrement`: `app/pptx_rag_quizzer/pptx.py :: rebuild_presentation_with_accessible_features`.
- `Prod-v1`: follows one of the above depending on the cherry-pick history.

The `check_alt_text_xml` invariant scans all of `app/` for a line that writes `descr` onto `cNvPr`, so moving the rebuild loop between modules is fine — *deleting* the XML write is not.

**Check.**

```bash
rg -n "cNvPr|descr=" app/
python scripts/check_invariants.py --only alt_text_xml
```

---

## #3 — Gemini rate-limit: 60 s sleep on `ResourceExhausted`

**Rule.** On `google.api_core.exceptions.ResourceExhausted` (HTTP 429), the call path must sleep **60 seconds** and retry. Do not reduce this interval, do not remove the retry.

**Why.** Gemini free-tier and flash-lite models rate-limit per minute. Shorter backoff busy-loops the quota; no retry drops slides silently.

**Where.** `app/pptx_rag_quizzer/rag_core.py` (retry decorator / loop).

**Check.**

```bash
rg -n "ResourceExhausted|sleep\(60\)|time.sleep" app/pptx_rag_quizzer/
```

---

## #4 — ChromaDB access goes through the FastAPI wrapper

**Rule.** Streamlit code must call ChromaDB via HTTP to `chroma-api` (the FastAPI wrapper at `:8001`), not via a direct `chromadb.HttpClient(...)` or `chromadb.PersistentClient(...)`.

**Why.** The wrapper normalizes request shape, handles retries, and isolates dependency upgrades. Bypassing it couples Streamlit to the Chroma client version directly and has broken prod historically when `chromadb` major versions shifted.

**Where.** `app/pptx_rag_quizzer/rag_core.py :: ChromaHTTPClient`. In production the wrapper listens at `http://127.0.0.1:8001` (see [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §3).

**Check.**

```bash
rg -n "chromadb\.(Http|Persistent)Client" app/ | grep -v chroma-api
# expect: only chroma-api/app.py
```

---

## #5 — Image normalization before hashing or Gemini

**Rule.** Any image bytes flowing into the RAG pipeline must be normalized to PNG or JPG with RGB (not `P` / `RGBA`) color mode before they hit Gemini or ChromaDB hashing.

**Why.** WMF/EMF and paletted PIL images break Gemini's MIME sniffing (400 errors) and cause opaque/black output. Also breaks base64 round-tripping in `Image.image_bytes`.

**Where.** `app/pptx_rag_quizzer/utils.py :: convert_image_to_png_or_jpg`.

**Check.**

```bash
rg -n "convert_image_to_png_or_jpg|mode\s*==\s*['\"]P['\"]|WMF|EMF" app/
```

---

## #6 — `chroma/` vector data directory is not to be cleaned in prod

**Rule.** On the production VM, `./chroma/` inside the repo holds persistent vector data. Do **not**:
- add it to `.gitignore` in a way that it gets `git clean`-ed,
- `rm -rf chroma/` on the VM,
- include `chroma/` in a Docker image COPY for dev images.

**Why.** It takes hours to rebuild on large decks, and a wipe has no warning. The path is also hard-coded in `start_scripts/chromadbd.sh`.

**Where.** `/home/mattarama443/Student-Accessible-Powerpoint/chroma/` on prod. `.gitignore` currently lists `/chroma-db` (a different, older path). Confirm both.

**Planned fix.** Move Chroma data outside the repo tree (`/var/lib/chroma-saac/`). Tracked as tech debt §10 below.

**Check.**

```bash
grep -E "^/?chroma" .gitignore
```

---

## #7 — `python-pptx` shape traversal must handle groups recursively

**Rule.** Any loop over `slide.shapes` that inspects images or text must recurse into group shapes (`shape.shape_type == MSO_SHAPE_TYPE.GROUP`).

**Why.** Many real-world decks group images with captions. A non-recursive loop silently drops everything inside groups — no alt text, no RAG indexing.

**Where.** `app/pptx_rag_quizzer/utils.py`. On the `Aggrement` branch, `app/pptx_rag_quizzer/pptx.py`.

**Check.**

```bash
rg -n "MSO_SHAPE_TYPE\.GROUP|is_group|\.shapes\b" app/
```

---

## #8 — Consent gate is IRB-mandated on `Aggrement`

**Rule.** On the `Aggrement` branch, the consent screen in `app/ppt_notes.py` must be rendered before any file upload UI. Writing to `consent_responses.csv` must continue for every "Agree" click. Do not bypass, shorten, or remove without an IRB amendment.

**Why.** SUNY Brockport IRB approval for the research deployment depends on this gate. Removing it breaks the human-subjects protocol.

**Where.** `app/ppt_notes.py` (Aggrement branch). `consent_responses.csv` on prod.

**Check.**

```bash
git show Aggrement:app/ppt_notes.py | rg -n "consent|IRB|agree"
```

---

## #9 — Branch is not landscape, it is a pipeline

**Rule.** Do not merge `Prod-v1`, `Aggrement`, or `nextjs-impl` laterally into each other. See [`BRANCHING.md`](BRANCHING.md).

**Why.** Each branch encodes a different deployment model (Docker vs bare-metal vs Next.js). Cross-merging has historically created subtle 3-way conflicts in `requirements.txt`, config, and parsing.

---

## #10 — Known weaknesses (tech debt)

These are *not yet* invariants — they are known gaps. Closing them is welcome.

| # | Gap | Link |
|---|---|---|
| 10.1 | Secrets stored as plaintext `.env` on the VM | [`../ops/SOP_SECRETS.md`](../ops/SOP_SECRETS.md) §5 |
| 10.2 | `start_scripts/chromadbapi.sh` has `Accessibility` typo (masked by systemd `WorkingDirectory`) | [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §3.3 |
| 10.3 | `start_scripts/` and `.streamlit/` live on the VM only — not in git | [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §6 |
| 10.4 | GCP firewall `allow-8501` / `allow-8001` expose internal services publicly | [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §8.4 |
| 10.5 | `chroma/` vector data inside repo tree (risk of `git clean`) | §6 above |
| 10.6 | No CI/CD — Jenkins exists on a separate box, unconfigured | [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §5 |
| 10.7 | Single prod instance, no LB, no autoscaler | [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §8.7 |
| 10.8 | Nginx and systemd unit files not version-controlled | [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md) §6 |

Each is a legitimate target for a [`REFACTOR.md`](../templates/REFACTOR.md).

---

## #11 — Configuration of record (emerging invariant)

**Rule (aspirational).** All nginx site configs, all systemd unit files, all startup shell scripts, and the Streamlit config must live in the repo under a `deploy/` directory with an idempotent sync script. Any change on the VM must be mirrored back via PR.

**Status.** Not yet implemented. This is the #1 framework-level refactor we owe ourselves. Until then, any change to these files on the VM **must** also land in an issue with label `drift-upstream`.
