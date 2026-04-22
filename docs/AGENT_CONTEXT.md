# AGENT_CONTEXT — Student-Accessible-Powerpoint

> **Purpose:** dense, structured context for AI coding agents. Ingest this file first. Pair with [`PROJECT_OVERVIEW.md`](./PROJECT_OVERVIEW.md) for narrative framing.
>
> **Last verified against:** `main` @ `1a31150` (Apache 2.0 license commit). Branch facts verified against `origin/*` refs.
>
> **Framework files** (added after this doc was written — check them before making changes):
> [`../AGENTS.md`](../AGENTS.md) · [`guardrails/INVARIANTS.md`](guardrails/INVARIANTS.md) · [`guardrails/CHANGE_CHECKLIST.md`](guardrails/CHANGE_CHECKLIST.md) · [`ops/SOP_DEPLOY.md`](ops/SOP_DEPLOY.md) · [`templates/`](templates/) · [`../scripts/`](../scripts/).

---

## 0. TL;DR for agents

- Product: Streamlit tool that adds WCAG/ADA-compliant alt text and speaker notes to `.pptx` decks, using Gemini vision + ChromaDB RAG.
- Stack: Python 3.10/3.11, Streamlit, python-pptx, Pillow, pytesseract, FastAPI, ChromaDB, Google `google-generativeai` (gemini-2.0-flash-lite), pydantic v2, ImageMagick (external binary).
- Three processes: Streamlit web (8501), FastAPI Chroma wrapper (8001), ChromaDB server (8000). The web app **never** imports `chromadb` directly — only HTTP.
- Primary accessibility write: `shape._element._nvXxPr.cNvPr.attrib["descr"] = alt_text` (native OOXML). Fallback: `shape.alternative_text = alt_text`.
- Entry point: `app/ppt_notes.py::main` (Streamlit). Orchestrator for the 4-stage UI (`upload` → `describe_images` → `final_processing` → `download`).
- Quick dev start: `python start_app.py` (choose option 3).
- Most feature-rich branch: `Aggrement` (consent wall, DOCX support, perf fixes, modular `pptx.py`/`word.py`, change log in `CATCH_UP.md`).
- Experimental TS port: `nextjs-impl` (Next.js 16 + React 19, incomplete; Python kept in `backups/python-legacy/`).

---

## 1. Repository map (branch `main`)

```
/
├── app/
│   ├── ppt_notes.py              # Streamlit app; orchestrates pipeline
│   ├── chroma-api/
│   │   ├── app.py                # FastAPI wrapper for ChromaDB (port 8001)
│   │   └── .env.example
│   ├── models/
│   │   ├── __init__.py           # empty
│   │   └── models.py             # Pydantic v2 data models
│   └── pptx_rag_quizzer/
│       ├── __init__.py           # empty
│       ├── utils.py              # parse_powerpoint(), OCR stub, format convert
│       ├── rag_core.py           # Gemini + ChromaHTTPClient + RAGCore
│       └── image.py              # Image class: 4-stage description pipeline
├── docs/
│   ├── PROJECT_OVERVIEW.md       # human narrative
│   └── AGENT_CONTEXT.md          # this file
├── start_app.py                  # dev launcher (Chroma API + Streamlit)
├── requirements.txt              # full server + app dependencies
├── requirements-app.txt          # app-only (no Chroma/FastAPI server)
├── RAG_INTEGRATION_README.md
├── LICENSE                       # Apache 2.0
├── .env.example                  # GOOGLE_API_KEY only
├── .gitattributes
└── .gitignore                    # ignores .env, *.pptx, /venv, /chroma-db, *.pyc
```

**Python package root:** `app/`. Modules import as `pptx_rag_quizzer.*` and `models.*`. Streamlit must be launched with `streamlit run app/ppt_notes.py` from repo root so that `app/` is on the import path.

---

## 2. Data model (`app/models/models.py`)

```python
class Type(Enum):
    image = "image"
    text  = "text"

class SlideItem(BaseModel):
    id: str
    slide_number: int
    content: str                  # for Image: the generated description; starts as "none"
    type: Type
    order_number: int             # position within the slide's reading order

class Image(SlideItem):
    image_bytes: bytes            # raw (or PNG-normalized) image blob
    extension: str
    def metadata(self) -> dict    # used for Chroma metadata

class Text(SlideItem):
    def metadata(self) -> dict

class Slide(BaseModel):
    id: str
    slide_number: int
    items: List[Union[Image, Text]]

class Presentation(BaseModel):
    id: str
    name: str
    slides: List[Slide]

class RAG_quizzer(BaseModel):     # currently unused, kept for parity with RAG-quizzer sibling project
    model_config = ConfigDict(arbitrary_types_allowed=True)
    id: str; name: str; presentation: Presentation; collection_id: str
```

**Critical invariant:** `order_number` is the *only* join key between a parsed `Image`/`Text` item and a `python-pptx` `shape`. `parse_powerpoint` increments `order_number` for BOTH text and image shapes, and `update_images_with_alt_text` (in the `Aggrement` branch's `pptx.py` — see §7) must mirror that exact sequence or alt text lands on the wrong shape.

**On the `Aggrement` branch**, `Image` adds:
```python
@field_serializer('image_bytes')
def serialize_image_bytes(self, value: bytes, _info) -> str:
    return base64.b64encode(value).decode('utf-8')
```
…and additional `WordDocument` / `WordSection` / `WordText` / `WordImage` models for DOCX support.

---

## 3. Process topology

```
Streamlit (web, 8501) ──HTTP──▶ chroma-api (FastAPI, 8001) ──HTTPClient──▶ chromadb (8000, persistent)
         │
         └─HTTPS──▶ generativelanguage.googleapis.com  (Gemini)
```

- `web` never imports `chromadb`. See `rag_core.ChromaHTTPClient`.
- `chroma-api` is a **thin** REST wrapper (`app/chroma-api/app.py`). Endpoints:
  - `GET  /health`
  - `GET  /collections` (list)
  - `POST /collections` (create, body: `{name, metadata}`)
  - `DELETE /collections/{name}`
  - `GET  /collections/{name}/exists`
  - `POST /collections/{name}/add` (body: `{documents, metadatas, ids}`)
  - `POST /collections/{name}/query` (body: `{query_texts, n_results, include}`)
  - `POST /collections/{name}/get` (body: `{include}`)
- `make_json_serializable()` converts numpy arrays returned by Chroma into plain lists so FastAPI can serialize them.

### Environment variables (authoritative list)

| Var | Consumer | Default | Notes |
|---|---|---|---|
| `GOOGLE_API_KEY` | web (`rag_core.get_llm_model`, `ppt_notes.ExtractText_LLM`) | **required** | Gemini auth |
| `CHROMA_API_URL` | web (`RAGCore.__init__`) | `http://localhost:8001` | Points web → FastAPI wrapper |
| `CHROMA_SERVER_HOST` | chroma-api | `localhost` | Where FastAPI finds ChromaDB |
| `CHROMA_SERVER_HTTP_PORT` | chroma-api | `8000` | ChromaDB port |
| `API_HOST` / `API_PORT` | chroma-api | `0.0.0.0` / `8001` | FastAPI bind |
| `ACME_EMAIL` | caddy (Prod-v1) | — | Let's Encrypt contact |

---

## 4. Primary control flow (`main` branch, `app/ppt_notes.py`)

Streamlit `session_state.processing_stage` state machine:

```
upload ──▶ describe_images ──▶ final_processing ──▶ download
   │                                                    │
   └──── Quick Generate & Download (skip review) ───────┘
```

### `upload` stage
1. `uploaded_file.read()` → `file_bytes`.
2. `parse_powerpoint_file(file_bytes, name)` → `PresentationModel` (text + image blobs, `content='none'` for images).
3. `RAGCore().create_collection(presentation_model)` → `collection_id` (initial collection: TEXT ONLY; images skipped while `content == 'none'` or `__DELETED__`).
4. Store everything in `st.session_state`; jump to `describe_images`.

**Quick-generate path** bypasses human review: iterates all image items, calls `ImageProcessor.describe_image` for each, strips the `"Description: "` prefix, then rebuilds the collection and jumps to `final_processing`'s equivalent.

### `describe_images` stage
- Images paged in batches of `batch_size=5`.
- `batch_ready` = all images in batch have non-empty, non-`'none'`/`'null'` content.
- If not ready: call `ImageProcessor.describe_image()` for each un-described image (this is where the 4-stage RAG pipeline fires).
- If ready: render `st.image` + `st.text_area` for edit; user navigates `Previous/Next/Finish`.

### `final_processing` stage
1. `rag_core.remove_collection(old_collection_id)`.
2. `rag_core.create_collection(presentation_model)` → **enhanced** collection (now includes image descriptions as documents).
3. Write PPTX bytes to `temp_{filename}`.
4. `process_powerpoint_with_rag_enhanced(temp_path, output_path, ...)` — see §5.
5. `os.remove(temp_path)`; jump to `download`.

### `download` stage
- `st.download_button` serves `output_path` bytes with MIME `application/vnd.openxmlformats-officedocument.presentationml.presentation`.
- Reset clears session state and deletes `output_path`.

---

## 5. Accessibility writer — `process_powerpoint_with_rag_enhanced`

Defined in `app/ppt_notes.py`. Opens the PPTX with `python-pptx`, iterates slides and shapes, and writes three things:

### 5.1 Alt text (for every `MSO_SHAPE_TYPE.PICTURE`)
- Find matching description in `presentation_model` by comparing `shape.image.blob` to `item.image_bytes` (byte-equality). If not found or description is empty/`'none'`, use the fallback `"Image content - detailed description not available"`.
- Write primary path:
  ```python
  shape._element._nvXxPr.cNvPr.attrib["descr"] = alt_text
  ```
- Write fallback path:
  ```python
  shape.alternative_text = alt_text
  ```
- Alt text is the full AI description; slide-deck-wide cap via `create_accessible_alt_text` (125-char guideline) is *defined but not currently invoked* on this path — the pipeline writes the full description directly.

### 5.2 Chart / Table alt text
- `MSO_SHAPE_TYPE.CHART` → `shape.alternative_text = f"Chart on slide {N} - {slide_title}"`
- `MSO_SHAPE_TYPE.TABLE` → `shape.alternative_text = f"Table on slide {N} - {slide_title}"`

### 5.3 Slide notes
- Collect per-shape snippets into `notes_texts`.
- If a `collection_id` is present, call `rag_core.get_context_from_slide_number(slide_index+1, collection_id)` and feed those docs into `generate_enhanced_notes_with_context` → `rag_core.prompt_gemini(...)` → enhanced notes.
- Assign:
  ```python
  notes_slide = slide.notes_slide                # auto-creates if missing
  notes_slide.notes_text_frame.text = "\n".join(notes_texts)
  ```

### 5.4 Persistence
- `prs.save(output_path)` where `output_path = f"accessible_{file_name}"`.

---

## 6. Parser — `parse_powerpoint` (`app/pptx_rag_quizzer/utils.py`)

```python
def parse_powerpoint(file_object, file_name) -> Presentation:
    prs = pptx_lib(file_object)
    for slide_idx, slide in enumerate(prs.slides):
        order_number = 0
        # 1. Speaker notes (if any) as first Text item (order 0)
        # 2. Sort shapes by (top, left) for reading order
        for shape in sorted(slide.shapes, key=lambda x: (x.top or 0, x.left or 0)):
            if shape.has_text_frame and shape.text_frame.text:
                # append Text(content=..., order_number=order_number); order_number += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # append Image(image_bytes=shape.image.blob, extension=shape.image.ext,
                #              content='none', order_number=order_number); order_number += 1
```

Current `main` does **not** recurse into groups, diagrams, charts, or background fills — those are Aggrement-branch improvements (§7).

`ExtractText_OCR(img_bytes)` is currently stubbed; the real Tesseract call is commented out and the function returns `"<THIS OCR TEXT IS IN DEVELOPMENT AND SHOULD BE DISREGARDED>"`. OCR is effectively disabled on `main`.

---

## 7. RAG image description — `app/pptx_rag_quizzer/image.py`

Class `Image` (careful: same name as `models.models.Image` — import as `ImageProcessor` in callers).

```python
class Image:
    def __init__(self, rag_core: RAGCore):
        self.rag_core = rag_core
        self.chat_history = []       # rolling, max 10
        self.context_cache = {}      # cache_key -> {description, timestamp}
        self.lambda_index = {}
        self.cache_ttl = 3600
```

### 4-stage `describe_image(image_bytes, image_format, slide_number, collection_id, use_chat=True)`
1. **OCR** — `ocr_image()` → `ExtractText_OCR()` (currently a stub, see §6).
2. **Enhanced description** — `get_enhanced_description()`:
   - Build prompt with `<ocr_text>`, `<slide_context>` (from `rag_core.get_context_from_slide_number`), and optional `<chat_history>`.
   - `rag_core.prompt_gemini_with_image(prompt, image_bytes, image_format, max_output_tokens=200)`.
   - Expected output prefix: `"Description: "`.
3. **Lambda-Index context retrieval** — `get_context_with_lambda_index()`:
   - `_build_lambda_query()` constructs a text query from description + top-10 key terms (stop-words filtered, len > 3).
   - `rag_core.query_collection(...)` with `n_results=3`.
   - `_rank_context_with_lambda()` scores each retrieved doc with `_calculate_lambda_score()`:
     - `+0.1` per overlapping term.
     - `+0.3` for image-type metadata, `+0.1` for slide-number metadata.
     - Length-normalized, capped at 1.0, threshold `> 0.3` for inclusion.
4. **Final description** — `get_final_description_with_chat()`:
   - Gemini refines the enhanced description using context + chat history. 1–3 sentences, prefixed `"Description: "`.

Response may be JSON (`{output: {Description: ...}}` or `{Description: ...}`); the code unwraps both.

**Caching:** `cache_key = f"{md5(image_bytes)}_{slide_number}_{collection_id}"`, 1-hour TTL.

**Chat-history:** up to 10 most-recent "Enhanced description: …" / "Final description: …" entries, used to stabilize vocabulary across slides.

---

## 8. `RAGCore` — `app/pptx_rag_quizzer/rag_core.py`

```python
class RAGCore:
    def __init__(self, chroma_api_url=None):
        self.llm_model = get_llm_model()                 # cached gemini-2.0-flash-lite
        self.chroma_api = get_chroma_http_client_instance(...)

    def create_collection(self, data: Presentation) -> str       # returns f"ppt_collection_{uuid8}"
    def remove_collection(self, collection_id: str)              # returns dict (already .json()'d)
    def query_collection(self, query_text, collection_id, n_results=1) -> dict
    def get_random_slide_context(self, collection_id) -> dict
    def get_random_slide_with_image(self, collection_id) -> dict|None
    def get_context_from_slide_number(self, slide_number, collection_id) -> dict
    def prompt_gemini(self, prompt, max_output_tokens=200) -> str
    def prompt_gemini_with_image(self, prompt, image_bytes, image_format='png', max_output_tokens=200) -> str
```

### Collection layout
- **One Chroma document per slide**. Each document is `" ".join(all text and image-description strings for that slide)`.
- **Metadata per document** includes per-item fields `item_{n}_type`, `item_{n}_slide_number`, `item_{n}_order_number`, and for images `item_{n}_image_extension`, `item_{n}_has_image`, `item_{n}_image_size`. Plus top-level `slide_number` and `slide_id`.
- Image items with `content == "__DELETED__"` are skipped at collection build time.

### Gemini call conventions
- `prompt_gemini_with_image` normalizes every image through PIL:
  - `P` mode with `"transparency"` → convert to `RGBA`.
  - `RGBA`/`LA` → paste onto white RGB background.
  - Re-encode as PNG in memory.
- Retries: `max_retries=3`, `delay=1s`. On `"Resource has been exhausted"` — sleeps `quota_refill_delay=60s`.

### `ChromaHTTPClient` — note that its methods return `response.json()` (already-parsed dicts). Do not call `.json()` on `RAGCore.remove_collection(...)` return values (this was a historical bug — see `Aggrement`'s `CATCH_UP.md`).

---

## 9. Branches — exhaustive map

### `main` (current; Apache 2.0)
- HEAD: `1a31150 Include Apache License 2.0`.
- State described throughout §§1–8.

### `RAG-integration-branch`
- Oldest. Single `Dockerfile`, `docker-compose.yml`, `DEPLOYMENT.md`.
- Python 3.10 slim base, runs `streamlit run new-app/ppt_notes.py` (note: path is `new-app/`, not `app/` — the layout differs).
- Relevant only for historical deployment references.

### `Prod-v1`
- Multi-service deployment scaffolding:
  - `Dockerfile.api` — FastAPI wrapper container (python:3.11-slim + build-essential).
  - `Dockerfile.web` — Streamlit container (python:3.11-slim + tesseract-ocr + libjpeg62-turbo-dev + zlib1g-dev + curl).
  - `docker-compose.yml` — four services: `chroma`, `chroma-api`, `web`, `caddy` on an implicit shared network. `chroma` persists to `./chroma-db:/chroma/chroma`.
  - `Caddyfile` — proxies `access.brockportsigai.org` → `web:8501`, ACME TLS via `{$ACME_EMAIL}`.
- Same app code as `main` at commit `e6fcd8f`.

### `Aggrement` — most feature-rich
HEAD `ee8606d`. Key additions:

- **IRB consent wall** (`app/ppt_notes.py`, stage `consent`):
  - Four radio options (agree / agree+email / no / under-18).
  - Under-18 → `processing_stage = 'blocked'`, `st.stop()`.
  - Agreed-with-email → appended to `consent_responses.csv` via `save_consent_email(email, choice)` (creates CSV with header `timestamp_utc,email,choice`).
- **AI-content warning banner** rendered at top of page.
- **Module split inside `app/pptx_rag_quizzer/`:**
  - `utils.py` — only `ExtractText_OCR`, `clean_text`, `clean_text_with_llm`, `convert_image_to_png_or_jpg` (uses subprocess + `magick`/`convert`; returns `(None, None)` for failed WMF/EMF).
  - `pptx.py` — `parse_powerpoint` (recursive for groups/diagrams/charts/background), `generate_accessible_notes`, `rebuild_presentation_with_accessible_features` with inner `update_images_with_alt_text`. **Entry point for writing accessibility:** `rebuild_presentation_with_accessible_features(presentation_model, powerpoint_file_bytes_io)`.
  - `word.py` — `parse_word_document`, `rebuild_word_document_with_accessible_features`. Uses `python-docx`; walks paragraphs, nested tables, headings (`_is_heading` checks `style.name.lower().startswith("heading")`). Images resolved by `a:blip/@r:embed` → `document.part.related_parts`. Alt text written to `wp:docPr/@descr`.
- **Extended models** (§2): `WordDocument`/`WordSection`/`WordText`/`WordImage`, `base64` serializer on `Image.image_bytes`.
- **Performance:**
  - `rebuild_presentation_with_accessible_features` creates `rag_core = RAGCore()` once and passes it down (~70% speedup per `CATCH_UP.md`).
  - `generate_accessible_notes` uses `max_output_tokens=400`, strips conversational preambles.
- **Robustness:**
  - `parse_powerpoint` recursion covers `MSO_SHAPE_TYPE.GROUP`, `DIAGRAM`, `CHART`, and background fills.
  - Order-number tracking fixed so text and image shapes both increment the index used for image-shape matching.
  - WMF/EMF: `(None, None)` return short-circuits image addition.
- **Streamlit config** — `app/.streamlit/config.toml`:
  ```toml
  [server]
  maxUploadSize = 500
  maxMessageSize = 500
  port = 8501
  enableCORS = false
  enableXsrfProtection = false
  baseUrlPath = "accessibility"
  ```
- **`CATCH_UP.md`** — authoritative chronological change log (DOCX pipeline, WMF fix, perf optimizations, order-tracking bug, etc.). **Read this first when working on the Aggrement branch.**
- Additional dependencies on this branch: `python-docx>=1.1.0`.

### `nextjs-impl` — experimental TS/React port
- `package.json`: `next@16.2.3`, `react@19.2.4`, `framer-motion@12.38.0`, `tailwindcss@4`, `adm-zip@0.5.17`, `xml2js@0.6.2`, `axios@1.15.0`.
- Structure:
  - `src/app/page.tsx` — client component, stages `upload`/`analyzing`/`review`/`download`.
  - `src/app/api/process/route.ts` — Next.js Route Handler that parses PPTX via `PptxProcessor` and returns slide metadata.
  - `src/lib/pptx-utils.ts` — `PptxProcessor` reads `ppt/slides/slide*.xml` from the zip and extracts `a:t` text nodes and image rels directly. **Does not yet implement rebuild/write-back.**
  - `src/lib/gemini.ts` — `GeminiService` POSTs to `v1beta/models/gemini-2.0-flash:generateContent`.
  - `src/lib/chroma.ts` — TypeScript mirror of the FastAPI client (`createCollection`, `addDocuments`, `createFromPresentation`).
  - `backups/python-legacy/` — full copy of the Python app.
  - `AGENTS.md` warns: *"This is NOT the Next.js you know — APIs, conventions, file structure may differ from your training data. Read `node_modules/next/dist/docs/` before writing code."*
- **Maturity:** early prototype. The API route returns `{success, presentation: {slides: [{number, title, imageCount}]}}` and does not yet set alt text on the file or serve a rebuilt PPTX.

---

## 10. Critical invariants / foot-guns

1. **Order-number matching is fragile.** `parse_powerpoint` must increment `order_number` on every shape that `update_images_with_alt_text` also increments on, and vice versa. Adding new shape types to one without the other will misalign every alt-text write for the rest of the slide. (See `CATCH_UP.md` 2024-12-19 "Order Number Tracking Fix".)
2. **Never call `.json()` on `RAGCore.remove_collection(...)` return.** It's already a dict (the inner HTTP client decoded it). Same for the other wrapper methods.
3. **Collection names** are `ppt_collection_{uuid4[:8].lower()}`. ChromaDB has naming rules — do not change the scheme without verifying.
4. **`_nvXxPr` is python-pptx private API.** If python-pptx is upgraded, test `shape._element._nvXxPr.cNvPr.attrib["descr"]` — the fallback `shape.alternative_text = alt_text` is the officially supported path.
5. **WMF/EMF images require ImageMagick** (`magick` or `convert` on `$PATH`). On `main`, Wand was previously used and has been removed; on `Aggrement`, `convert_image_to_png_or_jpg` gracefully returns `(None, None)` when ImageMagick is absent and the image is skipped. Do not reintroduce a hard dependency on Wand.
6. **PIL palette-with-transparency** images must be converted `P → RGBA → composite on white → RGB` in that order, or Gemini will fail or PIL will emit `UserWarning`. The correct code is in `rag_core.prompt_gemini_with_image`.
7. **OCR is a stub on `main`.** `ExtractText_OCR` returns a placeholder string. The `pytesseract` call is commented out. Do not assume OCR text is meaningful unless you re-enable it.
8. **Gemini quota** errors trigger 60-second sleeps. Long runs can stall — batch size of 5 is a guardrail.
9. **Session state keys** used by Streamlit (do not rename silently): `processing_stage`, `presentation_model`, `rag_core`, `image_processor`, `collection_id`, `current_batch`, `batch_size`, `output_path`, `uploaded_file_name`, `file_bytes`. On `Aggrement`: also `consent_completed`, `consent_choice`, `consent_email`, `new_presentation_model`.
10. **Streamlit rerun semantics**: every `st.rerun()` re-enters `main()`; the stage machine is driven entirely by `st.session_state.processing_stage`. Don't introduce blocking loops outside of `st.spinner` contexts.
11. **`requirements.txt` vs `requirements-app.txt`.** The app-only file omits `chromadb`, `fastapi`, `uvicorn`, `numpy`; use it for the web container. The full file is for the `chroma-api` container.
12. **`.env` is gitignored; `.env.example` has only `GOOGLE_API_KEY`** — the Chroma env vars live in `app/chroma-api/.env.example`.

---

## 11. Accessibility specification (authoritative for agents)

Target: **WCAG 2.1 Level AA**, per ADA Title II rule (28 CFR Part 35, 2024 amendment).

| WCAG SC | Level | How this project fulfills it | Module |
|---|---|---|---|
| 1.1.1 Non-text Content | A | Gemini-generated alt text written to native `cNvPr/@descr` for every extracted image, diagram, chart image, and background fill. | `ppt_notes.process_powerpoint_with_rag_enhanced` + `pptx_rag_quizzer/image.Image` |
| 1.3.1 Info & Relationships | A | Reading order preserved by `(top, left)` sort; slide notes regenerated in Markdown with headings + bullets. | `pptx_rag_quizzer/utils.parse_powerpoint` + `generate_enhanced_notes_with_context` |
| 1.3.2 Meaningful Sequence | A | Same top-to-bottom / left-to-right extraction order; `order_number` preserved end-to-end. | `utils.parse_powerpoint` |
| 1.4.5 Images of Text | AA | OCR text (when enabled) passes through Gemini so that raster-baked text is echoed in the description. | `pptx_rag_quizzer/image.Image.ocr_image` (stubbed on main) |
| 2.4.2 Page Titled | A | Slide title extracted and included in notes as `## Slide N: {Title}`. | `ppt_notes.process_powerpoint_with_rag_enhanced` |
| 2.4.6 Headings and Labels | AA | Notes are generated with explicit markdown headings; chart/table alt text includes slide title context. | Notes generator prompt in `ppt_notes.generate_enhanced_notes_with_context` |
| 3.1.5 Reading Level | AAA (partial) | Prompt asks for "clear, concise explanations"; AI-content warning prompts human verification. | Aggrement `pptx.generate_accessible_notes` prompt |
| 4.1.2 Name, Role, Value | A | All alt text written to standard OOXML (`cNvPr/@descr`) — no custom metadata. | python-pptx shape mutation |

**Known gaps / not yet implemented:**
- 1.4.3 Contrast (Minimum) — not inspected.
- 1.2.x Media-related SCs (captions, audio description) — not generated.
- 2.4.4 Link Purpose (In Context) — hyperlinks not rewritten.
- On-slide accessibility tab-order (the order a screen reader walks *visually* on a slide) is not actively reordered; only reading-order of *extraction* is stable.

When adding features, agents should:
- Preserve native OOXML attributes (never invent custom namespaces).
- Always populate both the primary path (`cNvPr/@descr`) and the fallback (`shape.alternative_text`).
- Keep generated text under 125 characters *for alt text* when possible (current writer passes the full description — if you add trimming, use `create_accessible_alt_text` which already exists in `ppt_notes.py` as an unused helper).

---

## 12. Test & verification notes

- **No formal test suite exists** in any branch. `app/test.py` is referenced in `Aggrement`'s `CATCH_UP.md` but not present in the tracked tree.
- Quick smoke-test flow: launch via `python start_app.py` (choose 3) → upload `sample.pptx` → use "Quick Generate & Download" → open output in PowerPoint → Review → Check Accessibility.
- Screen-reader end-to-end: open output in PowerPoint → run the built-in Accessibility Checker; run NVDA in Presenter View to verify alt text is announced.
- ChromaDB debug: `curl http://localhost:8001/health`, `curl http://localhost:8001/collections`.

---

## 13. Common extension recipes

- **Switch LLM provider.** Replace `get_llm_model()` in `rag_core.py` and adapt `prompt_gemini*`. Leave the `RAGCore` public surface unchanged.
- **Add a new accessibility transform** (e.g., heading style enforcement). Insert a new step in `process_powerpoint_with_rag_enhanced` (main) or `rebuild_presentation_with_accessible_features` (Aggrement). Always increment `order_number` consistently if you consume shapes.
- **Persist descriptions across sessions.** Add a `descriptions` table keyed on `md5(image_bytes)` (same hash already used for caching in `image.Image`).
- **Add a new file format.** Mirror the Aggrement `word.py` module: `parse_X_document` → `{Model}Document` → `rebuild_X_document_with_accessible_features`. Ensure an order-number join key exists.
- **Move off ChromaDB.** Only `RAGCore` and `ChromaHTTPClient` need replacing; the `chroma-api` service can be rebuilt against any vector DB. The web app's contract with the wrapper is six HTTP endpoints (§3).

---

## 14. Quick reference: file → purpose

| File | Purpose | Size |
|---|---|---|
| `app/ppt_notes.py` | Streamlit app & accessibility writer (`main`) | 27 KB |
| `app/pptx_rag_quizzer/utils.py` | PPTX parser + OCR stub + format convert | ~4 KB on `main`; ~5 KB on `Aggrement` |
| `app/pptx_rag_quizzer/rag_core.py` | LLM + Chroma HTTP client | 18 KB |
| `app/pptx_rag_quizzer/image.py` | 4-stage RAG image description | 23 KB |
| `app/pptx_rag_quizzer/pptx.py` | (Aggrement only) rebuild pipeline | 23 KB |
| `app/pptx_rag_quizzer/word.py` | (Aggrement only) DOCX support | 12 KB |
| `app/models/models.py` | Pydantic models | 1–3 KB |
| `app/chroma-api/app.py` | FastAPI wrapper | 6 KB |
| `start_app.py` | dev launcher | 3.7 KB |
| `requirements.txt` / `requirements-app.txt` | pins | <1 KB each |
| `RAG_INTEGRATION_README.md` | narrative feature doc | 8 KB |
| `docs/PROJECT_OVERVIEW.md` | human-readable overview | — |
| `docs/AGENT_CONTEXT.md` | **this file** | — |

---

## 15. Stable fact list (agents: trust these)

- License: **Apache 2.0** (`LICENSE`, added in commit `1a31150`).
- Organization: **SUNY Brockport ACM Student Chapter**.
- LLM model ID: `gemini-2.0-flash-lite` (both `ppt_notes.py` and `rag_core.py`).
- PPTX library: `python-pptx`. DOCX library: `python-docx` (Aggrement only).
- Vector DB: ChromaDB, accessed only via `http://…:8001` (FastAPI wrapper). Web process does NOT import `chromadb`.
- Default batch size: 5 images.
- Default notes token budget: `main` 500, `Aggrement` 400.
- Default RAG retrieval: `n_results=3`, Lambda score threshold `0.3`.
- Default image cache TTL: 3600 s.
- Default chat-history window: 10 messages.
- Native alt-text attribute written: `a:cNvPr@descr` on each `pic` element.

---

*When in doubt: consult `Aggrement/CATCH_UP.md` for the deepest history of what was tried, what broke, and why the code looks the way it does.*
