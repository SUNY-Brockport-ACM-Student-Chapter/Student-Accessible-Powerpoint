# Processing Service

This package contains the FastAPI `/jobs/*` orchestration surface described in `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`.

The service must only orchestrate existing Python modules:

- pull uploaded `.pptx` blobs from Supabase Storage,
- call existing parsing/RAG/rebuild code,
- write status transitions back through the Next.js webhook,
- upload rebuilt decks to Supabase Storage.

Do not move PPTX parsing, OOXML traversal, Gemini calls, or Chroma access to TypeScript.

Run locally with:

```sh
uvicorn app.processing_service.app:app --reload --port 8000
```

Required environment:

- `PY_SERVICE_SHARED_SECRET`
- `NEXT_PUBLIC_APP_URL`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_UPLOADS_BUCKET`
- `SUPABASE_OUTPUTS_BUCKET`
- `GOOGLE_API_KEY`
- `CHROMA_API_URL`
