# Processing Service

This package will contain the FastAPI `/jobs/*` orchestration surface described in `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`.

The service must only orchestrate existing Python modules:

- pull uploaded `.pptx` blobs from Supabase Storage,
- call existing parsing/RAG/rebuild code,
- write status transitions back through the Next.js webhook,
- upload rebuilt decks to Supabase Storage.

Do not move PPTX parsing, OOXML traversal, Gemini calls, or Chroma access to TypeScript.
