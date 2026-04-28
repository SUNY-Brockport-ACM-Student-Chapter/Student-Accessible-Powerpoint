# Next.js Migration Agent Notes

This app follows `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`.

- React Server Components are the default. Add `"use client"` only for local browser state, refs, or event handlers.
- Route Handlers and Server Actions must use the Node runtime. Do not opt into Edge runtime.
- Never expose `SUPABASE_SERVICE_ROLE_KEY`, `DATABASE_URL`, `DIRECT_URL`, or `PY_SERVICE_SHARED_SECRET` to Client Components.
- Do not parse `.pptx`, traverse OOXML, call Gemini, or talk to Chroma from TypeScript. Those responsibilities stay in Python.
- Keep the App Router route shape aligned with design section 2.1.
