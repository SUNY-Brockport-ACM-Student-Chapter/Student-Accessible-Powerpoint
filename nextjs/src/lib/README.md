# Next.js Server Libraries

Modules from the migration design:

- `db.ts` — Prisma singleton.
- `supabase.ts` — Supabase server/client helpers.
- `processor.ts` — HTTP client for the Python processing service.
- `storage.ts` — Supabase Storage helpers.

Keep all secret-bearing helpers server-only.
