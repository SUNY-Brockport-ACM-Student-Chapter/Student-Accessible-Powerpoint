# Next.js Server Libraries

Planned modules from the migration design:

- `db.ts` — Prisma singleton.
- `supabase.ts` — Supabase server/client helpers.
- `processor.ts` — HTTP client for the Python processing service.
- `storage.ts` — Supabase Storage helpers.

These are intentionally not implemented in the scaffold PR. Follow-up PRs should add them with focused tests and no client-side secret exposure.
