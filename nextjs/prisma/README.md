# Prisma

`schema.prisma` defines the initial Supabase PostgreSQL data model for the migration.

Do not commit database passwords or Supabase service-role credentials.

## Required Supabase RLS Policy

The request proxy reads the signed-in user's `Profile.consentAcceptedAt` value with the Supabase anon key and the user's session cookie. The `Profile` table therefore needs an RLS policy equivalent to:

```sql
create policy "Users can read their own profile"
on "Profile"
for select
to authenticated
using (id = auth.uid()::text);
```

Without this policy, authenticated users may be redirected to `/consent` even after accepting consent because the profile lookup will return no row.
