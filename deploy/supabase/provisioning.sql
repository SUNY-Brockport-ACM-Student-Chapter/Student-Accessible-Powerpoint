-- Supabase provisioning SQL for the Next.js migration.
-- Run after the Prisma schema has been applied to the Supabase Postgres database.

alter table public."Profile" enable row level security;

drop policy if exists "Users can read their own profile" on public."Profile";
create policy "Users can read their own profile"
on public."Profile"
for select
to authenticated
using (id = auth.uid()::text);

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'pptx-uploads',
  'pptx-uploads',
  false,
  52428800,
  array['application/vnd.openxmlformats-officedocument.presentationml.presentation']
)
on conflict (id) do update
set
  public = excluded.public,
  file_size_limit = excluded.file_size_limit,
  allowed_mime_types = excluded.allowed_mime_types;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'pptx-outputs',
  'pptx-outputs',
  false,
  52428800,
  array['application/vnd.openxmlformats-officedocument.presentationml.presentation']
)
on conflict (id) do update
set
  public = excluded.public,
  file_size_limit = excluded.file_size_limit,
  allowed_mime_types = excluded.allowed_mime_types;
