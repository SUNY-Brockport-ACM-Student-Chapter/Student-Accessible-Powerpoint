# Supabase Provisioning Notes

The Next.js migration expects two private Storage buckets:

- `pptx-uploads` — source PowerPoint decks uploaded by signed URL.
- `pptx-outputs` — rebuilt accessible PowerPoint decks.

The uploads bucket must enforce the same v1 limit as the application:

```sql
update storage.buckets
set file_size_limit = 52428800
where id = 'pptx-uploads';
```

The client-provided upload size is only a fast-fail UX check. Supabase Storage
must enforce the 50 MB ceiling so signed upload URLs cannot be abused with
oversized files.
