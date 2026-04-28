# Deployment Provisioning Runbook

This runbook provisions the split Next.js + Python deployment described in
`docs/refactor/NEXTJS_MIGRATION_DESIGN.md`.

## 1. Supabase

1. Confirm the Supabase project is the intended production project.
2. From `nextjs/`, set `DIRECT_URL` to the direct Supabase Postgres connection
   string and apply the Prisma schema:

   ```sh
   npm ci
   npm run prisma:generate
   npx prisma db push
   ```

3. In the Supabase SQL editor, run `deploy/supabase/provisioning.sql`.
4. Confirm Storage has private buckets:
   - `pptx-uploads`
   - `pptx-outputs`
5. Confirm `pptx-uploads` has `file_size_limit = 52428800`.
6. Configure Auth redirect URLs:
   - local: `http://localhost:3000/auth/callback`
   - production: `https://<vercel-app-host>/auth/callback`

## 2. Vercel

Deploy `nextjs/` as the Vercel project root.

Set these environment variables:

- `NEXT_PUBLIC_SUPABASE_URL`
- `NEXT_PUBLIC_SUPABASE_ANON_KEY`
- `DATABASE_URL`
- `DIRECT_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_UPLOADS_BUCKET=pptx-uploads`
- `SUPABASE_OUTPUTS_BUCKET=pptx-outputs`
- `PY_SERVICE_URL=https://<processor-api-host>`
- `PY_SERVICE_SHARED_SECRET`
- `NEXT_PUBLIC_APP_URL=https://<vercel-app-host>`

`PY_SERVICE_SHARED_SECRET` must exactly match the VM value.

## 3. GCP VM

Expected release layout:

```text
/opt/sap/releases/<release-id>
/opt/sap/current -> /opt/sap/releases/<release-id>
```

On the VM:

1. Copy the repo release to `/opt/sap/releases/<release-id>`.
2. Update `/opt/sap/current` to point at that release.
3. Copy `deploy/env/.env.example` to `deploy/env/.env` and fill:
   - `PY_SERVICE_SHARED_SECRET`
   - `NEXT_PUBLIC_APP_URL`
   - `GOOGLE_API_KEY`
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SUPABASE_UPLOADS_BUCKET=pptx-uploads`
   - `SUPABASE_OUTPUTS_BUCKET=pptx-outputs`
4. Copy `deploy/env/caddy.env.example` to `deploy/env/caddy.env` and fill:
   - `ACME_EMAIL`
   - `PROCESSOR_PUBLIC_HOST`
5. Lock down env files:

   ```sh
   chmod 600 /opt/sap/current/deploy/env/.env
   chmod 600 /opt/sap/current/deploy/env/caddy.env
   ```

6. Validate Compose:

   ```sh
   cd /opt/sap/current/deploy
   docker compose config
   ```

7. Install and start the unit:

   ```sh
   sudo cp /opt/sap/current/deploy/systemd/docker-compose@sap.service /etc/systemd/system/
   sudo systemctl daemon-reload
   sudo systemctl enable --now docker-compose@sap.service
   ```

## 4. Network Cutover

Before production traffic:

- Point the processor API DNS host at the VM.
- Keep only ports 80 and 443 open publicly on the VM.
- Delete or disable legacy public rules for Streamlit/Chroma ports (`8501`,
  `8001`, and raw Chroma `8000`).
- Confirm `https://<processor-api-host>/health` returns healthy.

## 5. Smoke Test

1. Visit the Vercel app.
2. Sign in by magic link.
3. Accept consent.
4. Upload a small `.pptx`.
5. Confirm the job reaches review.
6. Confirm/export descriptions.
7. Download the rebuilt accessible `.pptx`.
