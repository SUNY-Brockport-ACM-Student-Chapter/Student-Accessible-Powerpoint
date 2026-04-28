# Deployment Scaffold

This directory will become the configuration of record for the Python processing service and ChromaDB deployment on the GCP VM.

Planned files from `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`:

- `docker-compose.yml`
- `Dockerfile.api`
- `Caddyfile`
- `env/.env.example`
- `systemd/docker-compose@sap.service`

The scaffold PR creates the directory and environment template only. Compose, Caddy, and systemd files should land in a follow-up deployment PR after the service layout is implemented.
