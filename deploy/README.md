# Deployment Scaffold

This directory will become the configuration of record for the Python processing service and ChromaDB deployment on the GCP VM.

Files from `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`:

- `docker-compose.yml`
- `Dockerfile.api`
- `Caddyfile`
- `env/.env.example`
- `systemd/docker-compose@sap.service`

Copy `env/.env.example` to `env/.env` on the VM, fill it from the secured deployment secret store, and keep `env/.env` out of git.

The stack exposes only Caddy on ports 80/443. Chroma and the Python API stay inside the Docker network.
