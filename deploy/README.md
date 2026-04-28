# Deployment Scaffold

This directory will become the configuration of record for the Python processing service and ChromaDB deployment on the GCP VM.

Files from `docs/refactor/NEXTJS_MIGRATION_DESIGN.md`:

- `docker-compose.yml`
- `Dockerfile.api`
- `Caddyfile`
- `env/.env.example`
- `systemd/docker-compose@sap.service`

Copy `env/.env.example` to `env/.env` and `env/caddy.env.example` to `env/caddy.env` on the VM. Fill both from the secured deployment secret store and keep the concrete `.env` files out of git.

The stack exposes only Caddy on ports 80/443. Chroma and the Python API stay inside the Docker network.
