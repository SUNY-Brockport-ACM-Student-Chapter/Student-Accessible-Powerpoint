# Production Environment

Authoritative map of the **live** Student-Accessible-Powerpoint deployment on Google Cloud Platform.

> **Scope:** GCP project `brockport-acm-sigai-project`. Accessed via `gcloud compute ssh` as `davelonski12@gmail.com`. Of the 6 running VMs, 3 were SSH-accessible; the remaining 3 return `compute.instances.setMetadata` permission denied and are not covered here.
>
> **Companion files:** [`PROJECT_OVERVIEW.md`](./PROJECT_OVERVIEW.md) (human narrative) · [`AGENT_CONTEXT.md`](./AGENT_CONTEXT.md) (agent reference).
>
> **Last verified:** Apr 22, 2026.

---

## 1. TL;DR

- **Production URL:** `https://access.brockportsigai.org/accessibility`
- **Production VM:** `instance-20250905-023343-pub` (zone `us-central1-c`, `e2-medium`, Debian 12 Bookworm, external IP `34.42.255.126`).
- **Deployed branch:** `Aggrement` — *not* `main` and *not* the Docker-based `Prod-v1` layout documented in the repo.
- **Deployment style:** **Bare-metal Python** on systemd, **not Docker Compose**. The `Prod-v1` branch's `docker-compose.yml` + Caddy stack is not used in production.
- **Process manager:** 3 systemd units under user `mattarama443` (Chroma daemon, Chroma REST wrapper, Streamlit app).
- **Reverse proxy / TLS:** nginx + Let's Encrypt (certbot), *not* Caddy.
- **Repo on prod:** `/home/mattarama443/Student-Accessible-Powerpoint` — currently has **uncommitted local modifications** to `app/ppt_notes.py` and `start_app.py` plus untracked `.streamlit/` and `start_scripts/` directories (architectural drift, see §6).
- **Second app on the same VM:** a separate project (`davidlonski/RAG-Dev`, branch `Prod-v1`) runs alongside at `/rag-application` on port 8502. It is **not** part of this repo.
- **Older deployment still running:** `instance-20250610-144049` (us-east1-c) carries an older install under user `davelonski12` with the same nginx `server_name access.brockportsigai.org` — but DNS does not point there anymore. It is effectively orphaned prod.
- **Non-project CI/CD box:** `instance-20251114-154547-main-dev` runs Jenkins (port 8080) + a Next.js dev server for an unrelated project (`edually-backend`). No Student-Accessible-Powerpoint pipeline lives there.

---

## 2. Instance Inventory

From `gcloud compute instances list` (project `brockport-acm-sigai-project`):

| Instance | Zone | Type | External IP | SSH | Role |
|---|---|---|---|---|---|
| `instance-20250905-023343-pub` | us-central1-c | e2-medium | 34.42.255.126 | ✅ | **Production** (Student-Accessible-Powerpoint + RAG-Dev) |
| `instance-20250610-144049` | us-east1-c | e2-standard-2 | 35.196.195.118 | ✅ | Older prod install (orphan, still running) |
| `instance-20251114-154547-main-dev` | us-central1-a | e2-medium | 35.192.148.196 | ✅ | Jenkins CI + unrelated Next.js dev |
| `instance-20250628-145549` | us-central1-f | e2-medium | 35.238.7.137 | ❌ denied | unknown |
| `instance-20250704-210400-brockportsigai` | us-east1-c | e2-medium | 34.75.70.146 | ❌ denied | unknown (name suggests Brockport-branded deployment) |
| `instance-20250628-145549-yub` | northamerica-northeast1-c | e2-medium | 35.203.22.246 | ❌ denied | unknown (static IP `static-ip-for-yub`) |

Static IPs reserved in the project:

| Reservation | Address | Region | Status |
|---|---|---|---|
| `arraywall-ip-adress` | 34.47.4.212 | northamerica-northeast1 | RESERVED (unused) |
| `static-ip-for-yub` | 35.203.22.246 | northamerica-northeast1 | IN_USE by `-yub` |

Firewall (VPC `default`):

| Rule | Direction | Allow | Notes |
|---|---|---|---|
| `default-allow-ssh` | INGRESS | tcp:22 | |
| `default-allow-http` | INGRESS | tcp:80 | prod nginx |
| `default-allow-https` | INGRESS | tcp:443 | prod nginx (TLS) |
| `allow-8501` | INGRESS | tcp:8501 | **Streamlit exposed directly to the internet** — bypasses nginx |
| `allow-8001` | INGRESS | tcp:8001 | **Chroma REST API exposed directly to the internet** |
| `allow-port-1000` | INGRESS | tcp:1000 | unused |
| `default-allow-internal` | INGRESS | all TCP/UDP | intra-VPC |
| `default-allow-icmp` | INGRESS | icmp | |
| `default-allow-rdp` | INGRESS | tcp:3389 | unused on Linux boxes |

Instance `-pub` carries the network tags `http-server`, `https-server`, `lb-health-check`.

---

## 3. Production Topology (`instance-20250905-023343-pub`)

### 3.1 Request path

```
Browser ──HTTPS──▶ nginx :443  (access.brockportsigai.org)
                      │
                      ├─ /                → static /var/www/streamlit-home/index.html  (landing)
                      ├─ /accessibility   → 127.0.0.1:8501  (Streamlit → Student-Accessible-Powerpoint)
                      └─ /rag-application → 127.0.0.1:8502  (Streamlit → RAG-Dev, SEPARATE PROJECT)

Streamlit :8501 (student-access-ppt.service, user mattarama443)
       │
       └─HTTP──▶ 127.0.0.1:8001  chroma-api (FastAPI wrapper, chromadbapi.service)
                          │
                          └─HTTP──▶ [::1]:8000  chromadb server (chromadbd.service, chroma run)
                                              └─ data: /home/mattarama443/Student-Accessible-Powerpoint/chroma/  (~15 MB)

Streamlit :8502 (ad-hoc nohup process, user mattarama443, cwd=/home/mattarama443/RAG-Dev)
       │
       └─HTTP──▶ 127.0.0.1:8003  RAG-Dev chroma-api
                          │
                          └─HTTP──▶ 127.0.0.1:8010  RAG-Dev chroma (port 8010)
```

### 3.2 nginx config — `/etc/nginx/sites-enabled/streamlit-https`

```nginx
server {
    if ($host = access.brockportsigai.org) {
        return 301 https://$host$request_uri;
    }
    listen 80;
    server_name access.brockportsigai.org;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name access.brockportsigai.org;
    ssl_certificate     /etc/letsencrypt/live/access.brockportsigai.org/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/access.brockportsigai.org/privkey.pem;
    include /etc/letsencrypt/options-ssl-nginx.conf;
    ssl_dhparam /etc/letsencrypt/ssl-dhparams.pem;

    client_max_body_size 500M;           # matches Streamlit maxUploadSize

    root /var/www/streamlit-home;
    index index.html;

    location / { try_files $uri $uri/ =404; }

    location /accessibility {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";     # required for Streamlit WebSocket
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For  $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }

    location /rag-application {
        proxy_pass http://127.0.0.1:8502;
        # (identical upgrade / forwarded headers)
    }
}
```

Certbot is wired via `certbot.timer` (next run ~13 h out, last run ~8 h ago). Certs renew automatically.

### 3.3 systemd units

All three services run as user **`mattarama443`** with `WorkingDirectory=/home/mattarama443/Student-Accessible-Powerpoint`. Configured for restart on failure.

| Unit | Script | Command |
|---|---|---|
| `chromadbd.service` | `start_scripts/chromadbd.sh` | `venv/bin/chroma run` (binds `[::1]:8000`) |
| `chromadbapi.service` | `start_scripts/chromadbapi.sh` | `python app/chroma-api/app.py` (binds `0.0.0.0:8001`) |
| `student-access-ppt.service` | `start_scripts/streamlit.sh` | `streamlit run app/ppt_notes.py` (binds `0.0.0.0:8501`) |

Unit file for reference (Streamlit):

```ini
[Unit]
Description=Streamlit Frontend for Student Accessible Powerpoint Services
After=network.target

[Service]
Type=simple
User=mattarama443
WorkingDirectory=/home/mattarama443/Student-Accessible-Powerpoint
ExecStart=/bin/bash /home/mattarama443/Student-Accessible-Powerpoint/start_scripts/streamlit.sh
Restart=on-failure
RestartSec=5
KillMode=process

[Install]
WantedBy=multi-user.target
```

Start scripts, verbatim:

```bash
# start_scripts/chromadbd.sh
cd /home/mattarama443/Student-Accessible-Powerpoint
source venv/bin/activate
venv/bin/chroma run

# start_scripts/chromadbapi.sh
cd /home/mattarama443/Student-Accessibility-Powerpoint     # ⚠ TYPO: directory does not exist
source venv/bin/activate
python app/chroma-api/app.py

# start_scripts/streamlit.sh
cd /home/mattarama443/Student-Accessible-Powerpoint
source venv/bin/activate
streamlit run app/ppt_notes.py
```

**⚠ Footgun.** The `chromadbapi.sh` `cd` fails silently (`Accessibility` vs `Accessible`). The FastAPI wrapper still runs because systemd's `WorkingDirectory=` directive anchors the process before bash runs; `python app/chroma-api/app.py` is resolved against the systemd working directory, not the failed `cd` target. If anyone removes or changes the `WorkingDirectory` line, the service will silently break. **Fix the script.**

Ports on pub observed via `ss -tlnp`:

| Port | Process | Bind |
|---|---|---|
| 22 | sshd | 0.0.0.0 |
| 80 | nginx (3 workers) | 0.0.0.0 |
| 443 | nginx | 0.0.0.0 |
| 8000 | chromadbd (`chroma run`) | `[::1]` (localhost-only) |
| 8001 | chromadbapi (`python app/chroma-api/app.py`) | 0.0.0.0 — **also exposed via firewall `allow-8001`** |
| 8010 | RAG-Dev chroma | 127.0.0.1 |
| 8003 | RAG-Dev chroma-api | 0.0.0.0 |
| 8501 | student-access-ppt (Streamlit) | 0.0.0.0 — **also exposed via firewall `allow-8501`** |
| 8502 | RAG-Dev Streamlit | 0.0.0.0 |
| 25 | exim4 (localhost only) | 127.0.0.1 |
| 20201 / 20202 | Google Cloud Ops Agent (metrics + logs) | — |

### 3.4 Code layout on prod

```
/home/mattarama443/Student-Accessible-Powerpoint/   (branch: Aggrement, remote: SUNY-Brockport-ACM-Student-Chapter/...)
├── .env                     # GOOGLE_API_KEY=<redacted>  (403 bytes)
├── .streamlit/              # untracked — only config.toml (baseUrlPath=accessibility)
├── app/                     # Aggrement branch layout (pptx.py, word.py, image.py, etc.)
│   ├── ppt_notes.py         # Streamlit entry (LOCALLY MODIFIED — drift from Aggrement HEAD)
│   ├── chroma-api/app.py    # FastAPI wrapper
│   ├── pptx_rag_quizzer/    # {utils, pptx, word, image, rag_core}.py
│   └── models/models.py
├── chroma/                  # ChromaDB persistent data (~15 MB, untracked)
├── consent_responses.csv    # IRB consent log — 6 rows so far (first entry 2026-04-16)
├── start_scripts/           # untracked — the three systemd start scripts
├── venv/                    # Python 3.11.2 virtualenv
├── requirements.txt / requirements-app.txt
├── start_app.py             # LOCALLY MODIFIED
└── RAG_INTEGRATION_README.md
```

Git state (`sudo -u mattarama443 git status -s`):
```
 M app/ppt_notes.py
 M start_app.py
?? .streamlit/
?? app/app.out
?? app/ppt_notes.py.bak
?? chroma/
?? start_scripts/
```
Branch `Aggrement` at commit `6a8175a` (one commit behind `origin/Aggrement`'s HEAD `ee8606d` observed in the repo analysis — so prod is slightly stale even from its own branch). The modifications to `app/ppt_notes.py` and `start_app.py` are **not in any upstream branch**: there is untracked production-only configuration living on the VM.

Virtualenv key pins (Python 3.11.2):
```
streamlit, python-pptx, Pillow, google-generativeai==0.8.6, chromadb==1.5.7, fastapi==0.135.3,
pydantic, requests, pytesseract
```

### 3.5 Secrets

- `/home/mattarama443/Student-Accessible-Powerpoint/.env` holds the live `GOOGLE_API_KEY`. Not referenced by any external secret manager; lives on disk only.
- `/opt/vault/` exists on `main-dev` but is not wired into this production VM. There is no HashiCorp Vault, GCP Secret Manager, or other secret injection on prod.
- TLS cert + key live in `/etc/letsencrypt/live/access.brockportsigai.org/`, rotated by the `certbot.timer`.

### 3.6 User accounts on `pub`

`/home/` contents: `davel`, `davelonski12`, `mattarama443`, `yubrajkhatri977`, `zgsdwhwd`. The running application stack is wholly owned by `mattarama443`. Other homes exist for operators with SSH access.

### 3.7 Second app sharing the VM (`RAG-Dev`)

`/home/mattarama443/RAG-Dev/` — repository `https://github.com/davidlonski/RAG-Dev.git`, branch `Prod-v1`. Launched **without** systemd (ad-hoc `nohup` via `sudo -u mattarama443`). Listens on 8502 behind `location /rag-application` in nginx. Uses its own virtualenv, its own ChromaDB instance (ports 8010 + 8003), its own `chroma-data/` directory. **Not part of this repo.** Mentioned here only because any ops work on pub touches it.

---

## 4. Older production install — `instance-20250610-144049`

Still running, but superseded. It has the same `server_name access.brockportsigai.org` in its nginx config, yet DNS now resolves to `pub` (34.42.255.126), so this box is effectively dark for public traffic — although port 8501 is still listed in the firewall, making it reachable directly by IP.

| Fact | Value |
|---|---|
| Zone | us-east1-c |
| External IP | 35.196.195.118 |
| App user | `davelonski12` |
| App dir | `/home/davelonski12/Student-Accessible-Powerpoint` (~2.7 GB) |
| Process manager | **None** — raw `nohup` (e.g. `nohup venv/bin/python3 venv/bin/streamlit run app/ppt_notes.py --server.port 8501 --server.headless true > /tmp/streamlit.log 2>&1 &`) |
| Services visible | only `nginx.service`; chroma + streamlit are foreground `nohup` processes |
| nginx | HTTP-only, proxies `/` → `127.0.0.1:8501`; has `/.well-known/acme-challenge/` alias but no 443 listener |
| Chroma | `chroma run --path chroma-db` on `[::1]:8000` (same-venv pattern, same as pub) |
| chroma-api | `python app.py` on `0.0.0.0:8001` |

This box is the predecessor of pub. It confirms the historical deployment pattern (systemd-less, one-user, same dir layout). **Recommended action:** either shut it down or confirm retention is intentional — see §7.

---

## 5. Jenkins + Next.js dev box — `instance-20251114-154547-main-dev`

Not part of this product, included for completeness since it was SSH-accessible.

- Debian 12, up 158 days.
- **Jenkins** at `http://35.192.148.196:8080` (service `jenkins.service`, running since 2026-02-16). No jobs or workspaces configured (`/var/lib/jenkins/workspace/` does not exist). Jenkins is installed but unused for this project.
- **Next.js dev server** on port 3000: `next-server (v16.0.3)` run by user `legonate` from `/home/legonate/edually-backend/` — a separate project unrelated to Student-Accessible-Powerpoint.
- `/opt/vault/` present but no systemd unit, not actively serving.

There is **no CI/CD pipeline** for Student-Accessible-Powerpoint anywhere in the project. Deployment is manual.

---

## 6. Codebase ↔ Live — Deltas

The repository's `README`, `Prod-v1` branch, and `DEPLOYMENT.md` document one deployment model; production uses another. Agents working on deployments must understand these differences.

| Aspect | Repo says | Live actually is |
|---|---|---|
| Packaging | `Prod-v1` branch: Docker Compose (chroma + chroma-api + web + caddy) | Bare-metal systemd units; no Docker on the VM (`docker` command not installed) |
| Reverse proxy | Caddy (ACME via Let's Encrypt) | nginx + certbot |
| Config toml | `app/.streamlit/config.toml` with `baseUrlPath=accessibility` | Repo's `.streamlit/` is **untracked on prod** — the deployed config was created on the VM, not pulled from the branch (it matches the Aggrement-branch config anyway) |
| Start scripts | `start_app.py` (interactive Python) | Three shell scripts in `start_scripts/`, invoked by systemd; `start_app.py` is locally modified and unused |
| Deployed branch | — | `Aggrement` (consent-wall branch), not `main`, not `Prod-v1` |
| Process isolation | Per-service container | Same Linux user (`mattarama443`), same virtualenv |
| Persistent Chroma volume | `./chroma-db:/chroma/chroma` container mount | `/home/mattarama443/Student-Accessible-Powerpoint/chroma/` directory in the repo tree |
| Secrets | `GOOGLE_API_KEY` via compose env | `.env` file in the repo directory on the VM |
| Path-based hosting | `baseUrlPath = accessibility` + Caddy reverse | `baseUrlPath = accessibility` + nginx reverse |
| Second Streamlit app (`/rag-application`) | **Not in repo** | Present — different project (`RAG-Dev`) running under the same user |
| Load balancer | not in repo | Instance tag `lb-health-check` suggests one was considered; no LB resources currently attach |
| CI/CD | not in repo | None; Jenkins exists on a separate box but has no jobs |

**Observed drift on the VM:**
- Uncommitted edits to `app/ppt_notes.py` and `start_app.py`.
- `chromadbapi.sh` contains a typo (`Accessibility` ≠ `Accessible`) that is masked by the systemd `WorkingDirectory`.
- Prod is at commit `6a8175a` of `Aggrement`; upstream is at `ee8606d`. Not pulled.

---

## 7. Operator playbook

### 7.1 How to SSH into prod
```bash
gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c
```

### 7.2 How to view logs
```bash
# Streamlit app
sudo journalctl -u student-access-ppt -f

# Chroma REST wrapper
sudo journalctl -u chromadbapi -f

# Chroma DB
sudo journalctl -u chromadbd -f

# nginx
sudo tail -f /var/log/nginx/error.log /var/log/nginx/access.log
```

### 7.3 Restart the app
```bash
sudo systemctl restart student-access-ppt     # Streamlit
sudo systemctl restart chromadbapi            # FastAPI wrapper
sudo systemctl restart chromadbd              # ChromaDB (nuking this drops the vector collection cache for in-flight jobs)
```
Start order on boot: `chromadbd` → `chromadbapi` → `student-access-ppt`. The `After=` clauses in the unit files are weak (`After=network.target` only), so if Chroma is slow to come up on reboot the API may race it; the API has `Restart=on-failure`, so it eventually recovers.

### 7.4 Deploy a code change (current manual process)
```bash
gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c
sudo -u mattarama443 -H bash -lc '
  cd /home/mattarama443/Student-Accessible-Powerpoint
  git stash                                   # preserve the production-only edits
  git pull origin Aggrement
  source venv/bin/activate
  pip install -r requirements.txt             # if deps changed
  git stash pop                               # reapply local edits
'
sudo systemctl restart chromadbapi student-access-ppt
```

> Before pulling, capture the diff of local modifications — some of them are production-only (e.g. server-specific Streamlit flags). Plan to upstream them instead of stashing forever.

### 7.5 Rotate the Gemini API key
```bash
sudo -u mattarama443 nano /home/mattarama443/Student-Accessible-Powerpoint/.env
sudo systemctl restart student-access-ppt
```

### 7.6 Renew TLS (automatic, but manual if ever needed)
```bash
sudo certbot renew --nginx
sudo systemctl reload nginx
```

### 7.7 Inspect consent records
```bash
sudo -u mattarama443 cat /home/mattarama443/Student-Accessible-Powerpoint/consent_responses.csv
```
CSV columns: `timestamp_utc,email,choice`. At last check, 6 rows (5 responses + header).

### 7.8 Scale / capacity notes
- `e2-medium` = 2 vCPU, 4 GB RAM. Gemini calls are the bottleneck, not CPU. Large decks (>30 slides) are fine.
- ChromaDB store at ~15 MB — no pressure.
- `client_max_body_size 500M` in nginx matches Streamlit's `maxUploadSize = 500` in the repo config.
- Single instance, no autoscaling, no LB. If it dies, `instance-20250610-144049` can be repurposed by repointing DNS, but it has no TLS listener and no systemd units — reviving it is a manual job.

---

## 8. Known issues / recommended cleanups

1. **Typo in `start_scripts/chromadbapi.sh`** — `cd .../Student-Accessibility-Powerpoint` should be `.../Student-Accessible-Powerpoint`. Currently masked by `WorkingDirectory=`. Fix.
2. **Production drift vs git** — `app/ppt_notes.py` and `start_app.py` carry uncommitted local changes. Capture them in a branch or stop editing on the VM.
3. **`start_scripts/` is not in git** on any branch. Add them to the repo (possibly alongside the existing `start_app.py`).
4. **Firewall rules `allow-8001` and `allow-8501` expose the internal services directly to the public internet**, bypassing nginx and TLS. The Streamlit UI is reachable as `http://34.42.255.126:8501` and the FastAPI wrapper as `http://34.42.255.126:8001`. Unless this is deliberate for debugging, tighten to `lb-health-check` tag or delete.
5. **Two instances advertise the same nginx `server_name` (`access.brockportsigai.org`).** `pub` serves it via HTTPS; `instance-20250610-144049` still answers on HTTP. Retire the old instance or remove its nginx config to avoid confusion.
6. **No monitoring/alerting pipeline** beyond the Google Cloud Ops Agent. No uptime check configured (that would have caught the old-instance orphaning).
7. **Single point of failure.** No LB, no autoscaler, no redundancy. An outage on `pub` takes the product down.
8. **Secrets on disk.** `.env` with the Gemini key is plaintext in the home directory. Consider GCP Secret Manager.
9. **No CI/CD** for this project. Jenkins on `main-dev` is idle. Deployments are manual SSH + git pull.
10. **`chroma` data directory is inside the checked-out repo** (`./chroma/`). A careless `git clean -fdx` would wipe all vector data. Move it outside the repo and point `chroma run --path …` at the new location.

---

## 9. Agent-oriented quick reference (dense)

> Read this block before touching production.

```yaml
gcp_project: brockport-acm-sigai-project
gcloud_account: davelonski12@gmail.com
default_zone: us-east1-d     # cli default, NOT where prod lives
prod:
  instance: instance-20250905-023343-pub
  zone: us-central1-c
  external_ip: 34.42.255.126
  dns: access.brockportsigai.org -> 34.42.255.126
  os: debian-12-bookworm
  arch: amd64
  machine_type: e2-medium
  network_tags: [http-server, https-server, lb-health-check]
  app_user: mattarama443
  app_dir: /home/mattarama443/Student-Accessible-Powerpoint
  deployed_branch: Aggrement
  deployed_commit: 6a8175a       # one commit behind origin/Aggrement HEAD ee8606d (as observed)
  secondary_app:
    repo: https://github.com/davidlonski/RAG-Dev.git
    branch: Prod-v1
    path: /home/mattarama443/RAG-Dev
    process: nohup (no systemd)
    nginx_path: /rag-application
    chroma_port: 8010
    chroma_api_port: 8003
    streamlit_port: 8502
  services_systemd:
    chromadbd.service:
      user: mattarama443
      exec: /bin/bash /home/mattarama443/Student-Accessible-Powerpoint/start_scripts/chromadbd.sh
      listens: "[::1]:8000"
      cmd: "venv/bin/chroma run"
    chromadbapi.service:
      exec: /bin/bash /home/mattarama443/Student-Accessible-Powerpoint/start_scripts/chromadbapi.sh
      listens: "0.0.0.0:8001"
      cmd: "python app/chroma-api/app.py"
      bug: "script cd's to nonexistent 'Student-Accessibility-Powerpoint'; saved by systemd WorkingDirectory"
    student-access-ppt.service:
      exec: /bin/bash /home/mattarama443/Student-Accessible-Powerpoint/start_scripts/streamlit.sh
      listens: "0.0.0.0:8501"
      cmd: "streamlit run app/ppt_notes.py"
  nginx:
    config: /etc/nginx/sites-enabled/streamlit-https
    tls_cert: /etc/letsencrypt/live/access.brockportsigai.org/
    renewer: certbot.timer (daily ~22:00 UTC)
    client_max_body_size: 500M
    routes:
      /:                  "static /var/www/streamlit-home/index.html"
      /accessibility:     "127.0.0.1:8501  (Streamlit, Upgrade headers for WS)"
      /rag-application:   "127.0.0.1:8502"
  runtime:
    python: 3.11.2
    venv: /home/mattarama443/Student-Accessible-Powerpoint/venv
    key_pkgs: {chromadb: 1.5.7, fastapi: 0.135.3, google-generativeai: 0.8.6}
  secrets:
    env_file: /home/mattarama443/Student-Accessible-Powerpoint/.env
    contains: [GOOGLE_API_KEY]
    storage: plain file on disk; no Vault/Secret Manager
  persistence:
    chroma_data: /home/mattarama443/Student-Accessible-Powerpoint/chroma   # ~15 MB
    consent_log: /home/mattarama443/Student-Accessible-Powerpoint/consent_responses.csv
  firewall:
    public: [22, 80, 443, 8001, 8501]
    internal_only: [8000, 8002/unused, 8010 (RAG-Dev chroma)]
  drift_from_repo:
    - "Bare-metal systemd instead of Prod-v1 docker-compose"
    - "nginx + certbot instead of Caddy"
    - "Uncommitted edits to app/ppt_notes.py, start_app.py"
    - "start_scripts/ and .streamlit/ not tracked in any branch"
    - "Prod is 1 commit behind origin/Aggrement"

old_prod:
  instance: instance-20250610-144049
  zone: us-east1-c
  external_ip: 35.196.195.118
  status: "orphan — still running, no DNS, http-only nginx"
  app_user: davelonski12
  app_dir: /home/davelonski12/Student-Accessible-Powerpoint
  process_mgr: nohup (no systemd)
  nginx: "server_name access.brockportsigai.org → 127.0.0.1:8501 (port 80 only)"

ci_dev_box:
  instance: instance-20251114-154547-main-dev
  zone: us-central1-a
  external_ip: 35.192.148.196
  relevant_to_this_project: false
  services: [jenkins (:8080, no jobs configured), next-server (legonate/edually-backend, :3000), vault installed but inactive]

inaccessible_instances:
  - name: instance-20250628-145549
    zone: us-central1-f
    deny_reason: "compute.instances.setMetadata permission"
  - name: instance-20250704-210400-brockportsigai
    zone: us-east1-c
    deny_reason: "compute.instances.setMetadata permission"
    note: "name suggests Brockport-branded deployment; warrants follow-up with an account holding permissions"
  - name: instance-20250628-145549-yub
    zone: northamerica-northeast1-c
    deny_reason: "compute.instances.setMetadata permission"
    external_ip_reserved_as: static-ip-for-yub

static_ips:
  arraywall-ip-adress: {ip: 34.47.4.212, region: northamerica-northeast1, status: RESERVED}
  static-ip-for-yub:   {ip: 35.203.22.246, region: northamerica-northeast1, status: IN_USE}

deploy_procedure:
  - "gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c"
  - "sudo -u mattarama443 git -C /home/mattarama443/Student-Accessible-Powerpoint pull origin Aggrement"
  - "(optional) source venv/bin/activate && pip install -r requirements.txt"
  - "sudo systemctl restart chromadbapi student-access-ppt"

incident_playbook:
  app_hung:
    - "sudo systemctl restart student-access-ppt"
    - "sudo journalctl -u student-access-ppt -n 200"
  chroma_hung:
    - "sudo systemctl restart chromadbd chromadbapi"
    - "verify: curl -sf http://localhost:8001/health"
  tls_expired:
    - "sudo certbot renew --nginx && sudo systemctl reload nginx"
  disk_full:
    - "check /home/mattarama443/Student-Accessible-Powerpoint/chroma and /tmp/*.log"
  gemini_quota:
    - "app already sleeps 60s on 'Resource has been exhausted'; no action unless sustained"
```

---

*Documentation generated from live reconnaissance on Apr 22, 2026. Inaccessible instances should be revisited when permissions are granted — in particular `instance-20250704-210400-brockportsigai`, whose name strongly suggests it hosts a Brockport-branded deployment.*
