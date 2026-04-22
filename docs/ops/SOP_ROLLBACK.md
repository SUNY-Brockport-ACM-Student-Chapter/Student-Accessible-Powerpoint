# SOP: Rollback a Production Deploy

**When to use:** immediately after a deploy if the smoke test fails, or at any time prod is broken and the last deploy is the likely cause.

**Target instance:** `instance-20250905-023343-pub` / `us-central1-c` / branch `Aggrement`.

**Required:** the "HEAD before deploy" SHA you recorded in [`SOP_DEPLOY.md`](SOP_DEPLOY.md) §1. If you did not record it, get it from `git reflog`.

---

## 1. Decide: rollback or forward-fix?

| Situation | Action |
|---|---|
| Smoke test failed, no user has hit the site since the deploy | **Rollback.** Fast and cheap. |
| Users are on the site right now and the bug is cosmetic | **Forward-fix.** Open an incident ([`SOP_INCIDENT.md`](SOP_INCIDENT.md)) and push a hotfix. |
| Gemini API key issue | Not a code problem. See [`SOP_SECRETS.md`](SOP_SECRETS.md). |
| Prod is hard-down (502, process crash looping) | **Rollback first, diagnose later.** |

If in doubt: rollback. Downtime > serving wrong data.

---

## 2. Rollback (Python/git)

SSH in:

```bash
gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c
```

Revert the repo to the previous SHA:

```bash
APP=/home/mattarama443/Student-Accessible-Powerpoint
OLD_SHA=<paste-your-pre-deploy-sha>

sudo -u mattarama443 -H bash -lc '
  cd '"$APP"'
  git fetch origin
  git checkout Aggrement
  git reset --hard '"$OLD_SHA"'
  source venv/bin/activate
  pip install -r requirements.txt
'
```

Restart the three services in order:

```bash
sudo systemctl restart chromadbd
sleep 3
sudo systemctl restart chromadbapi
sleep 2
sudo systemctl restart student-access-ppt
```

---

## 3. Confirm rollback

From your workstation:

```bash
python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility --strict
```

Open the site and verify it renders. Attempt a trivial upload.

---

## 4. Rollback a dependency change

If the failing deploy changed `requirements.txt`, `git reset --hard` alone does not always restore the venv (pip rarely removes packages). The nuclear option is safe here:

```bash
sudo -u mattarama443 -H bash -lc '
  cd '"$APP"'
  rm -rf venv
  python3.11 -m venv venv
  source venv/bin/activate
  pip install -r requirements.txt
'
sudo systemctl restart chromadbd chromadbapi student-access-ppt
```

Takes ~2–3 minutes. Do this if the rollback leaves the app in a "ModuleNotFoundError" state.

---

## 5. Rollback a ChromaDB schema change

Chroma's vector data lives at `/home/mattarama443/Student-Accessible-Powerpoint/chroma/`. If the deploy silently broke the collection layout (rare — only on `chromadb` major bumps):

```bash
sudo systemctl stop chromadbd chromadbapi student-access-ppt
sudo -u mattarama443 mv $APP/chroma $APP/chroma.broken-$(date +%s)
sudo -u mattarama443 mkdir $APP/chroma
sudo systemctl start chromadbd chromadbapi student-access-ppt
```

This is equivalent to "fresh RAG index"; the app seeds it lazily on first upload. The app will be slower on cold start but functionally correct. **Do not delete the `.broken-*` directory** until the incident is closed — it is your forensic copy.

---

## 6. Rollback nginx / systemd unit changes

Unit files and nginx configs are **not in git** on production. If the deploy touched them:

```bash
# nginx
sudo cp /etc/nginx/sites-enabled/streamlit-https /etc/nginx/sites-enabled/streamlit-https.broken
sudo nano /etc/nginx/sites-enabled/streamlit-https    # manually revert
sudo nginx -t && sudo systemctl reload nginx

# systemd
sudo nano /etc/systemd/system/student-access-ppt.service
sudo systemctl daemon-reload
sudo systemctl restart student-access-ppt
```

**After the incident closes:** open a [`REFACTOR.md`](../templates/REFACTOR.md) to track `/etc/nginx/sites-enabled/*` and `/etc/systemd/system/*.service` in git (see Invariant #11 in [`../guardrails/INVARIANTS.md`](../guardrails/INVARIANTS.md)).

---

## 7. Record the rollback

Append to `docs/ops/DEPLOY_LOG.md`:

```
YYYY-MM-DD HH:MM UTC  <your-handle>  ROLLBACK  <bad-sha> → <old-sha>  reason: <one line>
```

Then open an issue using [`../templates/BUG.md`](../templates/BUG.md) so the faulty change can be fixed forward cleanly.
