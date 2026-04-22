# SOP: Deploy to Production

**Applies to:** `instance-20250905-023343-pub` (zone `us-central1-c`), deploying branch `Aggrement` to `/home/mattarama443/Student-Accessible-Powerpoint`.

**Audience:** any agent or developer pushing a code change live. Read [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) once before your first deploy.

**Do not skip steps.** This SOP has specific guards against the drift we have already seen on this VM.

---

## 0. Preconditions

- [ ] Your change is merged to `main` and cherry-picked (or merged) into `Aggrement` with a passing PR review.
- [ ] `python scripts/preflight.py` passes locally on the commit you intend to deploy.
- [ ] You have `gcloud` auth with `compute.instances.setMetadata` on the prod instance.
- [ ] It is **not** during a live class / demo window (check with the course owner).
- [ ] You have a rollback plan (git SHA to roll back to + link to [`SOP_ROLLBACK.md`](SOP_ROLLBACK.md) open in another tab).

If any of the above is not true, STOP.

---

## 1. Prep: capture the current prod state

```bash
gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c
```

Inside the VM:

```bash
APP=/home/mattarama443/Student-Accessible-Powerpoint
sudo -u mattarama443 -H bash -lc '
  cd '"$APP"'
  echo "=== HEAD before deploy ===";    git -C '"$APP"' rev-parse --short HEAD
  echo "=== Local modifications ===";   git -C '"$APP"' status -s
  echo "=== Last 3 commits ===";        git -C '"$APP"' log --oneline -3
'
```

**Record the "HEAD before deploy" SHA.** This is your rollback target. Paste it into your deploy issue/PR comment.

If `git status -s` shows modifications, they must be resolved *before* you pull. Historically prod carries uncommitted local edits on `app/ppt_notes.py` and `start_app.py` — see [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §3.4. Capture the diff first:

```bash
sudo -u mattarama443 git -C $APP diff > ~/predeploy-drift-$(date +%Y%m%d-%H%M).patch
sudo -u mattarama443 git -C $APP stash push -u -m "predeploy-$(date +%Y%m%d-%H%M)"
```

**Do not throw away the patch file.** If any hunk is production-only (server-specific flags, hotfix not yet upstreamed), treat it as a [`../templates/BUG.md`](../templates/BUG.md) to be upstreamed after the deploy.

---

## 2. Pre-deploy smoke test (current state)

From your workstation, before changing anything:

```bash
python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility
```

If the pre-deploy smoke test already fails, you are walking into a broken prod. Stop and follow [`SOP_INCIDENT.md`](SOP_INCIDENT.md) first.

---

## 3. Pull and install

On the VM, as root / with sudo:

```bash
APP=/home/mattarama443/Student-Accessible-Powerpoint
sudo -u mattarama443 -H bash -lc '
  cd '"$APP"'
  git fetch origin
  git checkout Aggrement
  git pull --ff-only origin Aggrement
  source venv/bin/activate
  pip install -r requirements.txt
'
```

`--ff-only` is mandatory. If it refuses, the tree is dirty or someone committed on the VM — go back to §1.

**If `requirements.txt` did not change**, you can skip `pip install` — but it is cheap and safe to run.

---

## 4. Re-apply production-only patches (if any were stashed)

Only if §1 produced a `.patch` file that contains production-only edits (not yet upstreamed):

```bash
sudo -u mattarama443 git -C $APP apply --index ~/predeploy-drift-YYYYMMDD-HHMM.patch
```

Resolve any conflicts manually. File an issue tagged `drift-upstream` to move those edits into git.

---

## 5. Restart services (order matters)

```bash
sudo systemctl restart chromadbd              # 1. vector DB
sleep 3
sudo systemctl restart chromadbapi            # 2. FastAPI wrapper
sleep 2
sudo systemctl restart student-access-ppt    # 3. Streamlit
```

Watch boot logs for the Streamlit service:

```bash
sudo journalctl -u student-access-ppt -n 50 -f
# Ctrl-C when you see "You can now view your Streamlit app..."
```

---

## 6. Post-deploy smoke test (required gate)

From your workstation:

```bash
python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility --strict
```

Expected: exit code 0 with all checks green. If any check fails, **immediately** follow [`SOP_ROLLBACK.md`](SOP_ROLLBACK.md).

Additionally, open the site in a browser and verify the four-stage UI renders: upload → describe images → final processing → download.

---

## 7. Record the deploy

Append a line to `docs/ops/DEPLOY_LOG.md` (commit on `main`):

```
YYYY-MM-DD HH:MM UTC  <your-handle>  <old-sha> → <new-sha>  (prod Aggrement)  notes: <one line>
```

If you had to re-apply a drift patch in §4, link the patch filename here as well.

---

## 8. When *not* to use this SOP

- **Docker deploy?** There is no Docker deploy for this app in prod. The `Prod-v1` branch's `docker-compose.yml` is not the deployed topology. Do not follow this SOP to apply it; that is a migration project, not a deploy.
- **Static content only?** If you only changed `/var/www/streamlit-home/index.html`, the landing page, you can edit it in place and `sudo systemctl reload nginx`. But add the file to git at the same time — landing-page drift is a recurring problem.
- **Secret rotation?** Use [`SOP_SECRETS.md`](SOP_SECRETS.md) instead.

---

## 9. Failure modes you will actually hit

| Symptom | Cause | Fix |
|---|---|---|
| `git pull` refuses with "local changes" | Someone edited on the VM | §1 stash → re-apply |
| `pip install` fails on `google-generativeai` | Python ABI mismatch after OS update | `rm -rf venv && python3.11 -m venv venv && source venv/bin/activate && pip install -r requirements.txt` |
| Streamlit says `Error: port 8501 already in use` | Previous process did not die | `sudo systemctl stop student-access-ppt && sudo pkill -f streamlit && sudo systemctl start student-access-ppt` |
| `curl localhost:8001/health` returns 404 | `chromadbapi.sh` typo bit you (the `cd` fails, and something else changed `WorkingDirectory`) | Fix the script: `cd /home/mattarama443/Student-Accessible-Powerpoint` (not `Accessibility`) |
| 502 from nginx after restart | Streamlit still booting; give it 15 s | Re-run the smoke test |
| Gemini quota exhausted during smoke | Daily quota hit by the test itself | Acceptable — the smoke test is `--lite` by default and does not call Gemini |

See also [`SOP_INCIDENT.md`](SOP_INCIDENT.md).
