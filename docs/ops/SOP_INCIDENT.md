# SOP: Production Incident Response

**Scope:** any situation where `https://access.brockportsigai.org/accessibility` is unavailable, returning errors, or producing clearly wrong output.

**Target instance:** `instance-20250905-023343-pub` / `us-central1-c`.

---

## 1. First 60 seconds

Pick one:

- [ ] Can you open `https://access.brockportsigai.org/accessibility` in a browser? → no / yes?
- [ ] Does `python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility` pass? → which check failed?

Based on that, jump to the matching section below.

**If you cannot answer either question (network blocked, etc.)**, skip to §7 (escalation) — do not guess at fixes blind.

---

## 2. Symptom matrix

| Symptom | First probe | Section |
|---|---|---|
| Site does not respond at all (connection refused / timeout) | `gcloud compute instances describe instance-20250905-023343-pub --zone=us-central1-c --format='value(status)'` | §3 |
| Site responds but shows `502 Bad Gateway` / `503` | Streamlit is down. Check systemd. | §4 |
| Site responds but shows Streamlit error page ("Oh no.") | App crashed. Check journalctl. | §5 |
| Site renders but Gemini calls error out | API key / quota | §6 |
| Site renders but ChromaDB calls error out | chromadbd / chromadbapi | §5 |
| TLS cert warning | certbot renewal failed | §8 |

---

## 3. VM down

```bash
gcloud compute instances describe instance-20250905-023343-pub --zone=us-central1-c --format='value(status)'
```

- If `TERMINATED`: `gcloud compute instances start instance-20250905-023343-pub --zone=us-central1-c`, then wait 60 s and retry smoke test.
- If `STOPPING`: wait; retry in 60 s.
- If `RUNNING` but unreachable: GCP networking issue. Check `gcloud compute instances get-serial-port-output ...` and escalate (§7).

On startup, systemd brings up `chromadbd → chromadbapi → student-access-ppt` automatically. If any unit failed, §4 applies.

---

## 4. Service crashed / unit failed

SSH in and check:

```bash
sudo systemctl status chromadbd chromadbapi student-access-ppt --no-pager
```

For any unit in `failed` state:

```bash
sudo journalctl -u <unit-name> -n 200 --no-pager
sudo systemctl restart <unit-name>
sudo systemctl status <unit-name> --no-pager
```

Order of restart if multiple failed: `chromadbd` → `chromadbapi` → `student-access-ppt`.

**If `chromadbapi.service` keeps failing with "no such file":** it is probably the `Student-Accessibility-Powerpoint` typo in `start_scripts/chromadbapi.sh`. Confirm:

```bash
sudo cat /home/mattarama443/Student-Accessible-Powerpoint/start_scripts/chromadbapi.sh
```

Fix the typo (`Accessibility` → `Accessible`). Full detail: [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §3.3.

---

## 5. App logs point at a code bug

```bash
sudo journalctl -u student-access-ppt -n 500 --no-pager | less
```

Look for `Traceback`. Common signatures:

| Traceback contains | Likely cause | Action |
|---|---|---|
| `KeyError: 'order_number'` | Parsing regression broke invariant #1 | Rollback ([`SOP_ROLLBACK.md`](SOP_ROLLBACK.md)) |
| `PIL.UnidentifiedImageError` | WMF/EMF not normalized | Check `app/pptx_rag_quizzer/utils.py` image conversion |
| `google.api_core.exceptions.ResourceExhausted` | Gemini quota | §6 |
| `chromadb.errors.*` | Chroma down or collection corrupted | §4 + [`SOP_ROLLBACK.md`](SOP_ROLLBACK.md) §5 |
| `streamlit.runtime.scriptrunner.script_runner.RerunException` | Not a real error, ignore | — |
| `AttributeError: 'Shape' object has no attribute '...'` | python-pptx version mismatch | `pip install -r requirements.txt` |

If the last deploy is the likely cause → [`SOP_ROLLBACK.md`](SOP_ROLLBACK.md).

---

## 6. Gemini issues

Check the last error:

```bash
sudo journalctl -u student-access-ppt --since "1 hour ago" | grep -iE "gemini|google.api_core|GOOGLE_API_KEY|ResourceExhausted|PermissionDenied"
```

- **`ResourceExhausted` (429)**: Daily quota hit. The app already sleeps 60 s on this — no action unless the whole day is exhausted. If so, it will auto-recover at the quota reset; alternatively, rotate to a higher-tier key following [`SOP_SECRETS.md`](SOP_SECRETS.md).
- **`PermissionDenied` (401/403)**: Key revoked or missing. Rotate — [`SOP_SECRETS.md`](SOP_SECRETS.md).
- **`InvalidArgument` on model name**: Someone changed `gemini-2.0-flash-lite` → unknown model. Rollback.

---

## 7. Escalate

If §3–6 do not resolve it in **15 minutes**:

1. Post the current symptom, the last 50 lines of `journalctl -u student-access-ppt`, and the `gcloud instances describe` output to the team channel.
2. Mention the course owner if a live class is in session.
3. If DNS is broken, you can temporarily redirect users to the old instance (`instance-20250610-144049`, `35.196.195.118`) — but it is HTTP-only and has no systemd, so it needs manual start (`nohup venv/bin/streamlit run ...`). See [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) §4.

---

## 8. TLS issues

```bash
sudo certbot certificates
sudo journalctl -u certbot.service --since "7 days ago"
```

Manual renewal:

```bash
sudo certbot renew --nginx
sudo systemctl reload nginx
```

If certbot can't reach Let's Encrypt, check that port 80 is still open in the GCP firewall (`default-allow-http`) — ACME HTTP-01 uses it.

---

## 9. Close out

Every incident that lasted more than a brief blip must produce a short post-incident note in `docs/ops/INCIDENT_LOG.md`:

```
YYYY-MM-DD  duration=<min>  owner=<handle>
Symptom:
Root cause:
Fix:
Prevention action (ticket link):
```

If the prevention action is a code or infra change, file it as a [`BUG.md`](../templates/BUG.md) or [`REFACTOR.md`](../templates/REFACTOR.md) immediately.
