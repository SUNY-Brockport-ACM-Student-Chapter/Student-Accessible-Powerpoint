# SOP: Secret Rotation & Management

**In scope:** `GOOGLE_API_KEY` (the only runtime secret we have today), TLS certificates (covered in [`SOP_INCIDENT.md`](SOP_INCIDENT.md) §8), and future secrets.

**Current storage:** `.env` file on the production VM at `/home/mattarama443/Student-Accessible-Powerpoint/.env`. Plaintext, 0600, owned by `mattarama443`. No Vault, no Secret Manager. This is a known weakness — see §5.

---

## 1. Rotate `GOOGLE_API_KEY`

### 1.1 Create the new key

- Console → Google AI Studio / Vertex AI → create a new API key scoped to the same project/model (`gemini-2.0-flash-lite`).
- Label it with the date and operator, e.g. `sigai-prod-2026-04-22-dal`.

Do **not** disable the old key yet. You need both live briefly for a zero-downtime swap.

### 1.2 Deploy the new key

```bash
gcloud compute ssh instance-20250905-023343-pub --zone=us-central1-c
APP=/home/mattarama443/Student-Accessible-Powerpoint
sudo -u mattarama443 cp $APP/.env $APP/.env.bak-$(date +%Y%m%d-%H%M)
sudo -u mattarama443 nano $APP/.env        # replace GOOGLE_API_KEY value
sudo systemctl restart student-access-ppt
```

Wait 15 seconds. Then:

```bash
python scripts/smoke_test.py --url https://access.brockportsigai.org/accessibility --strict
```

Open the site, upload a small test deck, and confirm alt text is generated for at least one image (proves the Gemini call path).

### 1.3 Disable the old key

Only **after** the smoke test passes and you have seen a successful Gemini response in the logs:

- Console → revoke the old key.
- Delete the `.env.bak-*` files on the VM (`sudo -u mattarama443 rm $APP/.env.bak-*`).

### 1.4 Record it

Append to `docs/ops/SECRET_ROTATION_LOG.md`:

```
YYYY-MM-DD  rotated=GOOGLE_API_KEY  operator=<handle>  reason=<scheduled|leak|staffing>
```

Never paste the key value in the log. The log is public-visible in git.

---

## 2. Add a new secret

If a feature requires a new secret (e.g. OAuth client secret, database password):

1. Add the **variable name** (not value) to `.env.example`.
2. Document it in [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) "Stable facts" section.
3. Reference it from code via `os.environ` (or `pydantic-settings`), never hard-code.
4. Deploy: append the variable to `/home/mattarama443/Student-Accessible-Powerpoint/.env` on the VM and `systemctl restart student-access-ppt`.
5. If the secret rotates routinely (e.g. >1× / year), add a section to this SOP.

Never:
- commit a real secret value to any branch,
- log a secret,
- bake a secret into a Docker image or systemd unit `Environment=` line.

---

## 3. Handling a suspected leak

If you believe a key has been exposed (accidentally committed, logged, screenshotted, shared in a ticket):

1. **Rotate the key immediately** — §1, but with zero grace period. Revoke the old key *before* confirming the smoke test if the exposure is public.
2. **Scan the git history** for the leaked value:

   ```bash
   git log --all -p -S '<first-few-chars-of-key>' | head -100
   ```

   If found in git, the repo is compromised. A rotation is sufficient remediation *only* if the old key is revoked; do **not** try to rewrite history on the shared remote without team agreement.
3. **File an incident** — [`SOP_INCIDENT.md`](SOP_INCIDENT.md) §9 — with root cause = "secret leak".
4. **Search for further exposure**: logs (`journalctl --since "2 weeks ago" | grep -i '<key-pattern>'`), screenshots, Slack/Discord, PRs.

---

## 4. TLS private keys

Managed by certbot at `/etc/letsencrypt/live/access.brockportsigai.org/`. Rotate automatically via `certbot.timer`. Manual intervention steps are in [`SOP_INCIDENT.md`](SOP_INCIDENT.md) §8.

**Never back up the private key to anywhere outside the VM.**

---

## 5. Known gaps (planned improvements)

These are tracked as tech debt; see `docs/guardrails/INVARIANTS.md` §10.

- **Plaintext `.env`**: migrate to GCP Secret Manager + a small runtime loader. Proposal: a `REFACTOR.md` titled "Move secrets to GCP Secret Manager" — this has not been filed yet.
- **No leaked-secret scanner in CI**: add `detect-secrets` or `trufflehog` to `.github/workflows/` once CI exists.
- **No audit log**: secret rotations are only captured by `SECRET_ROTATION_LOG.md` (this SOP). A GCP IAM audit trail covers only the console-side actions.

If you are the agent picking up this work: start with Secret Manager migration; it closes the highest-severity gap.
