# Update Template (dependency / version bumps, config changes)

> **Purpose.** Agent-optimized template for changes that alter versions, dependencies, configuration, or minor behavior of an existing capability — without adding a new user-visible feature. Copy into the issue/PR and delete this block.

Common uses: Python dep bump, Gemini model change, `client_max_body_size` tweak, Chroma version bump, Streamlit version bump, firewall rule edit.

---

## 1. TL;DR

One sentence: what's being updated and why.

## 2. What changes

- **Package / model / config:** e.g. `chromadb 1.5.7 → 1.6.0`
- **Scope:** `requirements.txt`, `requirements-app.txt`, `.streamlit/config.toml`, firewall, systemd, nginx — pick the files.
- **Expected behavior change (if any):** …
- **Expected behavior that must stay identical:** …

## 3. Motivation

- **Trigger:** security advisory | CVE | upstream deprecation | feature we need | sysadmin request
- **Link:** CVE ID, release notes URL, advisory URL

## 4. Branch

Per [`../guardrails/BRANCHING.md`](../guardrails/BRANCHING.md):

- Target branch: …
- If this is a version bump, confirm the *same* bump is applied (or skipped) on each of: `main`, `Aggrement`, `Prod-v1`. Pick one:
  - [ ] `main` only (canonical)
  - [ ] `main` + `Aggrement` (standard prod update)
  - [ ] `Prod-v1` only (Docker branch)
  - [ ] All branches (rare)

## 5. Compatibility audit

- [ ] I read the upstream changelog from the current pin → new pin.
- [ ] Breaking changes identified:
- [ ] Deprecations that will bite us at the **next** bump:

### Specific risk checks

- [ ] `python-pptx` changes — do they affect group traversal (invariant #7) or XML serialization (invariant #2)?
- [ ] `chromadb` changes — do they require a collection migration? (See [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md) §5.)
- [ ] `google-generativeai` changes — does the Gemini model name / endpoint still resolve?
- [ ] `fastapi` / `pydantic` changes — do our models in `app/models/models.py` still validate?
- [ ] `Pillow` changes — does `P` → `RGB` still work for our WMF/EMF path (invariant #5)?
- [ ] `streamlit` changes — does `baseUrlPath=accessibility` still route under nginx?

## 6. Test plan

- [ ] Full local stack restarted with new deps; `.pptx` end-to-end round-trip works.
- [ ] `python scripts/check_invariants.py` passes.
- [ ] `python -m pytest -q` passes.
- [ ] `python scripts/preflight.py` passes.
- [ ] Manual check of each invariant in §5 that I flagged "yes".

## 7. Prod considerations

- [ ] `pip install -r requirements.txt` on the VM (see [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md) §3).
- [ ] If this is a major version bump, I plan to `rm -rf venv && python3.11 -m venv venv && ...` instead of in-place upgrade. (`pip install` rarely removes packages cleanly.)
- [ ] If this changes Gemini model: I've coordinated with the course owner on cost/quota implications.
- [ ] If this changes nginx / systemd / firewall: I've noted that those files live on the VM, not in git (invariant #11), and I've captured the diff.

## 8. Rollback

- Rollback is usually "revert the PR, restart services". Call out deviations:
  - If the update pushed a Chroma schema forward, rollback also needs [`../ops/SOP_ROLLBACK.md`](../ops/SOP_ROLLBACK.md) §5.
  - If the update was a model name, rollback only needs a config edit + `systemctl restart`.

## 9. Docs updates

- [ ] [`../AGENT_CONTEXT.md`](../AGENT_CONTEXT.md) "Stable facts" — pin bumped.
- [ ] [`../PRODUCTION_ENVIRONMENT.md`](../PRODUCTION_ENVIRONMENT.md) — infra detail updated if relevant.
- [ ] [`../ops/SOP_DEPLOY.md`](../ops/SOP_DEPLOY.md) — if the deploy procedure itself changes (rare).
