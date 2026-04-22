# Deploy Log

One line per prod deploy or rollback. Newest at the top. Keep it terse — post-mortems live in [`INCIDENT_LOG.md`](INCIDENT_LOG.md).

Format:

```
YYYY-MM-DD HH:MM UTC  <handle>  <old-sha> → <new-sha>  (branch)  notes: <one line>
YYYY-MM-DD HH:MM UTC  <handle>  ROLLBACK  <bad-sha> → <old-sha>  reason: <one line>
```

---

<!-- Append new entries above this line. -->

2026-04-22 —  framework  —  n/a → n/a  (docs-only)  notes: initial maintainability framework added (no prod change)
