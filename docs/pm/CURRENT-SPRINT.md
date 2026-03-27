# Current Sprint

## Bolt 3 — Security Hardening & Code Quality (Staff Panel + Security Audit)

**Status:** COMPLETE
**Goal:** Fix all P0/P1/P2 findings from the combined staff panel review and security audit.
**Opened:** 2026-03-27
**Branch:** bolt/20260327-security-quality-sweep

---

### P0 — Fix Now

| # | Item | Size | Status | Source |
|---|------|------|--------|--------|
| 1 | Path traversal in attachment downloads — sanitize filenames | S | todo | Security F1 |
| 2 | `--json` outputs human text + JSON — skip human output when --json | S | todo | Staff (Rob) |
| 3 | Token cache write race condition — atomic file creation | S | todo | Security F4 |

### P1 — Fix This Sprint

| # | Item | Size | Status | Source |
|---|------|------|--------|--------|
| 4 | OData injection via subject_contains in Graph API | S | todo | Security F2 |
| 5 | Graph API IDs not URL-encoded in download URL | S | todo | Security F3 |
| 6 | Backend data-transformation tests (mocked COM/subprocess/requests) | M | todo | Staff (Fran) |
| 7 | Deduplicate multi-folder search logic into shared helper | S | todo | Staff (Rob) |
| 8 | subject_contains double-application + truncation-before-filter | M | todo | Staff (Rob) |
| 9 | Add timeout/retry/429 handling to Graph API calls | M | todo | Staff (Tim) |

### P2 — Harden

| # | Item | Size | Status | Source |
|---|------|------|--------|--------|
| 10 | importance silently ignored on AppleScript send — warn or pass | S | todo | Staff (Fran) |
| 11 | Unstable hash() for AppleScript event IDs — use hashlib | S | todo | Staff (Fran) |
| 12 | Win32 COM cache — scope to single search call | S | todo | Staff (Tim) |
| 13 | msal/requests as optional extras in pyproject.toml | S | todo | Staff (Al) |
| 14 | Validate msg_index/att_index as int in save_attachment | S | todo | Security F7 |
| 15 | Strip null bytes in _escape_applescript | S | todo | Security F14 |

### Success Criteria

- [ ] All P0 items fixed and tested
- [ ] All P1 items fixed and tested
- [ ] All P2 items fixed
- [ ] 50+ tests passing (expect 60+ with new backend tests)
- [ ] Security audit re-run shows 0 High findings
