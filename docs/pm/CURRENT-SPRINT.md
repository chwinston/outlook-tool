# Current Sprint

## Bolt 4 — Inbox Triage + Write Verbs (NEW 2026-11-16)

**Status:** OPEN — awaiting Chris to kick off Phase 1
**Goal:** Ship inbox triage capability: classify messages into safe-delete / needs-action / keep buckets, then safely apply archive/delete/mark-read operations. Three-phase ship (write methods → `/inbox-triage` skill → batching).
**Opened:** 2026-11-16
**Panel report:** [`docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`](../key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md)
**Plan:** [`docs/pm/plans/20261116-1100-inbox-triage-plan.md`](plans/20261116-1100-inbox-triage-plan.md)
**Backlog detail:** `BACKLOG.md` § Bolt 4 (13 F-items graduated to rows B4.1.1–B4.3.3)

---

### Phase 1 — Write methods on `OutlookClient` (~8-10 hours)

Three new methods (`move`, `delete`, `mark_read`) with unified `OperationResult` contract. COM + AppleScript + Graph backends. Round-trip verify mandatory. Dry-run default on CLI. Soft-delete default; hard-delete requires `--yes-really-permanent`.

| # | Item | Size | Status |
|---|------|------|--------|
| B4.1.1 | `OperationResult` dataclass | XS | todo |
| B4.1.2 | `OutlookClient.move()` dispatch | XS | todo |
| B4.1.3 | AppleScript `move()` + Archive folder detect-and-create + round-trip verify | S | todo |
| B4.1.4 | Graph `move()` returns `unsupported` without `Mail.ReadWrite` | S | todo |
| B4.1.5 | COM `move()` with round-trip verify | S | todo |
| B4.1.6 | `delete(permanent=False, yes_really_permanent=False)` across backends | S | todo |
| B4.1.7 | `mark_read(read=True)` across backends | S | todo |
| B4.1.8 | CLI verbs `move/delete/mark-read` — dry-run default | S | todo |
| B4.1.9 | Unit + integration tests (3 backends × 3 ops) | M | todo |
| B4.1.10 | CHANGELOG + README docs | S | todo |

### Phase 2 — `/inbox-triage` skill (~9 hours; blocked on Phase 1)

| # | Item | Size | Status |
|---|------|------|--------|
| B4.2.1 | `.claude/skills/inbox-triage/SKILL.md` | S | blocked |
| B4.2.2 | `scripts/triage.py` classifier | M | blocked |
| B4.2.3 | `scripts/triage_rules.py` (5 buckets) | S | executable (parallel with B4.2.2) |
| B4.2.4 | `scripts/apply_triage.py` executor | M | blocked |
| B4.2.5 | Rule provenance in classifier | XS | blocked |
| B4.2.6 | Live validation on real inbox (≥18/20 correct ship-gate) | S | blocked |
| B4.2.7 | `docs/manuals/inbox-triage.md` | S | blocked |

### Phase 3 — Batching (~3-4 hours; deferred ≥ 1 week post-ship)

See `BACKLOG.md` § B4.3.* — AppleScript bulk-move + Graph `$batch`. Triggered by live-use feedback.

### Success criteria

- [ ] Phase 1 complete — three write methods ship with unified `OperationResult`, round-trip verify, dry-run default CLI
- [ ] Phase 2 complete — `/inbox-triage` skill ships with classifier + executor + live-validation confirming ≥ 18/20 correct on safe-delete bucket sample
- [ ] Chris successfully triages inbox end-to-end: `triage` → review markdown → `apply --bucket safe-delete --apply`
- [ ] CHANGELOG + README + operator manual all updated
- [ ] 0 regressions in existing `search` / `send` / calendar surface

---

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
