# Current Sprint

## Bolt 5 — Silent-Success Fix + Batching (OPEN 2026-11-17)

**Status:** OPEN — not yet kicked off
**Trigger:** B4.3.* graduated early. Observed 44% silent-success rate on 43-msg apply run — reliability pain the deferred trigger was a proxy for.
**Goal:** Root-cause AppleScript silent-success, implement bulk-move primitive, harden verify step to eliminate phantom UNVERIFIED-success reports.
**Spec:** [`docs/pm/plans/20261117-0740-bolt5-batching-plan.md`](plans/20261117-0740-bolt5-batching-plan.md)
**Backlog:** `BACKLOG.md` § Bolt 5
**Rough sizing:** ~7-9 hours

### Entry point

**B5.1** (root-cause) is blocking B5.2 (bulk-move). Three hypotheses to test against disposable folder with instrumentation. The pick-off experiment 2026-11-17 (one-at-a-time with 2s sleep in descending-index order) is partial evidence toward hypothesis 1 (index-shift race).

---

## Bolt 4 — Inbox Triage + Write Verbs (SHIPPED 2026-11-16)

**Status:** ✅ COMPLETE — Phase 1 + Phase 2 shipped same-day via autonomous execution
**Outcome:** 134 real soft-deletes on Chris's Inbox (266→132). Ship-gate 20/20 met.
**Opened:** 2026-11-16 · **Closed:** 2026-11-16 (~5 hours of autonomous build + walk-through)
**Panel report:** [`docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`](../key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md)
**Plan:** [`docs/pm/plans/20261116-1100-inbox-triage-plan.md`](plans/20261116-1100-inbox-triage-plan.md)
**Design doc:** [`docs/superpowers/specs/2026-11-16-bolt4-autonomous-execution-design.md`](../superpowers/specs/2026-11-16-bolt4-autonomous-execution-design.md)
**Captain's log:** [`docs/captains_log/20261116-2230-bolt4-inbox-triage-shipped.md`](../captains_log/20261116-2230-bolt4-inbox-triage-shipped.md)
**Backlog detail:** `BACKLOG.md` § Bolt 4 (all B4.1.* + B4.2.* ✅ done; B4.3.* stays deferred)
**Branch:** `bolt/20261116-inbox-triage-phase1-2` (14 commits; PR open to main)

---

### Phase 1 — Write methods on `OutlookClient` (~8-10 hours)

Three new methods (`move`, `delete`, `mark_read`) with unified `OperationResult` contract. COM + AppleScript + Graph backends. Round-trip verify mandatory. Dry-run default on CLI. Soft-delete default; hard-delete requires `--yes-really-permanent`.

| # | Item | Size | Status |
|---|------|------|--------|
| B4.1.1 | `OperationResult` dataclass | XS | ✅ done |
| B4.1.2 | `OutlookClient.move()` dispatch | XS | ✅ done |
| B4.1.3 | AppleScript `move()` + Archive folder detect-and-create + round-trip verify | S | ✅ done |
| B4.1.4 | Graph `move()` returns `unsupported` without `Mail.ReadWrite` | S | ✅ done |
| B4.1.5 | COM `move()` with round-trip verify | S | ✅ done |
| B4.1.6 | `delete(permanent=False, yes_really_permanent=False)` across backends | S | ✅ done |
| B4.1.7 | `mark_read(read=True)` across backends | S | ✅ done |
| B4.1.8 | CLI verbs `move/delete/mark-read` — dry-run default | S | ✅ done |
| B4.1.9 | Unit + integration tests (3 backends × 3 ops) | M | ✅ done |
| B4.1.10 | CHANGELOG + README docs | S | ✅ done |

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

- [x] Phase 1 complete — three write methods ship with unified `OperationResult`, round-trip verify, dry-run default CLI
- [x] Phase 2 complete — `/inbox-triage` skill ships with classifier + executor + live-validation confirming ≥ 18/20 correct on safe-delete bucket sample (20/20 hit)
- [x] Chris successfully triages inbox end-to-end: `triage` → review markdown → `apply --bucket safe-delete --apply` (134 soft-deletes, 266→132)
- [x] CHANGELOG + README + operator manual all updated
- [x] 0 regressions in existing `search` / `send` / calendar surface (158 tests pass)

### Build footnotes (mid-run findings)

- **Verify bug found + fixed** (commit f52645a): post-move verify was scanning only the first 50 messages of Deleted Items, false-failing on large mailboxes. Now uses source-absence as primary signal + target-presence scan 500 as secondary; UNVERIFIED-success tier surfaces correctly.
- **4 exemption rules added from walk-through** (commit c701cfc): Google Cloud project-at-risk, Google Account recoverability, MSI billing alerts, SharePoint permission changes — all route to `needs-action` so classifier can't sweep them again.
- **15 real silent-success failures** remain — genuine AppleScript limitation where the `move` command returns OK but the message stays. Known in Phase 1 handoff notes; Phase 3 batching (deferred) will address via bulk-move primitive.

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
| 1 | Path traversal in attachment downloads — sanitize filenames | S | ✅ done | Security F1 |
| 2 | `--json` outputs human text + JSON — skip human output when --json | S | ✅ done | Staff (Rob) |
| 3 | Token cache write race condition — atomic file creation | S | ✅ done | Security F4 |

### P1 — Fix This Sprint

| # | Item | Size | Status | Source |
|---|------|------|--------|--------|
| 4 | OData injection via subject_contains in Graph API | S | ✅ done | Security F2 |
| 5 | Graph API IDs not URL-encoded in download URL | S | ✅ done | Security F3 |
| 6 | Backend data-transformation tests (mocked COM/subprocess/requests) | M | ✅ done | Staff (Fran) |
| 7 | Deduplicate multi-folder search logic into shared helper | S | ✅ done | Staff (Rob) |
| 8 | subject_contains double-application + truncation-before-filter | M | ✅ done | Staff (Rob) |
| 9 | Add timeout/retry/429 handling to Graph API calls | M | ✅ done | Staff (Tim) |

### P2 — Harden

| # | Item | Size | Status | Source |
|---|------|------|--------|--------|
| 10 | importance silently ignored on AppleScript send — warn or pass | S | ✅ done | Staff (Fran) |
| 11 | Unstable hash() for AppleScript event IDs — use hashlib | S | ✅ done | Staff (Fran) |
| 12 | Win32 COM cache — scope to single search call | S | ✅ done | Staff (Tim) |
| 13 | msal/requests as optional extras in pyproject.toml | S | ✅ done | Staff (Al) |
| 14 | Validate msg_index/att_index as int in save_attachment | S | ✅ done | Security F7 |
| 15 | Strip null bytes in _escape_applescript | S | ✅ done | Security F14 |

### Success Criteria

- [ ] All P0 items fixed and tested
- [ ] All P1 items fixed and tested
- [ ] All P2 items fixed
- [ ] 50+ tests passing (expect 60+ with new backend tests)
- [ ] Security audit re-run shows 0 High findings
