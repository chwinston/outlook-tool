# Bolt 4 — Inbox Triage + Write Verbs: SHIPPED

**Date:** 2026-11-16
**Branch:** `bolt/20261116-inbox-triage-phase1-2`
**Mode:** Autonomous (/pm + /staff), /gitsmith wrapped
**Duration:** ~5 hours (Phase 1 build + Phase 2 build + walk-through + 3 apply passes + bug fix)
**Outcome:** 134 real soft-deletes on Chris's Inbox (266 → 132). Ship-gate 20/20 met.

## What shipped

**Phase 1 — Write Verbs (10 B4.1.* rows)**
- `OperationResult` dataclass contract across all 3 backends
- `move`, `delete` (soft default + hard-delete `--yes-really-permanent` gate), `mark_read` on AppleScript + COM + Graph stub
- CLI verbs with dry-run default; `--apply` required
- Round-trip verify (Tim F1 silent-success guard) — bug found + fixed mid-run
- 158 tests pass (3 Mac-integration deselected via marker)

**Phase 2 — /inbox-triage skill (7 B4.2.* rows)**
- `scripts/triage_rules.py` — 23 rules across 5 buckets (19 initial + 4 exemption rules from walk-through)
- `scripts/triage.py` classifier — emits `output/triage-YYYYMMDD-HHMM.json`
- `scripts/apply_triage.py` executor — dry-run default, `--apply` gate, 1-hour report freshness refusal, `--exclude-ids` for walk-through overrides
- `.claude/skills/inbox-triage/SKILL.md` + `docs/manuals/inbox-triage.md`
- Live validation: 20/20 correct on safe-delete spot-check

## Build sequence (14 commits)

1. `02bb29e` — gitsmith setup: branch foundation + gitignore + design doc
2. `3ef09b3` — OperationResult contract + test harness
3. `07fc1b2` — AppleScript mark_read, move, delete + round-trip verify
4. `7ca7133` — COM parity for all three verbs
5. `16b3d19` — Graph write-op stubs (scope-gated unsupported)
6. `1060c9d` — OutlookClient dispatch with double-flag guard
7. `64de005` — CLI verbs move/delete/mark-read (dry-run default)
8. `eb44a13` — CHANGELOG + README docs
9. `4aaa235` — triage_rules.py (19 rules, 5 buckets)
10. `e978864` — triage.py classifier
11. `f3ed107` — apply_triage.py executor + 39 tests
12. `240698a` — /inbox-triage skill + operator manual
13. `f52645a` — **fix:** large-mailbox verify (source-absence primary signal)
14. `c701cfc` — 4 exemption rules from walk-through

## What went well

- **Autonomous /pm + /staff cadence held together.** Phase 1 implementer + Phase 2 implementer each produced self-contained summaries; I read them, verified tests, moved on.
- **HARD PAUSE worked.** Chris reviewed dry-run markdown previews (4 bucket previews, 161 total rows) before any real write.
- **Walk-through surfaced real value.** The cluster-by-cluster review caught 4 subject categories the classifier was misfiling. Those became permanent exemption rules — next run won't ask.
- **Mid-run bug fix was clean.** First apply pass surfaced the verify-false-fail; diagnosed it from the result file + Inbox scan, fixed `_verify_present_in_folder` semantics (source-absence beats target-presence on large mailboxes), updated tests, re-ran. ~45 min detour.

## What hurt

- **First apply pass was polluted** by the verify bug — 87 attempts reported 72 failures, but 63 real moves happened. Had to scan the Inbox manually to figure out real success count. Now that verify is fixed (`status="success", verified=False` for UNVERIFIED-success), future runs will tally accurately.
- **AppleScript silent-success is real and persistent.** Final pass showed 15 real failures where `move` script returned OK but message stayed in Inbox. Likely an index-shift issue during bulk operations (move message N invalidates indices of messages N+1..). Phase 3 batching (bulk-move primitive) will be the real fix. For now: re-run until clean.
- **No upfront sample of Deleted Items size.** If I'd known Chris has 5000+ messages in Deleted Items, I'd have specced the scan_limit at 500 from the start instead of discovering it in production. Lesson: ask about mailbox size during brainstorm, not after first apply.

## Metrics

| Metric | Value |
|---|---|
| Lines added | ~1,800 (not counting tests) |
| Lines of tests | ~450 (39 new triage tests + 5 updated write-ops tests) |
| Test count | 158 passing (3 integration deselected) |
| Inbox size before | 266 messages in 30-day window |
| Inbox size after | 132 messages |
| Real soft-deletes | 134 |
| Ship-gate | 20/20 correct on spot-check |
| 5 messages kept per walk-through decision | Google Cloud x2, MSI billing alert, Google Account recoverability, SharePoint Jungle permission |
| 15 genuine AppleScript silent-success failures remaining | Documented; Phase 3 batching will address |

## Follow-ups

**Open (nothing blocking):**
- Phase 3 batching (B4.3.*) stays deferred with named trigger: "Phase 1+2 in use ≥ 1 week OR 500+ message op complaint." The silent-success rate observed (~15/70 = 21%) might be the complaint.
- Continue running `/inbox-triage` daily or on-demand. Each run will add more aged messages to safe-delete as they cross the 7/14/30-day thresholds.
- Re-run on remaining Inbox (132 messages) in a few days — some will have aged into safe-delete.

**Not opened:**
- Graph `Mail.ReadWrite` scope — still requires IT ticket, still not needed because Chris is Mac-primary with Outlook always running.
- ML classifier — rule count is 23, override rate is 5/92 ~= 5%, both below trigger thresholds.
- Cron scheduling — manual-trigger is fine for the current review-first workflow.

## Links

- Design doc: `docs/superpowers/specs/2026-11-16-bolt4-autonomous-execution-design.md`
- Panel approval: `docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`
- Detailed plan: `docs/pm/plans/20261116-1100-inbox-triage-plan.md`
- Apply result (pass 2): `output/apply-result-20261116-2210.md`
- Apply result (pass 3): `output/apply-result-20261116-2225.md`
- Operator manual: `docs/manuals/inbox-triage.md`
