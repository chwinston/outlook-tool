# Bolt 5 — Phase 3 Batching + Silent-Success Root Cause Fix

**Date opened:** 2026-11-17
**Prompted by:** Live-run reliability issue observed across 4 apply passes on real Inbox
**Prior panel approval:** 2026-11-16 (Tim/Rob/Fran/Al/Will) — plan was green-lit as B4.3.*, deferred with named triggers

## Trigger evidence (why the deferred trigger is now met)

Observed across 4 apply-triage runs 2026-11-16 → 2026-11-17 on Chris's real Inbox:

| Round | Attempted | Verified-success | UNVERIFIED-success | Real silent-success failures | Failure rate |
|---|---|---|---|---|---|
| Round 1 (buggy verify) | 87 | 15 | 0 | 72 (false!) | N/A (metric wrong) |
| Round 2 | 67 | 39 | 8 | 20 | 30% |
| Round 3 | 56 | 30 | 11 | 15 | 27% |
| Round 4 | 43 | 24 | 11 | 8 | 19% |
| **Round 4 post-hoc audit** | 43 | 15 (truly moved) | — | 19 real failures (Inbox scan) | **44%** |

The trigger stated in BACKLOG § B4.3 was: "Phase 1+2 in use ≥ 1 week OR 500+ message op complaint." The 1-week clock hasn't run, but the reliability characteristic that "500+ message op complaint" was a proxy for has **already materialized at 43-message sizes** — AppleScript silent-success is not an edge case, it's the modal failure. Opening Bolt 5 instead of waiting.

## Scope

### B5.1 — Root-cause the AppleScript silent-success (BLOCKING)

Before implementing bulk-move, understand why single-message `move message N of inbox` silently fails. Hypotheses to test:

1. **Index-shift race** — moving message N invalidates messages N+1..N+k's indices mid-loop. The script issues the next move against an index that now points at a different message (or an empty slot).
2. **Exchange sync lag** — AppleScript's move returns when the client-side state changes, but Exchange server hasn't ACKed. A subsequent read-back operation sees the message still in Inbox because the sync hasn't settled.
3. **Verify-step side effect** — the post-move search in `_verify_present_in_folder` opens a new AppleScript session that resets some client-side state.

Instrument the apply loop to log AppleScript stdout/stderr per call. Compare behavior with:
- 2-second sleep between calls (apparent fix observed in round 4 pick-off experiment)
- Index-descending order (reserve against hypothesis 1)
- Single `tell application "Microsoft Outlook"` block wrapping multiple moves (probe hypothesis 2)

### B5.2 — AppleScript bulk-move primitive (was B4.3.1)

Based on B5.1 findings, implement bulk-move. Two candidate shapes:

**Shape A (if hypothesis 1 confirmed):** AppleScript-side bulk inside one tell block
```applescript
tell application "Microsoft Outlook"
  set m1 to message 5 of inbox
  set m2 to message 7 of inbox
  set m3 to message 9 of inbox
  move m1 to mail folder "Archive" of (first exchange account)
  move m2 to mail folder "Archive" of (first exchange account)
  move m3 to mail folder "Archive" of (first exchange account)
end tell
```
This captures all target references BEFORE any move, so index shifts don't matter.

**Shape B (if hypothesis 2 confirmed):** Python-side batching with Exchange sync wait
```python
for id in ids:
    move_one(id)
    wait_for_sync()  # poll Exchange sync status via AppleScript
```

**Shape C (if all else fails):** OMG-wait pattern with descending-index order + 2s sleep
```python
for id in sorted(ids, reverse=True):
    move_one(id)
    time.sleep(2)
```

Decision per hypothesis-test results in B5.1.

### B5.3 — Verify-step hardening (was in-scope of B4.2.6 but kept surfacing)

Round-trip verify's source-absence check uses `scan_limit=200` on the source folder. Empirically this misses messages when the Inbox is actively being mutated by other running processes. Options:

- Keep IDs resolved during pre-fetch; verify that EntryID is no longer resolvable (more reliable than fingerprint scan)
- Scan the whole source folder (slow but correct for N < 500)
- Accept that AppleScript source-absence is fundamentally unreliable mid-batch and move the primary verify signal to target-presence after a settle delay

### B5.4 — Graph `$batch` POST (was B4.3.2)

**Still deferred.** Requires `Mail.ReadWrite` scope, which requires IT ticket. Open once scope is granted.

### B5.5 — Benchmark 500-message archive (was B4.3.3)

Add `scripts/bench_triage.py`:
- Seed 500 messages into `_bolt5_bench` folder via AppleScript
- Run 3 paths: current one-at-a-time, descending-index with sleep, bulk-move primitive
- Report: wall-clock time, success count, silent-success count

Ship-gate: ≥ 95% verified-success on the bulk-move path.

## Non-goals

- COM parity for bulk-move (defer until Windows user appears — same rationale as Phase 1)
- Graph `$batch` (blocked on scope acquisition)
- Changes to `/inbox-triage` skill UX (the two-command workflow is working)

## Rough sizing

- B5.1 root-cause: M (3-4 hours) — hypothesis testing against disposable folder + instrumentation
- B5.2 bulk-move: M (2-3 hours, depends on B5.1 outcome)
- B5.3 verify hardening: S (90 min)
- B5.4 Graph batch: blocked
- B5.5 benchmark: S (60 min)

Total: ~7-9 hours.

## Success criteria

- [ ] B5.1: silent-success root cause identified and documented (which hypothesis won)
- [ ] B5.2: Bulk-move primitive ships with `OperationResult[]` return and per-message verified status
- [ ] B5.3: Verify step achieves ≥ 95% precision on real Inbox (no more phantom UNVERIFIED-success)
- [ ] B5.5: Benchmark proves bulk-move ≥ 5× faster than one-at-a-time with ≥ 95% success rate
- [ ] Zero regressions on existing `/inbox-triage` workflow — same two-command UX, same CLI
- [ ] CHANGELOG entry documenting the silent-success fix

## Appendix: pick-off experiment (2026-11-17)

Mid-session workaround: after round 4 apply returned 19 real + 11 UNVERIFIED-success failures (44% silent-success), ran 34 remaining safe-delete messages one-at-a-time via `cli.py delete <id> --apply` with 2s sleep between calls. Processed from highest message_id to lowest to avoid index-shift. Result TBD pending completion — if success rate approaches 100%, that confirms hypothesis 1 (index-shift race) and shape C becomes the minimum viable fix.
