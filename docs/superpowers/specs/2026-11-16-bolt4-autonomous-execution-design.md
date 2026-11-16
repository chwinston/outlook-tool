# Bolt 4 — Autonomous Execution Design (Phase 1 + Phase 2)

**Date:** 2026-11-16
**Status:** Approved by Chris (2026-11-16)
**Mode:** Full autonomous — /pm + /staff drive from start to finish; /gitsmith wraps.

## Intent & Why

Ship Bolt 4 Phase 1 (write verbs) + Phase 2 (`/inbox-triage` skill) in one autonomous run.
Chris is Mac-primary, inbox is noisy, panel pre-approved the plan 2026-11-16.
Goal: end session with live-validated triage capability on real inbox, ≥ 18/20 correct on safe-delete bucket.

## Goals

1. `move`, `delete`, `mark_read` exposed as first-class CLI verbs across AppleScript + COM + Graph stub backends
2. Unified `OperationResult` contract; round-trip verify mandatory
3. Dry-run default; `--apply` required for real writes; `--yes-really-permanent` for hard-delete
4. `/inbox-triage` skill: classify → preview markdown → apply by bucket
5. Live validation on real inbox, ship-gate = ≥ 18/20 correct on safe-delete sample

## Non-Goals

- No Graph `Mail.ReadWrite` scope acquisition (future IT ticket)
- No ML classifier (rules only; graduation trigger named in backlog)
- No cron scheduling
- No Phase 3 batching (deferred ≥ 1 week post-ship)

## Anti-Goals (would mean failure even if "working")

- Any `--apply` fires on real mailbox before explicit Chris approval
- Any hard-delete during live validation (soft-delete only)
- Premature abstraction of `OperationResult` for hypothetical future verbs (`flag`, `categorize`)
- Tests that pass against mocks but silently fail against real Outlook (Tim's F1 — silent success)

## Constraints

- Chris is on Mac; AppleScript is primary path; COM is untestable here but code-complete for Windows
- Graph `move/delete` return `status="unsupported"` without `Mail.ReadWrite`; do NOT conflate with 403
- AppleScript message IDs are session-scoped — triage report `>1h` old is rejected; fallback identity = sender+subject+received_date
- COM `.Move()` returns the moved item; capture it (original handle is invalidated after move)
- Test fixture = disposable `_bolt4_test` folder on real mailbox (created + torn down per run)

## Approaches considered

**A. Horizontal:** `OperationResult` → all three backends `move()` → all three `delete()` → …
- Fast for refactor; weak for de-risking. Panel rejected: contract isn't proven until a second backend consumes it.

**B. Vertical slice (chosen):** `OperationResult` → test harness → AppleScript `mark_read` end-to-end → replicate pattern per verb → COM parity → Graph stub → CLI → tests → docs
- Staff panel's recommendation. First commit stresses the full stack. Every subsequent row is pattern-replication.

**C. Classifier-first:** Build Phase 2 rules before write methods ship
- Rejected. Rules without executable verbs is theater. Also risks speculative rule-writing (Tim: "40 rules, 20% precision").

**Chosen: B.** Rationale: proves the contract on real code before replication; harness lands as side-effect of first vertical slice.

## The How (as decided)

1. **/gitsmith setup** — branch `bolt/20261116-inbox-triage-phase1-2` off `main`; verify clean working tree; stash `.audit.db` untracked state
2. **Phase 1 implementer (general-purpose agent, staff-panel voice)** — executes in this order:
   - B4.1.1 `OperationResult` dataclass
   - Test harness scaffold (`tests/test_write_ops.py` + disposable-folder fixture)
   - B4.1.7 AppleScript `mark_read` (safest verb first; proves contract + verify loop)
   - B4.1.3 AppleScript `move` (Archive detect-and-create)
   - B4.1.6 AppleScript `delete` (soft-delete default)
   - B4.1.5 COM parity for all three verbs (code-complete; integration-skipped locally)
   - B4.1.4 Graph stub (returns `status="unsupported"` with clear error)
   - B4.1.2 `OutlookClient` dispatch methods
   - B4.1.8 CLI verbs (`move` / `delete` / `mark-read`) — dry-run default; `--apply` gate; `--yes-really-permanent` for hard-delete
   - B4.1.9 Unit tests (all) + integration tests (Mac only; skip COM/Graph integration)
   - B4.1.10 CHANGELOG + README docs
   - Commit per logical unit; green test suite before moving on
3. **Phase 2 implementer (same shape agent, fresh context)**:
   - B4.2.1 `.claude/skills/inbox-triage/SKILL.md`
   - Sample real inbox (`outlook-tool search --limit 500 --json`) → hand-bucket in head → write B4.2.3 rules from observed patterns
   - B4.2.2 `scripts/triage.py` classifier
   - B4.2.5 rule provenance in classifier output
   - B4.2.4 `scripts/apply_triage.py` executor (reads report; prints markdown preview; refuses reports >1h old; `--apply` gate)
   - B4.2.7 `docs/manuals/inbox-triage.md`
4. **HARD PAUSE** — autonomous run halts. Agent produces dry-run markdown preview for Chris to eyeball. Chris types "apply" or redirects.
5. **Live validation (B4.2.6)** — after Chris approval:
   - Run `apply_triage.py --bucket safe-delete --apply` on fresh report
   - Verify destination folder contents
   - Compute hit-rate; log results to captain's log
   - Ship-gate: ≥ 18/20 correct on sample
6. **/pm close-out**:
   - Update `BACKLOG.md` (mark B4.1.* + B4.2.* done; keep B4.3.* deferred)
   - Update `CURRENT-SPRINT.md` (status: COMPLETE)
   - Write captain's log entry under `docs/captains_log/`
   - Open PR against `main`
7. **/gitsmith audit** — verify branch clean; no stray files; commits well-formed; `.audit.db` not committed; `output/` artifacts gitignored

## Hidden risks baked into code

- **COM move handle invalidation:** `moved = msg.Move(target); return result_from(moved)` — never re-query by old EntryID
- **AppleScript session ID staleness:** fallback identity on apply = (sender, subject, received_date) match if direct ID lookup fails; abort cleanly if ambiguous
- **Graph 403 vs scope-missing:** distinguish by checking granted scopes on token; surface both with distinct error messages
- **Round-trip verify:** `OperationResult.verified=True` requires re-fetch after op; `verified=False` surfaces to caller even if op returned success

## What's left to the builder

- Rule-set contents (sample-driven from Chris's real inbox, not prescribed here)
- Test-folder cleanup specifics (best-effort teardown in `finally`)
- CHANGELOG wording
- Markdown preview format for dry-run (concise table; sender · subject · bucket · matched_rule)

## Success criteria

- [ ] `bolt/20261116-inbox-triage-phase1-2` branch exists, clean, pushed
- [ ] All Phase 1 tests green (unit + Mac integration against `_bolt4_test` folder)
- [ ] `/inbox-triage` skill present and documented
- [ ] Dry-run preview rendered on real inbox and reviewed by Chris
- [ ] Live `--apply` run executed on safe-delete bucket with ≥ 18/20 correct
- [ ] Captain's log entry written
- [ ] PR opened against `main`
- [ ] `.audit.db` and `output/` artifacts not tracked

## Approval

Approved by Chris 2026-11-16 ("go"). Authorizes autonomous execution through live validation, with the one hard pause at step 4 before any `--apply` touches real mail.
