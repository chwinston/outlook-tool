# Backlog

**Last groomed:** 2026-11-16 (added Bolt 4 — Inbox Triage + Write Verbs. Panel-approved 2026-11-16 by Tim/Rob/Fran/Al/Will. Three-phase ship: Phase 1 = write methods + `OperationResult` contract, Phase 2 = `/inbox-triage` skill + classifier/executor, Phase 3 = batching for throughput. 13 F-items graduated to rows below. Panel report: `docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`. Plan: `docs/pm/plans/20261116-1100-inbox-triage-plan.md`.) Prior: 2026-03-27 (Bolt 3 close).

## Completed

| # | Item | Bolt |
|---|------|------|
| 1-8 | Security hardening (F1-F10), CI pipeline | Bolt 1 |
| 9 | Create SECURITY.md | Bolt 2 |
| 10 | Create REQUIREMENTS.md with REQ IDs | Bolt 2 |
| 11 | Calendar event search (all 3 backends) | Bolt 2 |

## Prioritized Items

| # | Item | Size | Status | Notes |
|---|------|------|--------|-------|
| 1 | Fix COM cache key (use stable ID instead of `id(msg)`) | S | executable | Deferred — panel said revisit if users report issues |
| 2 | Move token cache to OS keychain | M | executable | Deferred — chmod sufficient for now |
| 3 | Create calendar events / send meeting invites (REQ-010) | M | executable | Could priority — add when needed |
| 4 | List available calendars (REQ-009) | S | executable | Could priority |
| 5 | Contacts search | M | needs-spec | Not yet in requirements |

## Bolt 4 — Inbox Triage + Write Verbs (NEW 2026-11-16, panel-approved)

Spec: `docs/pm/plans/20261116-1100-inbox-triage-plan.md`. Panel report: `docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`.

### Phase 1 — Write methods on `OutlookClient` (~8-10 hours total)

| # | Item | Size | Status | Panel F-tag | Notes |
|---|------|------|--------|-------------|-------|
| B4.1.1 | Add `OperationResult` dataclass to `outlook_tool.py` (shared contract across backends) | XS (15 min) | executable | F4 (Rob F1) | `{status, verified, message_id, operation, before_state, after_state, error_detail, backend, elapsed_ms}` |
| B4.1.2 | Add `OutlookClient.move(message_id, target_folder) → OperationResult` dispatch method | XS (15 min) | executable | — | Thin dispatch to backend-specific impl |
| B4.1.3 | `_AppleScriptBackend.move()` with Archive folder detect-and-create + round-trip verify | S (90 min) | executable | F1 (Tim) + F3 (Al) | Mac-primary path; silent-success is the dominant risk; verify catches it |
| B4.1.4 | `_GraphBackend.move()` returns `status="unsupported"` without `Mail.ReadWrite` (don't auto-upgrade) | S (45 min) | executable | F2 (Rob) | Surface the scope cliff loudly |
| B4.1.5 | COM backend `move()` with round-trip verify (Windows parity) | S (60 min) | executable | F1 (Tim) | Windows path |
| B4.1.6 | `OutlookClient.delete(message_id, permanent=False, yes_really_permanent=False)` — soft-delete default, hard-delete requires double-flag | S (60 min) | executable | F6 (Tim + unanimous) | Dangerous op; guarded by design |
| B4.1.7 | `OutlookClient.mark_read(message_id, read=True)` across all three backends | S (45 min) | executable | — | Lowest-risk write; rounds out surface |
| B4.1.8 | CLI verbs `outlook-tool move/delete/mark-read` — dry-run default, `--apply` required for real write | S (60 min) | executable | F5 (Panel unanimous) | Safe-by-default CLI shape |
| B4.1.9 | Unit + integration tests: 3 backends × 3 ops. Graph mocked with `respx`. AppleScript tested against disposable mailbox folder locally (not in CI). | M (120 min) | executable | — | Correctness guard |
| B4.1.10 | CHANGELOG entry + README section documenting new verbs + `yes_really_permanent` safety | S (30 min) | executable | — | Operator docs |

### Phase 2 — `/inbox-triage` skill (~9 hours total, blocked on Phase 1 complete)

| # | Item | Size | Status | Panel F-tag | Notes |
|---|------|------|--------|-------------|-------|
| B4.2.1 | Create `.claude/skills/inbox-triage/SKILL.md` | S (30 min) | blocked on Phase 1 | — | Two-step workflow: classify → apply |
| B4.2.2 | `scripts/triage.py` — classifier reading inbox via `outlook_tool.search`, emitting `output/triage-YYYYMMDD-HHMM.json` | M (180 min) | blocked on Phase 1 | F10 (Fran) | Classifier is its own concern |
| B4.2.3 | `scripts/triage_rules.py` — initial rule base: 5 buckets (safe-delete / resolved-ticket / low-signal-client / needs-action / keep) | S (90 min) | executable (parallel with B4.2.2) | F10 (Fran) | Single Python file; grows additively |
| B4.2.4 | `scripts/apply_triage.py` — reads report, prints markdown preview, applies when `--apply` passed. Supports `--bucket`. Refuses reports > 1 hour old. | M (120 min) | blocked on Phase 1 | F10 + F11 | Executor is separate from classifier |
| B4.2.5 | Add rule provenance (`matched_rule`) to classifier output | XS (45 min) | blocked on B4.2.2, B4.2.3 | F9 (Rob) | Debuggability — tune the rule, not retrain a model |
| B4.2.6 | Live validation on Chris's real inbox: produce triage report, review markdown, spot-check safe-delete bucket, confirm ≥ 18/20 correct before shipping | S (60 min) | blocked on B4.2.1-5 | — | Ship-gate acceptance criterion |
| B4.2.7 | Operator manual `docs/manuals/inbox-triage.md` | S (45 min) | blocked on B4.2.1-5 | — | Two-command workflow docs |

### Phase 3 — Batching (~3-4 hours total, deferred until Phase 1 + 2 in use ≥ 1 week)

| # | Item | Size | Status | Panel F-tag | Notes |
|---|------|------|--------|-------------|-------|
| B4.3.1 | AppleScript bulk-move: single `move messages {id1, id2, ...} to folder "Archive"` per target folder | S (90 min) | deferred (with trigger) | F7 (Al) | 5-10× speedup; trigger = Phase 1+2 in use ≥ 1 week OR 500+ message op complaint |
| B4.3.2 | Graph `$batch` POST up to 20 ops/call, honor `Retry-After` on 429 | S (90 min) | deferred (with trigger) | F8 (Al) | Only relevant after `Mail.ReadWrite` granted; same trigger as 4.3.1 |
| B4.3.3 | Benchmark 500-message archive before + after Phase 3 (`scripts/bench_triage.py`) | XS (30 min) | deferred | — | Prove the win |

### Deferred from Bolt 4 (with named triggers)

| Item | Source | Trigger to revisit |
|---|---|---|
| ML classifier for sender/subject clustering | Panel F12 (not-graduating) | Rule count > 40 OR Chris's bucket-override rate > 10% of classifications |
| Recurring scheduled triage via cron | Panel F13 (not-graduating) | After 4 consecutive weeks of manual use with < 5% misclassifications |
| Cross-mailbox delegation (triage a shared mailbox) | Deferred in plan § out-of-scope | First time Chris asks to triage someone else's mailbox |
| Graph `Mail.ReadWrite` scope acquisition (IT ticket) | Deferred in plan § out-of-scope | First time Chris needs to triage from a machine where Outlook isn't running |
| Hard-delete batch mode | Panel ruled explicit NO-BUILD | Never — foot-gun by design; `--permanent --yes-really-permanent` is per-message on purpose |
