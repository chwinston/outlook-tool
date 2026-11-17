# Bolt Plan — Inbox Triage + Write Verbs

**Created:** 2026-11-16 11:00
**Author:** `/pm` + `/staff` panel (Tim, Rob, Fran, Al, Will Larson)
**Companion panel report:** [`docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`](../../key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md)
**Status:** APPROVED — ready for implementation, Chris to execute.

---

## Goal

Add inbox-triage capability to `outlook-tool`: classify messages into safe-delete / needs-action / keep buckets, then safely apply archive/delete/mark-read operations. Three-phase ship: Phase 1 (write methods), Phase 2 (`/inbox-triage` skill), Phase 3 (batching for throughput, deferred).

## Scope

### In scope for v1 (Phases 1 + 2)

- Three new methods on `OutlookClient`: `move`, `delete`, `mark_read`
- Unified `OperationResult` dataclass contract across backends (COM / AppleScript / Graph)
- Round-trip verify on every write
- AppleScript backend: detect-and-create Archive folder on first use
- Graph backend: returns `status="unsupported"` without `Mail.ReadWrite` scope
- CLI verbs: `outlook-tool move/delete/mark-read` — dry-run default, `--apply` required for real writes
- `yes_really_permanent=True` requirement for hard-delete; soft-delete to Deleted Items is the default
- New Claude Code skill `/inbox-triage` with two-verb workflow (`triage` → `apply`)
- Classifier with 5 buckets + rule provenance (`matched_rule` field)
- Executor that reads a triage JSON report, prints markdown preview, applies when `--apply` passed
- Executor rejects triage reports > 1 hour old
- Live validation against Chris's real inbox before shipping
- Operator manual at `docs/manuals/inbox-triage.md`

### Out of scope for v1 (explicit)

- **Batching.** Phase 3 (AppleScript bulk-move + Graph `$batch`) ships separately after ≥ 1 week of live use.
- **ML classifier.** V1 is heuristic rules only. Trigger to revisit: rule count > 40 OR override rate > 10%.
- **Recurring scheduled triage.** Manual trigger is v1; scheduled after proven over weeks.
- **Cross-mailbox delegation.** Chris's mailbox only; delegation revisits when first needed.
- **Graph `Mail.ReadWrite` scope.** Requires IT ticket; AppleScript/COM paths are sufficient on Chris's primary machine.

## Phased approach

### Phase 0 — Design gate ✅ COMPLETE 2026-11-16

Panel-approved; see companion report. No open decisions.

### Phase 1 — Write methods on OutlookClient

**Deliverable:** Three write methods on `OutlookClient` returning a unified `OperationResult`, backed by COM / AppleScript / Graph implementations. CLI verbs expose them with dry-run defaults.

**Tasks:**

| # | Task | File(s) | Effort | Depends on |
|---|---|---|---|---|
| 1.1 | Add `OperationResult` dataclass to `outlook_tool.py` (shared contract per Rob F1) | `outlook_tool.py` | 15 min | — |
| 1.2 | Add `OutlookClient.move(message_id, target_folder) → OperationResult` (dispatch method) | `outlook_tool.py` (OutlookClient) | 15 min | 1.1 |
| 1.3 | `_AppleScriptBackend.move()` with Archive folder detect-and-create (Al F3) + round-trip verify (Tim F1) | `outlook_tool.py` (AppleScript backend) | 90 min | 1.2 |
| 1.4 | `_GraphBackend.move()` — returns `status="unsupported"` without `Mail.ReadWrite` (Rob F2) | `outlook_tool.py` (Graph backend) | 45 min | 1.2 |
| 1.5 | COM backend `move()` with round-trip verify (Windows parity) | `outlook_tool.py` (COM backend) | 60 min | 1.2 |
| 1.6 | `OutlookClient.delete(message_id, permanent=False, yes_really_permanent=False) → OperationResult` — soft-delete default; hard-delete requires double-flag (Tim F6) | `outlook_tool.py` all backends | 60 min | 1.1 |
| 1.7 | `OutlookClient.mark_read(message_id, read=True) → OperationResult` across all three backends | `outlook_tool.py` all backends | 45 min | 1.1 |
| 1.8 | CLI verbs: `outlook-tool move/delete/mark-read` — dry-run default, `--apply` explicit opt-in (Fran F5) | `outlook_tool.py` CLI | 60 min | 1.2, 1.6, 1.7 |
| 1.9 | Unit + integration tests: all 3 backends × 3 ops. Graph backend mocked with `respx`. AppleScript tested against a disposable mailbox folder on Chris's machine (not in CI). | `tests/test_outlook_tool.py`, new test files | 120 min | 1.2, 1.6, 1.7 |
| 1.10 | CHANGELOG entry + README section documenting new verbs + `yes_really_permanent` safety flag | `CHANGELOG.md`, `README.md` | 30 min | all of 1.* |

**Phase 1 total: ~8-10 hours**

**Acceptance criteria:**

- `OperationResult` has the exact shape from the panel report § Rob
- All three backends return a result with `status`, `verified`, `before_state`, `after_state`, `backend`
- Round-trip verify fires on every write on COM + AppleScript; result's `verified` field is `True` only when the post-state confirms the op
- Mac: running `outlook-tool move <msg_id> Archive` creates the Archive folder if missing, then moves, then verifies
- Running `outlook-tool delete <msg_id> --permanent --apply` is REFUSED without `--yes-really-permanent`
- Running `outlook-tool move <msg_id> Archive` (no `--apply`) prints "DRY RUN: would move message <id> to folder Archive" and exits 0 WITHOUT touching the mailbox
- Graph backend with current read-only token returns `OperationResult(status="unsupported", error_detail="OAuth grant lacks Mail.ReadWrite — use COM/AppleScript or request scope upgrade")` for every write call

### Phase 2 — `/inbox-triage` skill

**Deliverable:** Claude Code skill + two scripts (`triage.py` + `apply_triage.py`) + rule base. Workflow: classify to a JSON report, review markdown preview, apply with `--apply` + `--bucket`.

**Tasks:**

| # | Task | File(s) | Effort | Depends on |
|---|---|---|---|---|
| 2.1 | Create `.claude/skills/inbox-triage/SKILL.md` describing the two-step workflow + guardrails | `.claude/skills/inbox-triage/SKILL.md` (NEW) | 30 min | Phase 1 complete |
| 2.2 | `scripts/triage.py` — reads inbox via `outlook_tool.search`, applies rules from `scripts/triage_rules.py`, emits `output/triage-YYYYMMDD-HHMM.json` | `scripts/triage.py` (NEW) | 180 min | 1.* complete |
| 2.3 | `scripts/triage_rules.py` — initial rule base with 5 buckets (Fran). Buckets: `safe-delete` / `resolved-ticket` / `low-signal-client` / `needs-action` / `keep` | `scripts/triage_rules.py` (NEW) | 90 min | — (can parallelize with 2.2) |
| 2.4 | `scripts/apply_triage.py` — reads a triage report, prints markdown preview, applies writes when `--apply` passed. Supports `--bucket safe-delete\|resolved-ticket\|low-signal-client`. Refuses reports > 1 hour old (Tim F11). | `scripts/apply_triage.py` (NEW) | 120 min | 2.2, 1.* |
| 2.5 | Rule provenance in classifier (Rob F9) — every classified message carries a `matched_rule` field pointing at the rule id that flagged it | `scripts/triage.py`, `scripts/triage_rules.py` | 45 min | 2.2, 2.3 |
| 2.6 | Live validation run on Chris's real inbox: produce a triage report, review markdown, spot-check the safe-delete bucket with Chris, confirm 0 false positives on the sample before shipping | — | 60 min | 2.1-2.5 complete |
| 2.7 | Operator manual at `docs/manuals/inbox-triage.md` — the two-command workflow, how to tune rules, how to interpret a report | `docs/manuals/inbox-triage.md` (NEW) | 45 min | 2.1-2.5 |

**Phase 2 total: ~9 hours**

**Acceptance criteria:**

- `outlook-tool triage --since "last 30 days" --output triage-20261116-1500.json` produces a JSON report with 5 buckets + rule provenance per message
- The JSON has schema: `{captured_at, inbox_size, buckets: {safe-delete: [...], resolved-ticket: [...], low-signal-client: [...], needs-action: [...], keep: [...]}}` with each message = `{message_id, subject, sender, received_date, matched_rule, confidence}`
- `outlook-tool apply --triage-report triage-20261116-1500.json --bucket safe-delete` (no `--apply`) prints a markdown preview + exits 0 without touching the mailbox
- `outlook-tool apply --triage-report <2-hours-old.json>` REFUSES with error "triage report expired (>1h old); re-run `outlook-tool triage`"
- Live validation: of the safe-delete bucket's first 20 messages on Chris's real inbox, Chris confirms at least 18 are correctly classified before shipping

### Phase 3 — Batching (deferred)

**Trigger to revisit:** Phase 1 + 2 in live use for ≥ 1 week AND (a user complaint about wall-clock OR Graph `Mail.ReadWrite` scope granted).

**Tasks (not yet broken down to acceptance criteria):**

| # | Task | File(s) | Effort |
|---|---|---|---|
| 3.1 | AppleScript bulk-move: single `move messages {id1, id2, ...} to folder "Archive"` per target folder. Verify after. | `outlook_tool.py` (AppleScript backend) | 90 min |
| 3.2 | Graph `$batch` POST up to 20 ops per call. Honor `Retry-After` on 429. | `outlook_tool.py` (Graph backend) | 90 min |
| 3.3 | Benchmark 500-message archive before + after Phase 3 | `scripts/bench_triage.py` (NEW) | 30 min |

**Phase 3 total: ~3-4 hours.**

## Risks + mitigations

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| AppleScript silent-success — write "succeeds" but Outlook refused | HIGH | HIGH (wrong-message blast) | Round-trip verify on every write. `verified: True` required before counting as success. |
| Mac Archive folder doesn't exist | MEDIUM | MEDIUM (op fails) | Detect-and-create on first use (Al F3). Refuse with clear error if creation fails. |
| Classifier misclassifies, Chris archives a real email | MEDIUM | HIGH | Dry-run default. Markdown preview mandatory. Live validation on real inbox before shipping. Rule provenance for every classification. Bucket ordering surfaces most-confident first; Chris reviews top-down. |
| Graph backend silently fails under throttle | LOW (v1; Graph write path returns `unsupported`) | LOW | Graph is `unsupported` until scope is granted. When scope arrives, Phase 3 `$batch` with `Retry-After` honor. |
| Phase 1 ships, Phase 2 never lands, and Chris has dangerous verbs with no safe workflow around them | MEDIUM | MEDIUM | CLI verbs are dry-run default. `--apply` is required for every real write. The naked verbs are safe even without the skill wrapping them. |
| Triage report gets stale between classify and apply | LOW | MEDIUM (wrong msg_id, op fails harmlessly) | Executor rejects reports > 1 hour old (Tim F11). |
| Rule base grows unmanageable | LOW (v1) | LOW | Rules are a single Python file. When count > 40, panel re-run to decide ML classifier upgrade. |

## Open questions (none blocking)

None. Panel-approved on all design axes.

## Companion docs

- Panel report: [`docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`](../../key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md)
- Will update when shipped: `CHANGELOG.md`, `README.md`, `docs/manuals/inbox-triage.md` (new)
