# Backlog

**Last groomed:** 2026-11-19 (Bolt 8 In-Session LLM Classification SHIPPED. 489 messages from `keep` bucket classified via 1 inline + 4 parallel sub-agents, ~56k tokens. Verdicts: 117 junk + 81 reference + 197 action + 2 unclear. Applied: 118 delete + 71 archive = 189 verified-success. Captain's log: `docs/captains_log/20261119-2230-bolt8-llm-inbox-shipped.md`. Also fixed SOC-awareness in claude-instructions.md after Novawatch case 19014.) Prior: 2026-11-19 (Bolt 7 JXA Recipient Fix + Chunked Scan SHIPPED. Unblocked behavioral rules — ignored-sender-cold + dead-thread-reference now fire SAFELY with working sent-context (3,546 sent msgs / 401 unique recipients scanned). Live validation: 149 safe-delete + 160 archive = 309 messages moved. Captain's log: `docs/captains_log/20261119-0720-bolt7-jxa-fix-shipped.md`.) Prior: 2026-11-19 (Bolt 6 Advanced Triage SHIPPED. 5 behavioral rules + List-Unsubscribe rule + LLM classifier + marketplace scan. Live-run: 141/141 verified-success, Inbox 1,425 → 1,313. Dangerous ignored-sender-cold bug found + fixed mid-run — empty sent-scan now treated as broken scan not "never replied". Captain's log: `docs/captains_log/20261119-0620-bolt6-advanced-triage-shipped.md`.) Prior: 2026-11-17 (Opened Bolt 5 — Phase 3 batching trigger met early. Observed 44% silent-success failure rate on 43-msg apply run, far exceeding acceptable reliability. B4.3.* graduating to B5.*. Spec: `docs/pm/plans/20261117-0740-bolt5-batching-plan.md`.) Prior: 2026-11-16 (Bolt 4 Phase 1 + Phase 2 SHIPPED via autonomous execution. 134 real-world soft-deletes on Chris's Inbox, 20/20 ship-gate hit, 1 verify bug found + fixed mid-run, 4 exemption rules added from walk-through. Captain's log: `docs/captains_log/20261116-2230-bolt4-inbox-triage-shipped.md`.) Prior: 2026-03-27 (Bolt 3 close).

## Completed

| # | Item | Bolt |
|---|------|------|
| 1-8 | Security hardening (F1-F10), CI pipeline | Bolt 1 |
| 9 | Create SECURITY.md | Bolt 2 |
| 10 | Create REQUIREMENTS.md with REQ IDs | Bolt 2 |
| 11 | Calendar event search (all 3 backends) | Bolt 2 |
| 12 | Write verbs: move / delete / mark-read across 3 backends with OperationResult contract | Bolt 4 Phase 1 |
| 13 | `/inbox-triage` skill: classifier + executor + 23 rules + operator manual | Bolt 4 Phase 2 |
| 14 | Live validation: 134 real soft-deletes on Chris's Inbox (266→132) | Bolt 4 Phase 2 |
| 15 | Bolt 5: bulk-move primitive + silent-success fix. 6/6 verified on live sweep | Bolt 5 |
| 16 | Bolt 6: 5 behavioral rules (ignored-sender, dead-thread, cold-outreach, bulk-marketing-tail, unread-automation) | Bolt 6 Phase A1 |
| 17 | Bolt 6: Claude Haiku LLM classifier module (29 mocked tests) | Bolt 6 Phase A2 |
| 18 | Bolt 6: Marketplace scan — zero Outlook plugins exist in Anthropic marketplace | Bolt 6 Phase A3 |
| 19 | Bolt 6: List-Unsubscribe body rule + unread-junk gate + --llm wiring | Bolt 6 Phase B |
| 20 | Bolt 6 live validation: 141/141 verified-success soft-deletes (1,425→1,313) | Bolt 6 Phase C |
| 21 | Bolt 7: JXA recipient extraction fix (`emailAddress().address` property access) | Bolt 7 B7.1 |
| 22 | Bolt 7: Chunked classifier scan (90-day windows, transparent to callers) | Bolt 7 B7.2 |
| 23 | Bolt 7 live validation: 149 safe-delete + 160 resolved-ticket archive (309/309 verified) | Bolt 7 B7.3 |
| 24 | Bolt 8: In-session LLM batching helpers (extract_keep_batch.py, apply_llm_verdicts.py) | Bolt 8 |
| 25 | Bolt 8: 489 keep-bucket msgs classified via 5 batches (1 inline + 4 parallel sub-agents) | Bolt 8 |
| 26 | Bolt 8 live validation: 118 delete + 71 archive (189/189 verified-success) | Bolt 8 |
| 27 | SOC-awareness note added to claude-instructions.md after Novawatch case 19014 | ops |

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
| B4.1.1 | Add `OperationResult` dataclass to `outlook_tool.py` (shared contract across backends) | XS (15 min) | ✅ done | F4 (Rob F1) | `{status, verified, message_id, operation, before_state, after_state, error_detail, backend, elapsed_ms}` |
| B4.1.2 | Add `OutlookClient.move(message_id, target_folder) → OperationResult` dispatch method | XS (15 min) | ✅ done | — | Thin dispatch to backend-specific impl |
| B4.1.3 | `_AppleScriptBackend.move()` with Archive folder detect-and-create + round-trip verify | S (90 min) | ✅ done | F1 (Tim) + F3 (Al) | Mac-primary path; silent-success is the dominant risk; verify catches it |
| B4.1.4 | `_GraphBackend.move()` returns `status="unsupported"` without `Mail.ReadWrite` (don't auto-upgrade) | S (45 min) | ✅ done | F2 (Rob) | Surface the scope cliff loudly |
| B4.1.5 | COM backend `move()` with round-trip verify (Windows parity) | S (60 min) | ✅ done | F1 (Tim) | Windows path |
| B4.1.6 | `OutlookClient.delete(message_id, permanent=False, yes_really_permanent=False)` — soft-delete default, hard-delete requires double-flag | S (60 min) | ✅ done | F6 (Tim + unanimous) | Dangerous op; guarded by design |
| B4.1.7 | `OutlookClient.mark_read(message_id, read=True)` across all three backends | S (45 min) | ✅ done | — | Lowest-risk write; rounds out surface |
| B4.1.8 | CLI verbs `outlook-tool move/delete/mark-read` — dry-run default, `--apply` required for real write | S (60 min) | ✅ done | F5 (Panel unanimous) | Safe-by-default CLI shape |
| B4.1.9 | Unit + integration tests: 3 backends × 3 ops. Graph mocked with `respx`. AppleScript tested against disposable mailbox folder locally (not in CI). | M (120 min) | ✅ done | — | Correctness guard |
| B4.1.10 | CHANGELOG entry + README section documenting new verbs + `yes_really_permanent` safety | S (30 min) | ✅ done | — | Operator docs |

### Phase 2 — `/inbox-triage` skill (~9 hours total, blocked on Phase 1 complete)

| # | Item | Size | Status | Panel F-tag | Notes |
|---|------|------|--------|-------------|-------|
| B4.2.1 | Create `.claude/skills/inbox-triage/SKILL.md` | S (30 min) | ✅ done | — | Two-step workflow: classify → apply |
| B4.2.2 | `scripts/triage.py` — classifier reading inbox via `outlook_tool.search`, emitting `output/triage-YYYYMMDD-HHMM.json` | M (180 min) | ✅ done | F10 (Fran) | Classifier is its own concern |
| B4.2.3 | `scripts/triage_rules.py` — initial rule base: 5 buckets + 4 exemption rules added from live walk-through | S (90 min) | ✅ done | F10 (Fran) | 23 rules total (19 initial + 4 exemption) |
| B4.2.4 | `scripts/apply_triage.py` — reads report, prints markdown preview, applies when `--apply` passed. `--exclude-ids` added for walk-through overrides. | M (120 min) | ✅ done | F10 + F11 | Executor is separate from classifier |
| B4.2.5 | Add rule provenance (`matched_rule`) to classifier output | XS (45 min) | ✅ done | F9 (Rob) | Debuggability — tune the rule, not retrain a model |
| B4.2.6 | Live validation: 134 real-world soft-deletes on Chris's Inbox (266→132). Ship-gate: 20/20 correct on spot-check. | S (60 min) | ✅ done | — | SHIPPED 2026-11-16 |
| B4.2.7 | Operator manual `docs/manuals/inbox-triage.md` | S (45 min) | ✅ done | — | Two-command workflow docs |

### Phase 3 — Batching (GRADUATED to Bolt 5 on 2026-11-17 — trigger met early)

Original B4.3.* rows have moved to Bolt 5 below. Observed 44% silent-success failure rate on round-4 apply run made this reliability work, not throughput work.

## Bolt 5 — Silent-Success Fix + Batching (NEW 2026-11-17)

Spec: `docs/pm/plans/20261117-0740-bolt5-batching-plan.md`

**Trigger evidence:** 4 apply-triage runs on real Inbox showed silent-success rate climbing from 30% → 44%. The deferred trigger for B4.3.* ("500+ msg complaint OR 1 week of use") is a proxy for reliability pain; that pain is already here at 43-msg sizes.

| # | Item | Size | Status | Notes |
|---|------|------|--------|-------|
| B5.1 | Root-cause AppleScript silent-success (index-shift vs sync-lag vs verify-side-effect hypotheses) | M (3-4h) | executable | BLOCKING for B5.2 — pick a shape based on root cause |
| B5.2 | AppleScript bulk-move primitive (shape A/B/C per B5.1) — single tell block capturing refs before any move | M (2-3h) | blocked on B5.1 | Was B4.3.1 |
| B5.3 | Verify-step hardening — fix phantom UNVERIFIED-success (scan-limit-200 miss under mutation) | S (90 min) | executable | Addresses apply_triage.py reporting accuracy |
| B5.4 | Graph `$batch` POST | S (90 min) | still deferred (needs IT ticket) | Was B4.3.2 — blocked on `Mail.ReadWrite` scope |
| B5.5 | Benchmark 500-msg archive — current vs bulk-move | S (60 min) | blocked on B5.2 | Was B4.3.3; ship-gate ≥ 95% verified-success |

### Deferred from Bolt 4 (with named triggers)

| Item | Source | Trigger to revisit |
|---|---|---|
| ML classifier for sender/subject clustering | Panel F12 (not-graduating) | Rule count > 40 OR Chris's bucket-override rate > 10% of classifications |
| Recurring scheduled triage via cron | Panel F13 (not-graduating) | After 4 consecutive weeks of manual use with < 5% misclassifications |
| Cross-mailbox delegation (triage a shared mailbox) | Deferred in plan § out-of-scope | First time Chris asks to triage someone else's mailbox |
| Graph `Mail.ReadWrite` scope acquisition (IT ticket) | Deferred in plan § out-of-scope | First time Chris needs to triage from a machine where Outlook isn't running |
| Hard-delete batch mode | Panel ruled explicit NO-BUILD | Never — foot-gun by design; `--permanent --yes-really-permanent` is per-message on purpose |
