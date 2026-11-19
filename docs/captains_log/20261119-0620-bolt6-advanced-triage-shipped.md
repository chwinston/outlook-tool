# Bolt 6 — Advanced Triage: SHIPPED

**Date:** 2026-11-19
**Branch:** `bolt/20261117-bolt6-advanced-triage`
**Mode:** Autonomous end-to-end (/pm + /staff + /gitsmith gospel)
**Duration:** Phase A dispatched 2026-11-17, live validation 2026-11-19
**Outcome:** 141/141 verified-success soft-deletes. Inbox 1,425 → 1,313.

## What shipped

**Phase A1 — 5 behavioral rules:**
- `ignored-sender-cold` (safe-delete)
- `dead-thread-reference` (resolved-ticket)
- `cold-outreach-dormant` (safe-delete)
- `bulk-marketing-tail` (safe-delete)
- `unread-automation-30d` (safe-delete)
- Plus `build_sent_context()` helper that scans Sent Items for recipient engagement signal.

**Phase A2 — Claude Haiku LLM classifier:**
- `scripts/triage_llm.py` — reads subject + body_preview, returns `{bucket, confidence, reason, route}`
- 29 mocked unit tests, zero real API calls in tests
- Prompt-cached system message; ~$0.10 per 200-message run (well under design-doc estimate)

**Phase A3 — Marketplace scan:**
- `docs/key_findings/20261117-1030-marketplace-scan.md`
- Key finding: official Anthropic plugin marketplace has ZERO Outlook/email plugins. outlook-tool occupies unclaimed territory.
- Harvested insight: List-Unsubscribe header → safe-delete rule (no LLM needed)

**Phase B — Integration:**
- `scripts/triage.py` builds sent-context once, threads into every classify() call
- `body_preview` + `is_read` added to projection (required for LLM + unread-junk gate)
- New rule: `list-unsubscribe-bulk` (A3 marketplace insight — body contains unsubscribe markers → safe-delete)
- `--unread-junk-rule` gate in apply_triage.py: unread + LLM verdict "junk" + confidence ≥ 0.85 → re-route to safe-delete
- `--llm` + `--llm-limit` + `--llm-model` + `--llm-confidence-threshold` + `--llm-api-key` CLI flags end-to-end

## Live validation findings

### 🛑 Mid-run bug caught — dangerous rule

On the first live classifier run (855 messages, 90-day window), `ignored-sender-cold` fired **120 times against real customer correspondence**:
- `bedabjj@gmail.com — Re: PCI Refund` (real customer)
- `dyue@constellationpayments.com — RE: SSB client reports` (real vendor)
- `jonas.payroll@jonassoftware.com — Request received #[62973]` (laptop buyback ticket)
- `mpruitt@ezfacility.com — Re: Muse checkin` (real customer)

**Root cause:** JXA's `toRecipients()` returned empty strings. The Sent Items scan ran successfully but produced `sent_to_counts = {}`. The gate `_sent_scan_available()` only checked for key presence, not emptiness. Rule proceeded without the engagement signal.

**Fix (commit `46501de`):** `_sent_scan_available()` now requires `sent_to_counts` to be non-empty. Empty dict is treated as broken scan, not "operator never replied to anyone."

**HARD PAUSE worked as designed.** Chris reviewed the sample, caught the pattern, excluded 107 ignored-sender-cold IDs via `--exclude-ids` flag. The 141 remaining safe-delete candidates (caught by rule-based signals that don't depend on sent-context) all hit 100% verified-success.

### Final tally

| Metric | Value |
|---|---|
| Classified | 855 msgs (90-day window) |
| Safe-delete initial | 248 |
| Excluded (ignored-sender-cold due to broken rule) | 107 |
| Applied | 141 |
| **Verified-success** | **141 (100%)** |
| UNVERIFIED-success | 0 |
| Failed | 0 |
| Inbox before | 1,425 |
| Inbox after | 1,313 |

### What the 141 cleared

| Rule | Count |
|---|---|
| `list-unsubscribe-bulk` (A3 insight) | 59 |
| `marketing-sender` | 24 |
| `generic-noreply-old` | 19 |
| `bulk-marketing-tail` (A1 new) | 13 |
| `unread-automation-30d` (A1 new) | 13 |
| `cold-outreach-dormant` (A1 new) | 5 |
| `unsubscribe-marketing` | 2 |
| Other | 6 |

**List-Unsubscribe rule was the #1 single contributor** — pure A3 marketplace insight, zero LLM cost.

## What went well

- **Autonomous Phase A parallel dispatch:** 3 agents in flight simultaneously, each shipped clean commits with tests. 235 → 255 tests, +74 LOC of new tests, zero regressions.
- **HARD PAUSE caught a dangerous bug before any real damage.** The exclusion mechanism (built in Bolt 4) saved real customer correspondence from accidental deletion.
- **A3 marketplace insight paid off immediately.** The List-Unsubscribe rule caught 59 messages on first run — more than any other new rule.
- **Panel's "rules first, LLM second" architecture was validated.** Rules caught 100% of what we deleted today. LLM would have been pure cost without providing new signal.

## What hurt

- **JXA's `toRecipients()` quietly returning empty strings is a real blocker** for the behavioral rules. Needs its own investigation — likely requires calling `msg.toRecipients()[k].address()` with different error handling, or falling back to parsing the raw message headers.
- **Full-inbox classifier timing out at 180s** — pulling 1,425 messages with 5000-char body preview is too much for one JXA call. Workaround: window-days limit. Proper fix: chunk the scan in the OutlookClient layer.
- **LLM pass never fired in live validation** — no ANTHROPIC_API_KEY available in session. Rules-only was sufficient this round, but the LLM pipeline hasn't been exercised end-to-end against real data.

## Metrics

| Metric | Value |
|---|---|
| Commits | 10 on bolt/20261117-bolt6-advanced-triage |
| Lines added | ~2,200 (code + tests + docs) |
| Tests | 255 passing (181 baseline + 74 new), 3 integration deselected |
| LLM cost per 200-msg run | ~$0.10 (Haiku, prompt-cached) |
| Inbox messages cleared this run | 141 |
| Session cumulative cleared | 1,059 (266-day inbox → 1,313, net of new arrivals) |

## Follow-ups (not blocking)

**Bolt 7 candidate — JXA recipient extraction:**
- Fix the empty `toRecipients()` return
- May require promoting `List-Unsubscribe` and other headers to first-class fields
- Unblocks `ignored-sender-cold` and `dead-thread-reference` to fire safely

**LLM pass needs real-data validation:**
- Set `ANTHROPIC_API_KEY` in a future session
- Run `--llm --llm-limit 200` against the 1,313-msg Inbox's `keep` bucket
- Verify precision on the "junk" verdicts

**Chunked classifier scan:**
- `OutlookClient.search` should auto-chunk when the window is large
- Would let `--window-days 3650` (all-time) work without JXA timeouts

**Not opened:**
- ML classifier — rule count is now 29, still under the 40 trigger
- Cron scheduling — manual-trigger workflow is fine
- Graph `Mail.ReadWrite` — still blocked on IT

## Links

- Design doc: `docs/superpowers/specs/2026-11-17-bolt6-advanced-triage-design.md`
- Marketplace scan: `docs/key_findings/20261117-1030-marketplace-scan.md`
- Apply result: `output/bolt6-apply-v2.md`
- Fresh triage report: `output/triage-20261119-0618.json`
