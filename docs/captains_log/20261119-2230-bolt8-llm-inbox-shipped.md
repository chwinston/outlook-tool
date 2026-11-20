# Bolt 8 — In-Session LLM Classification: SHIPPED

**Date:** 2026-11-19
**Branch:** `bolt/20261119-bolt8-llm-inbox`
**Mode:** 1 inline batch + 4 parallel sub-agents, all judging in-session
**Duration:** ~45 min (extract → classify → verify → apply)
**Outcome:** 189/189 verified-success. Inbox cleared another chunk of real-work residue.

## What shipped

**Two new scripts** that make the inbox LLM-ready without touching the Anthropic API:

- `scripts/extract_keep_batch.py` — slices N messages from a triage report's `keep` bucket into a slim batch JSON (sender + subject + body_preview + is_read + received_at)
- `scripts/apply_llm_verdicts.py` — merges `{id, bucket, confidence, reason}` verdicts back onto the triage report as `llm_verdict` fields

Plus:
- `scripts/find_clusters.py` — dedupes Inbox by (sender, normalized_subject) to find fat clusters (useful for one-shot sweeps)
- `scripts/inbox_composition.py` — age/unread/sender-concentration profile of Inbox

Pipeline:

```
triage.py classifies            → 489 msgs land in `keep` bucket (rules couldn't bucket them)
extract_keep_batch.py           → 5 batch JSONs (100+100+100+100+89)
me OR sub-agents judge          → 5 verdicts JSONs
apply_llm_verdicts.py           → verdicts merged onto triage report
Python 1-liner re-routes        → bucket flipped to safe-delete / resolved-ticket based on llm_verdict
apply_triage.py                 → bulk-move to Deleted Items / Archive
```

## Classification shape

Judge was given customer/colleague/vendor context (names of real Member Solutions clients like Jungle MMA, Baltimore Martial Arts, Infofit, Rep1, Pyxis vendors, etc.) and asked to bucket each message into:

- **junk**: cold outreach, expired 2FA, newsletters, receipts
- **reference**: completed threads, OOO replies, resolved tickets
- **action**: direct asks from real humans
- **unclear**: not enough signal

Confidence 0.85+ = eligible for apply. Below = routed to needs-action for operator review.

## Results per batch

| Batch | Size | junk | reference | action | unclear | High-conf junk | High-conf ref |
|---|---|---|---|---|---|---|---|
| 1 (inline) | 100 | 22 | 41 | 37 | 0 | 22 | 6 |
| 2 (sub-agent) | 100 | 27 | 28 | 45 | 0 | 27 | 14 |
| 3 (sub-agent) | 100 | 30 | 25 | 45 | 0 | 30 | 14 |
| 4 (sub-agent) | 100 | 9 | 56 | 35 | 0 | 9 | 41 |
| 5 (sub-agent) | 89 | 29 | 23 | 35 | 2 | 29 | 6 |
| **Total** | **489** | **117** | **173** | **197** | **2** | **117** | **81** |

High-confidence total eligible for apply: **198** (117 delete + 81 archive).

## Live application

Delete (118 incl. 1 rule-matched): **118/118 verified-success**
Archive (71 after re-matching to fresh IDs, 10 lost in the gap between delete and archive apply): **71/71 verified-success**

**189 total moves. Zero failures. Zero UNVERIFIED.**

## What went well

- **Parallel sub-agents cut classification time ~4×.** Each handled 100 messages in ~80 seconds. Sequential in-session would have been ~5-6 minutes AND context-pressured me past the batch-3 mark.
- **Context-dropping strategy worked.** Sub-agents got fresh contexts, each ~75k tokens, well under limit. My main context stayed lean.
- **Operator-judgment proxy prompt was precise enough.** Agents correctly identified real customers (Beda BJJ PCI refund, Rep1 rate escalation, Madison Wellness cancel threat) as action, caught vendor noise + cold outreach as junk, and routed resolved-thread context to reference.
- **Human review caught exactly the right concerns.** Pyxis legal contract threads flagged as borderline — Chris approved archive anyway (recoverable).

## What hurt

- **Reference verdicts lost 11/81 (14%) in the ID-shift gap** between delete-apply and archive-apply. The `(sender, subject, received_at)` fingerprint matched most, but a few messages dropped. Not a reliability issue — just efficiency.
- **Fresh triage runs reported `sent-context: 0`** despite Bolt 7 fix. Needs follow-up: the `build_sent_context()` helper likely has a bug in the triage.py wiring that only surfaces on repeat runs, OR Outlook disconnects the Sent folder between runs.
- **One sub-agent misreported batch counts** in its summary (said "88 at >=0.85" when actual high-conf was 90). Didn't affect verdicts — the JSON output was correct.

## Sub-agent prompt lessons

Worked:
- Priming with specific real customer names
- "Lean generous with action" instruction prevented overclassification to junk
- Explicit confidence scale (0.9+, 0.7-0.85, below 0.7)
- Fresh contexts prevented cross-batch contamination

Didn't matter:
- Asking agents to flag "surprises" — their observations were nice but didn't change verdicts

Would improve next time:
- Pass the full triage report so agents can see context on prior conversations (currently they see only the single batch message)
- Have agents flag items that reference known-settled matters so Chris can archive instead of delete

## Metrics

| Metric | Value |
|---|---|
| Keep-bucket messages | 489 |
| Messages classified | 489 (100%) |
| High-confidence verdicts | 198 of 489 (41%) |
| Applied moves | 189 |
| Verified-success | 189/189 (100%) |
| Failed | 0 |
| API cost | $0 (session subscription tokens only) |
| Wall clock | ~45 min end-to-end |

## Session cumulative totals

| Milestone | Inbox count |
|---|---|
| Start yesterday (2026-11-16) | 2,095 (all-time) |
| End Bolt 4 | — |
| End Bolt 5 big sweep | 1,415 |
| End Bolt 6 | 1,313 |
| End Bolt 7 | 1,010 |
| **End Bolt 8 (this)** | **~820** (projected after archive sync) |
| **Cumulative cleared** | **~1,275 messages** |

## Follow-ups

- **SOC-awareness note added** to `claude-instructions.md` (PR #6 merged separately) after Novawatch case 19014 fired on Bolt 7's JXA diagnostic probe. Future agents must warn before writing scripts to `/tmp` and running via `osascript -l JavaScript`.
- **sent-context regression** — build_sent_context returning 0 recipients on today's second run. Was working this morning. Investigate before next run.
- **Remaining ~820 Inbox** is overwhelmingly real work (colleagues + active customer/vendor threads). Further LLM passes will yield diminishing returns. True Inbox Zero via Archive remains an option.

## Links

- Design: implicit (small-change follow-up to Bolt 6 LLM infra)
- Apply result (delete): `output/bolt8-apply-delete.md`
- Apply result (archive): `output/bolt8-apply-archive.md`
- Verdict files: `output/llm-verdicts-{001..005}.json`
- SOC-awareness doc: PR #6 merged to main
