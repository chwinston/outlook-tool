# Bolt 7 — JXA Recipient Fix + Chunked Scan: SHIPPED

**Date:** 2026-11-19
**Branch:** `bolt/20261119-bolt7-jxa-recipients`
**Mode:** Autonomous end-to-end
**Duration:** ~45 min (spec → agent → 2 live applies)
**Outcome:** 309/309 verified-success. Unblocked behavioral rules permanently.

## What shipped

**B7.1 — JXA recipient extraction fix (commit `335289b`)**
- Root cause: JXA `recipients[k].address()` returned empty strings
- Fix: use `recipients[k].emailAddress().address` instead — the `emailAddress()` call returns an object `{name, address, type}`
- Agent probed with `JSON.stringify(Object.keys(recipients[k]))` to find the real accessor
- Live-verified across 10 sampled messages

**B7.2 — Chunked JXA scan (commit `93f88a3`)**
- Triggers when `max_results > 500` OR window `> 90` days
- Splits into 90-day windows, walks backward from `end_date`
- Dedup by `(id, received_date)`
- Transparent to callers — no API change
- 365-day window → 5 JXA calls; previously timed out at 180s

**B7.3 — Integration test (commit `1501cc1`)**
- `@pytest.mark.integration` test against live Outlook, scans Sent Items, asserts at least one `to` field non-empty
- Deselected in default CI

## Live validation

| Phase | Classified | Sent context | Safe-delete | Resolved-ticket | Verified |
|---|---|---|---|---|---|
| Classify | 1,317 msgs | **3,546 sent / 401 unique recipients** | 149 | 160 | — |
| Delete apply | — | — | 149 | — | **149/149** |
| Archive apply | — | — | — | 160 | **160/160** |

### Rule distribution (post-fix)

**safe-delete (149 total):**
- `ignored-sender-cold`: **80** (previously catastrophically misfiring; now firing cleanly on CSI Pay reports, Vertice promo, GRC trainings, CloudMailin, Vive registrations)
- `bulk-marketing-tail`: 21
- `list-unsubscribe-bulk`: 19
- `generic-noreply-old`: 10
- `unread-automation-30d`: 10
- `cold-outreach-dormant`: 5
- Other: 4

**resolved-ticket (160 total):**
- `dead-thread-reference`: **155** (this rule was effectively dead in Bolt 6; now firing correctly on Re: chains Chris replied to and nothing came back — eFit case 78475, Rep1 Nuvei signer confirmations, Pyxis follow-ups)
- Other: 5

### Spot-check validation

**ignored-sender-cold sample (15 shown, all real cold):**
- ✅ `merchant@csipay.com` transaction report (automated)
- ✅ `philipgrave@vertice.one` Vertice vendor pitch
- ✅ `webtrainings@insights.grciq.com` training promo
- ✅ `invoice+statements@mail.anthropic.com` receipts
- ✅ `vive@events.hlth.com` ViVE 2027 registration
- ✅ No real customers (Beda BJJ, Rep1, Constellation, EZFacility) in the list

**dead-thread-reference sample (10 shown, all legitimate archives):**
- ✅ `Tsnyder@efitfinancial.com — Re: Case 78475` (resolved case)
- ✅ `eyvon@rep1.ca — Re: Nuvei signer` (4 copies, resolved doc signing)
- ✅ `chenao@csipay.com — Re: Rep 1 follow-up` (resolved thread)

## What went well

- **Agent found the right JXA accessor in one shot** by writing a diagnostic probe instead of guessing. Classic "look at the data before you assume."
- **Chunking is clean** — 5 JXA calls, dedup via `(id, received_date)`, early-exit on `max_results`. No overhead for small windows.
- **sent-context went from 0 → 3,546 scanned / 401 unique recipients.** That's the signal that unblocked everything else.
- **263 tests pass** (+8 unit, +1 integration vs Bolt 6 baseline of 255).

## What hurt

- Nothing significant. Clean sprint.

## Session cumulative

| Milestone | Inbox count | Cleared this phase |
|---|---|---|
| Start of session 2 days ago | 2,095 (all-time) | — |
| End Bolt 4 | 48 (30d view) | 218 |
| Bolt 5 big sweep | 1,415 | 692 |
| Bolt 6 | 1,313 | 141 |
| **Bolt 7 (this)** | **~1,004** | **309** |
| **Session total cleared** | | **1,360 messages** |

## Follow-ups parked

- **LLM pass still untested end-to-end** — needs `ANTHROPIC_API_KEY` in a future session. Infra ready since Bolt 6.
- **Archive folder cleanup** — Archive still has 10,192 messages. Bolt 8 candidate (apply same rules to Archive, demote to Deleted Items what's obvious noise).
- **COM backend parity for chunked scan** — only AppleScript backend got the fix. Windows path unchanged.

## Links

- Design doc: `docs/superpowers/specs/2026-11-19-bolt7-jxa-recipient-fix-design.md`
- Apply result (delete): `output/bolt7-apply-delete.md`
- Apply result (archive): `output/bolt7-apply-archive.md`
- Fresh triage: `output/triage-20261119-0713.json`
