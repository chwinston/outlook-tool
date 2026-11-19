# Bolt 7 — JXA Recipient Extraction + Chunked Classifier Scan

**Date:** 2026-11-19
**Status:** Approved by Chris 2026-11-19 ("Keep crushing")
**Size:** Small change — scoped bugfix + one perf improvement

## Intent

Two Bolt 6 follow-ups ship together:

1. **Fix JXA `toRecipients()` returning empty strings** — broke `build_sent_context()` and gated out `ignored-sender-cold` + `dead-thread-reference` rules during Bolt 6 live validation
2. **Chunked classifier scan** — full-Inbox classification times out at 180s due to body_preview volume

Both block "clean the remaining 1,313 Inbox" from working. Both are small, scoped fixes.

## Scope

### B7.1 — JXA recipient extraction

Current implementation at `outlook_tool.py` ~line 429 (`_search_applescript` JXA block):
```javascript
var recipients = msg.toRecipients();
var addrs = [];
for (var k = 0; k < recipients.length; k++) {
    try { addrs.push(recipients[k].address()); } catch(e) {}
}
toStr = addrs.join(", ");
```

Bug: `recipients[k].address()` returns empty strings even when the recipient exists. Likely a JXA accessor issue — needs `address` not `address()`, or a different property name.

Fix approach (in order of preference):
1. Try `recipients[k].address` (no parentheses — JXA property access)
2. Try `recipients[k].emailAddress()` or `recipients[k].emailAddress`
3. Try `recipients[k].name()` + address parsing
4. Fall back to raw header parsing from `msg.content()` or similar

Add an integration test (opt-in, `pytest -m integration`) that seeds a message in `_bolt4_test`, scans Sent Items, asserts non-empty `to` field.

### B7.2 — Chunked classifier scan

Current: `OutlookClient.search(folder='Inbox', date_from='2020-01-01', date_to='2026-11-19', max_results=5000)` tries to pull everything in one JXA call → timeout.

Fix: add internal chunking in `_search_applescript`. If `max_results > 500` OR `(date_to - date_from).days > 90`, split the date range into 90-day windows, call JXA per chunk, concatenate results.

Transparent to callers — no API change.

### B7.3 — Re-run behavioral rules on full Inbox

With B7.1 + B7.2 shipped:
1. Fresh classify of 1,313 Inbox with working sent-context
2. `ignored-sender-cold` + `dead-thread-reference` should now fire WITHOUT false-positiving on real customers
3. HARD PAUSE for Chris to spot-check
4. Apply the clean safe-delete bucket

## Non-goals

- COM backend parity for JXA fix (only affects AppleScript backend)
- Graph backend chunking (Graph already pages)
- New rules or classifier changes
- LLM pass (still parked until API key)

## Approaches considered

**A. Fix JXA now + classifier chunking + re-run** (chosen)
- Scoped, unblocks the ignored-sender-cold tail cleanup, ~2 hour agent build

**B. Rewrite the whole JXA scan to use raw EWS/Graph instead** 
- Over-scoped. Rejected.

**C. Skip the fix, keep excluding ignored-sender-cold indefinitely**
- Leaves 100+ msgs per classifier run misfiring. Rejected — this is a reliability bug.

## How

1. **/gitsmith setup** — branch `bolt/20261119-bolt7-jxa-recipients` off main (done)
2. **Dispatch implementer agent** — JXA fix + chunked scan + tests
3. **Verify tests pass** (target 255 → 258+)
4. **Live validation** — classify 1,313 Inbox, HARD PAUSE, apply clean safe-delete bucket
5. **/pm close + PR + merge**

## Success criteria

- [ ] `sent-context: N sent messages scanned` where N > 0 in live classifier output
- [ ] `ignored-sender-cold` fires on actual cold senders only (spot-check 20, find ≤ 1 false positive)
- [ ] Full-Inbox classification (1,313 msgs) completes without timeout
- [ ] 300+ additional messages cleared from Inbox
- [ ] All 255 existing tests + new tests pass

## Approved

Chris 2026-11-19 ("Keep crushing"). Autonomous execution authorized.
