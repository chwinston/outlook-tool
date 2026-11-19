# Bolt 6 — Advanced Triage (LLM + Patterns + Marketplace Scan)

**Date:** 2026-11-17
**Status:** Approved by Chris 2026-11-17 ("full autonomous, /pm + /staff + /gitsmith")
**Mode:** Autonomous end-to-end

## Intent & Why

Chris's Inbox is 1,425 messages after Bolt 4+5 cleared 900+. The remaining 1,425 is harder territory:
- 430 internal colleagues
- 163 residual automation
- 832 "real external people" — sales reps, vendors, prospects, customer-service threads
- 447 unread
- Sender-based bulk is topping out: next wave requires body/subject semantic judgment

Chris wants "brilliant." That's three threads in one bolt:

1. **Pattern-based detection** (level 1) — surface ignored senders, dead threads, cold outreach via measurable signals, not hand-curated lists
2. **LLM semantic classification** (level 2) — read subject + body preview, classify intent: action-required / reference / dead / junk
3. **Marketplace scan** (level 3) — find prior art: Claude Code plugins, skills, MCP servers that already solve pieces of this

Plus: **unread + junk = delete** (Chris's explicit rule — if it meets junk criteria AND Chris never opened it, it was never going to be read).

## Goals

1. Add **5 pattern-based rules** to `scripts/triage_rules.py` covering ignored-sender, dead-thread, cold-outreach-dormant, bulk-marketing-tail, unread-and-junk
2. Add an **LLM classifier** (`scripts/triage_llm.py`) that reads subject + body_preview for borderline messages and returns a structured verdict
3. Produce a **marketplace reference** doc (`docs/key_findings/20261117-marketplace-scan.md`) listing relevant Claude Code plugins, skills, and MCP servers with evaluation notes
4. Integrate into `/inbox-triage` workflow: rules → LLM fallback for "keep" bucket → human review → apply
5. Live validate on Chris's actual Inbox: target 300+ additional messages cleared, ≥ 95% precision on the LLM-classified "delete" bucket

## Non-Goals

- No gmail/other-provider support
- No account-takeover or bulk-send operations (read + classify + move/delete only)
- No training our own model — use Claude API via SDK
- No persistent LLM cache that touches message content on disk (ephemeral in-run only; privacy stays local)

## Anti-Goals (would mean failure even if "works")

- LLM mistakenly deletes a real message Chris cared about (precision > recall)
- LLM runs on every message (cost explosion) — only on classifier uncertain / `keep` bucket borderline cases
- Exposes mail content to any third-party service beyond Claude API (which Chris already uses)
- Sender-list sprawl — rules should stay < 40 total (panel's trigger for "graduate to ML")

## Constraints

- Mac-primary (AppleScript backend)
- Bolt 4 contract: `OperationResult` + dry-run default + `--apply` gate + `--exclude-ids`
- Bolt 5 contract: bulk primitives honored (LLM verdict feeds bulk-move)
- Reports > 1h old refused (panel F11)
- LLM: use existing `anthropic` SDK; no new auth plumbing (relies on `ANTHROPIC_API_KEY` env)
- Chris's new rule: **unread + classifies as junk = delete**. Read + classifies as junk = archive (he saw it, may want to reference later)

## Approaches considered

**A. Rules-only expansion** — add 20 more pattern rules
- Easy, zero LLM cost. But rules sprawl and can't read body. Rejected: already at 23 rules, panel trigger for ML is 40.

**B. Pure LLM pipeline** — send every Inbox message to Claude API, classify end-to-end
- Highest ceiling. But expensive (~$4-8 per full sweep at Haiku), slow (10-20 min for 1400), and overkill for the 500+ messages that rules already handle reliably.

**C. Hybrid (chosen):** rules → LLM for `keep` bucket fallback + unread-junk gate → bulk apply
- Rules handle the obvious (free, fast). LLM handles only what rules can't bucket confidently.
- Leverages Chris's "unread + junk = delete" insight — most unread tail is unopened cold outreach.
- Marketplace scan runs in parallel (independent research deliverable).

**Chosen: C.** Rationale: 3 levels compose. Level 1 pays off even if level 2 fails. Level 3 is pure research that informs future bolts.

## Pattern-based rules (level 1)

Add to `scripts/triage_rules.py`:

| Rule name | Signal | Bucket |
|---|---|---|
| `ignored-sender-cold` | ≥ 3 messages from sender, zero replies from Chris, oldest > 60 days | safe-delete |
| `dead-thread-reference` | Thread > 90 days old, last reply from Chris, no incoming for 60+ days | resolved-ticket (archive) |
| `cold-outreach-dormant` | Sender's domain in known-cold list (gartner.com, outreach, cold-email patterns) + no reply from Chris + > 30 days | safe-delete |
| `bulk-marketing-tail` | Sender matches marketing domain pattern + > 30 days + unread | safe-delete |
| `unread-automation-30d` | is_read=False + automated-sender pattern + > 30 days | safe-delete |

Implementation detail: "zero replies from Chris" requires scanning Sent Items for messages TO this sender. Add a `_scan_sent_to_sender(email)` helper.

## LLM classifier (level 2)

**Module:** `scripts/triage_llm.py`

**When invoked:**
- By `scripts/triage.py` for messages that classify to `keep` bucket (the default — rules couldn't bucket them)
- Opt-in via `--llm` flag (off by default; on costs money)
- Only for messages where `body_preview` + `subject` + `sender_email` are available

**Model:** `claude-haiku-4-5-20251001` (cheap, fast, good enough for binary junk/keep triage)

**Prompt shape** (inline, not a file):
```
You are an expert at classifying business email for a software engineering executive.

Classify the following message into EXACTLY ONE bucket:
- junk: cold outreach, sales pitches, newsletter, promotional, automated-noise
- reference: useful historical content, archive but don't delete
- action: contains a direct ask or action required of the recipient
- unclear: not enough signal

Return JSON only: {"bucket": "...", "confidence": 0.0-1.0, "reason": "..."}

Message:
From: <sender>
Subject: <subject>
Body preview: <first 500 chars>
```

**Cost controls:**
- Prompt caching on the system prompt (reuse across batch)
- Batch 10 messages per API call via content blocks
- Confidence < 0.7 → route to `needs-action` (operator decides)
- Confidence ≥ 0.85 + "junk" → eligible for delete
- Max messages per `--llm` run: 200 (safety cap; override with `--llm-limit`)
- Estimated cost: ~$0.50 per 200-message LLM pass at Haiku

**Output:** augments the triage.json report with `llm_verdict` field alongside `matched_rule`.

## Chris's "unread + junk = delete" rule

Implemented as a **post-classification gate** in `apply_triage.py`:

- When bucket == "safe-delete" AND is_read == False → eligible for immediate soft-delete (standard)
- When bucket == "resolved-ticket" or "low-signal-client" AND is_read == False AND LLM verdict == "junk" with confidence ≥ 0.85 → **re-route to safe-delete** (user's explicit opt-in)
- When bucket == "keep" AND is_read == False AND LLM verdict == "junk" with confidence ≥ 0.85 → **re-route to safe-delete**

Gate is on by default; `--no-unread-junk-rule` disables it.

## Marketplace scan (level 3)

**Deliverable:** `docs/key_findings/20261117-marketplace-scan.md`

**Scope:** Research what exists today in:
1. **Claude Code plugins** — the official `anthropic-plugins/` marketplace + community lists
2. **Claude skills** (Chris's own + superpowers + public marketplaces)
3. **MCP servers** for Outlook/Exchange/Graph/Gmail — anything that ships ready-to-use
4. **Non-Claude commercial tools** — SaneBox, Superhuman, Shortwave, Boomerang — note their approaches, not to install, but to steal patterns

**Format per entry:**
- Name + link
- What it does
- Signal strength (actively maintained? adoption?)
- Relevance to outlook-tool (steal / integrate / ignore / already-reimplemented)

**Output:** one-page summary at top, detailed per-entry below.

## The How (as decided)

1. **/gitsmith setup** — branch `bolt/20261117-bolt6-advanced-triage` off main (already done)
2. **Phase A (parallel):**
   - **A1 implementer** — 5 pattern rules + `_scan_sent_to_sender` helper + tests
   - **A2 implementer** — `scripts/triage_llm.py` + Claude API integration + tests + `--llm` wiring in triage.py
   - **A3 researcher** — marketplace scan, writes `docs/key_findings/20261117-marketplace-scan.md`
3. **Phase B — Integration:**
   - `apply_triage.py --unread-junk-rule` gate
   - `scripts/triage.py` LLM fallback wiring
   - Add `--llm` flag end-to-end
4. **Phase C — Live validation:**
   - Fresh triage pass with `--llm` on Chris's Inbox
   - **HARD PAUSE** — show Chris the counts + sample verdicts
   - On approval: apply delete bucket (soft-delete, recoverable)
   - Target: 300+ additional cleared, ≥ 95% precision on LLM "delete" verdicts
5. **/pm close-out** — BACKLOG + CURRENT-SPRINT + captain's log
6. **/gitsmith audit + push + PR + merge**

## What's left to the builder

- Prompt engineering details for the LLM classifier (model will iterate on specific wording during testing)
- Exact confidence thresholds — may tune during live validation
- Marketplace scan source list — researcher picks which marketplaces to check (Smithery, Claude's official plugin list, GitHub topic searches)
- Specific cold-outreach domain patterns — researcher adds to `_COLD_OUTREACH_PATTERNS` list

## Risks + mitigations

| Risk | Mitigation |
|---|---|
| LLM mis-classifies a real message as junk | HARD PAUSE before apply; confidence threshold 0.85+; soft-delete only (recoverable) |
| LLM cost overruns | 200-msg default cap, explicit `--llm-limit`, prompt caching |
| Sent-Items scan is slow | Cache per-sender result in-memory during one classifier run |
| Pattern-rule false positives cascade with LLM | Rules fire FIRST, LLM only runs on keep bucket — can't double-flag |
| Marketplace scan returns stale/dead projects | Researcher notes "last commit" date + stars + download count |

## Success criteria

- [ ] 5 new pattern rules shipped + tested
- [ ] LLM classifier shipped + 10+ unit tests (mocked API)
- [ ] Marketplace scan doc delivered
- [ ] `--llm` + `--unread-junk-rule` flags end-to-end
- [ ] 300+ additional messages cleared on Chris's real Inbox
- [ ] ≥ 95% precision on LLM "junk" verdicts (Chris spot-checks 20, finds ≤ 1 miscategorized)
- [ ] Zero regressions on Bolt 4+5 workflow
- [ ] All 181 existing tests + new tests pass
- [ ] Captain's log + CHANGELOG updated
- [ ] PR opened + merged to main

## Approved

Chris 2026-11-17 ("1, 2, 3 — also, we can delete things I have not read that meet the criteria of junk. /pm and /staff on autopilot, full autonomous following /gitsmith gospel, let's go!").
