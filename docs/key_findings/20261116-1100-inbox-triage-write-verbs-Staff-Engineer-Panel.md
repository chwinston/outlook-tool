# Inbox Triage + Write Verbs — Staff Engineer Panel Analysis

**Date:** 2026-11-16 11:00
**Panel:** Tim (SpaceX), Rob (Roblox), Fran (Meta), Al (AWS Mac/COM/Graph), Will Larson (Moderator)
**Trigger:** Chris needs to triage his Outlook inbox and bulk-archive/delete safely. The Claude Code M365 connector is read-only (`Mail.Read` only). The local `outlook-tool` has read + send methods but no `move`/`delete`/`mark_read`. He wants a repeatable, safe pipeline: classify first, act second, with dry-run defaults.

---

## Problem Statement

### What's happening

Chris's Outlook inbox has accumulated a large backlog mixing:
- Obvious junk (Zenoti password resets, HubSpot digest emails, LinkedIn notifications, calendar invites to past events)
- Resolved threads on closed HubSpot tickets
- Active client threads that still need action
- Reference material (contracts, financial PDFs)

Today there is no automated or semi-automated way to separate these buckets. The Claude Code M365 connector can read but can't move/delete. The local `outlook-tool` has `search`/`send` but no inbox-management methods. SaneBox / Superhuman / Clean Email are the market alternatives but introduce another subscription + don't know Chris's specific client list, ticket state, or AM roster.

### Why it matters

- Chris's inbox is a known bottleneck for client-fire detection (he's missed real signals because they sat below the notification noise)
- The /tasko walk pattern surfaces client fires from HubSpot tickets but doesn't surface fires that live only in email until they've already escalated
- Every /pm cycle begins with "did Chris miss anything in his inbox since last session?" and the answer is "we don't know without a manual scan"

### Root cause chain

1. `outlook-tool` has been read + send only since Bolt 1 — write-to-mailbox was deferred as "not needed yet"
2. The M365 Graph connector in Claude Code is also read-only by OAuth scope
3. So there's no programmatic write path from any AI session to the mailbox
4. So triage has to be manual — which means it doesn't happen
5. So the inbox grows

### Constraints

- Mac-primary (Outlook Legacy/Classic mode via AppleScript); Windows must still work (COM path); Graph fallback must gracefully degrade when scope is missing
- Cannot require admin tenant consent on short notice (`Mail.ReadWrite` on Graph is deferred)
- Cannot rely on the message being ACTUALLY deleted by AppleScript — need round-trip verify (AppleScript can silently fail)
- Dry-run must be the default — wrong-message deletion is high-blast-radius and not reversible from Chris's point of view even when the message goes to Deleted Items
- Build must fit the existing `_AppleScriptBackend` / `_GraphBackend` / `OutlookClient` three-layer architecture (defined at outlook_tool.py:131, :596, :740)

### Work type
Feature (multi-backend write surface + new skill)

### Impact summary
Adds a write dimension to a previously read-only tool. Touches 3 existing backend classes + 1 new CLI verb + 1 new Claude Code skill (`/inbox-triage`). No breaking changes to existing read/send surface. New skill is additive.

---

## Panel Analysis

### Tim — Staff Engineer, SpaceX

**Risk assessment:**
Write operations against a live inbox are irreversible from a user's perspective even when the message goes to Deleted Items. Three distinct risk axes:

1. **Wrong-message blast** — classifier misclassifies, we move/delete a real email
2. **Wrong-folder blast** — move succeeds but lands in a folder the user can't find later
3. **Silent-success blast** — AppleScript "succeeds" but Outlook refused the operation and we never notice. This is the dominant failure mode on Mac.

**Options evaluated:**
- A — Ship write methods with best-effort error handling: NO. The silent-success case is not a best-effort problem; it's a correctness problem.
- B — Ship write methods with round-trip verify (read the message BEFORE and AFTER the write, confirm state change): YES. This is the only way to catch AppleScript silent-success and Graph throttle swallowing.
- C — Wait until we have Graph `Mail.ReadWrite` so we have a reliable API: NO. That's a multi-week IT dependency and the Mac AppleScript path is actually more reliable than Graph for single-user operations today.

**Key quote:**
> *"The failure mode we actually see on Mac is AppleScript returning success while Outlook refused the operation internally. Dry-run and preview are not enough — every write wraps a round-trip verify, or we're not shipping safely."*

**Recommendation:**
Option B. Dry-run is the DEFAULT. Round-trip verify is mandatory. `--permanent` flag exists on `delete` but requires double-flag (`--permanent --yes-really-permanent`) to actually hard-delete; soft-delete to Deleted Items is the default. All ops log before+after state to a session log file.

**On the proposed `/inbox-triage` skill:**
Separate classifier from executor. Classifier outputs a JSON report. Executor consumes the report. Both can be invoked independently. The skill's default behavior is classify-only; executing requires an explicit `--apply` flag + a signed-off JSON report that's at most 1 hour old.

**Unique contribution:**
The round-trip-verify requirement. Without it the AppleScript backend is a liability.

---

### Rob — Staff Engineer, Roblox

**Risk assessment:**
The three backends (COM / AppleScript / Graph) will diverge the moment one gets a new feature. This is already visible today in how `search` behaves slightly differently across backends. If we ship `move`/`delete`/`mark_read` as three sibling implementations without a shared contract, the triage workflow consuming them has to know which backend ran, and divergence compounds over time.

**Options evaluated:**
- A — Three sibling implementations with "best effort" compatibility: NO. Same problem as `search` has today.
- B — Define the contract on `OutlookClient` as a unified `OperationResult` dict, each backend returns the same shape: YES. The triage layer never needs to know which backend ran.
- C — Pick ONE backend for writes (just COM + AppleScript, leave Graph out): NO. The whole point of Graph is for the cases when Chris is on a machine without Outlook running.

**Key quote:**
> *"The contract is what matters. Each backend can lie about how it got the answer; the triage layer only needs to trust that the answer is shaped the same way."*

**Recommendation:**
Option B. Define `OperationResult` as:
```python
@dataclass
class OperationResult:
    status: Literal["ok", "error", "unsupported"]
    verified: bool               # did round-trip confirm the state change?
    message_id: str
    operation: Literal["move", "delete", "mark_read"]
    before_state: dict           # folder, is_read, etc. before the op
    after_state: dict | None     # None if status != "ok"
    error_detail: str | None
    backend: Literal["com", "applescript", "graph"]
    elapsed_ms: int
```

Graph backend returns `status="unsupported"` with `error_detail="OAuth grant lacks Mail.ReadWrite"` when scope is missing. **Doesn't silently fail. Doesn't auto-upgrade the scope.** Surface the cliff loudly so Chris knows what to ask IT for if he wants cloud writes.

**On the `/inbox-triage` skill:**
The classifier should emit the triage report with explicit per-message rule provenance — "flagged as safe-delete because: sender in junk-domain-list AND age > 30d AND no attachments." That way when Chris disagrees with a classification, there's a clear line to the rule to tune, not a black-box retrain.

**Unique contribution:**
The unified `OperationResult` contract + the cost-of-divergence framing. Also: rule provenance in the triage report.

---

### Fran — Staff Engineer, Meta

**Risk assessment:**
Trying to do classification and execution in one pass creates a situation where every rule change requires re-running the whole pipeline. That's the opposite of what Chris needs for a tool he'll run weekly for years.

**Options evaluated:**
- A — Classifier + executor in one CLI command: NO. Doesn't separate concerns.
- B — Classifier and executor as two separate CLI verbs, with a JSON triage report as the hand-off: YES.
- C — Rules engine + ML classifier hybrid: NO. V1 is heuristic only. ML is v2 at earliest.

**Key quote:**
> *"Classify first, act second. Two commands, one JSON file passing between them. Tuning the rules is editing a Python file, not wrestling with an end-to-end pipeline."*

**Recommendation:**
Option B. Two CLI verbs:
- `outlook-tool triage [--inbox-only] [--since <date>] [--rules <file>] --output triage-YYYYMMDD-HHMM.json` — classifies, doesn't write
- `outlook-tool apply --triage-report <file> [--bucket safe-delete|needs-action|keep] [--apply]` — reads the report, surfaces the preview, applies when `--apply`

Bucket ordering (most confident first):
1. **Safe-delete** — automated notifications, password resets, calendar invites to past events, HubSpot digest emails. Highest-confidence bucket, archive these first.
2. **Resolved-ticket threads** — replies on HubSpot tickets that are now closed. Second-highest confidence. Move to an `Archive/Resolved/<year-month>` folder.
3. **Low-signal client threads** — client emails older than 30 days with no outstanding action from Chris (replies already sent, no open task/ticket). Move to `Archive/Client/<year-month>`.
4. **Needs-action** — emails where there's an open task/ticket, or where the thread ends on an incoming message with no reply from Chris. **Never auto-touched.** Surfaced as a sorted list for Chris to review.
5. **Keep** — contracts, financial PDFs (attachments > 1MB that look like documents), explicit "keep" sender list. Never auto-touched.

**On the apply step:**
The `apply` command's default when `--apply` is NOT passed should be: print a markdown preview, write nothing. Only when `--apply` is explicitly passed does it actually move/delete. This is Tim's dry-run-by-default rule, codified as the CLI shape.

**Unique contribution:**
The two-CLI-verb separation + bucket ordering by confidence. Also: the `--bucket` filter so Chris can apply just one bucket at a time.

---

### Al — Staff Engineer, AWS (Mac AppleScript + COM + Microsoft Graph platform expertise)

**Risk assessment:**
Three platform-specific landmines:

1. **AppleScript on Mac is slow.** Single `move` takes 400-800ms. 500-message archive serially = 3-5 minutes with UI freeze. Chris runs this weekly; 5-minute operations don't get used.
2. **Mac "Archive" is per-account.** The "Archive" button in the Outlook UI maps to each account's own Archive folder. If you point AppleScript at `folder "Archive"` it may not exist. Must detect-and-create on first use.
3. **Graph API throttling is aggressive.** 10 req/sec/user. Naive loop over 500 messages burns the whole budget and gets 429'd. `$batch` endpoint (up to 20 ops per call) is required for anything over ~50 messages.

**Options evaluated:**
- A — Ship v1 serial, optimize later: ACCEPTABLE for correctness but Chris won't use a 5-minute tool. Need batching for throughput.
- B — Ship v1 with batching from day one: NO. Batching + round-trip-verify at the same time is too much to validate at once. Ship v1 serial-with-verify; v2 adds batching.
- C — Pick one backend only for writes: NO. Same rejection as Rob gave.

**Key quote:**
> *"AppleScript bulk-move is 5-10× faster than serial because of how dispatch works. But you don't build v1 on it — you build v1 serial-with-verify, prove correctness, THEN add batching as v2. Ship-then-optimize is the right order."*

**Recommendation:**
Phase 1 (v1 correctness):
- Serial `move`/`delete`/`mark_read` with round-trip verify
- 100ms sleep between ops (politeness, not throttling — AppleScript dispatch is the bottleneck, not the Outlook app)
- Progress log every 10 messages
- Mac: detect-and-create Archive folder on first use
- Graph: refuse write ops with `status="unsupported"` when scope is missing

Phase 2 (v2 batching):
- AppleScript: single `move messages {id1, id2, ...} to folder "Archive"` per target folder (not per message). Still verify after, but batch the write.
- Graph: `$batch` POST up to 20 ops per call. Honor throttle backoff via `Retry-After` header.
- Progress log every 50 messages.

**Phase 3 (v2.5, deferred):**
- Cross-mailbox delegation (if Chris needs to triage a shared mailbox)
- Recurring scheduled triage (cron-style)

**On Mac Archive folder detection:**
```python
# Mac AppleScript — detect-and-create Archive folder for the default account
try:
    folders = applescript_run('tell application "Microsoft Outlook" to name of every folder of default account')
    if "Archive" not in folders:
        applescript_run('tell application "Microsoft Outlook" to make new folder at default account with properties {name:"Archive"}')
except AppleScriptError as e:
    # Fall back to Deleted Items? No — refuse the op with a clear error.
    return OperationResult(status="error", error_detail=f"Archive folder detection failed: {e}", ...)
```

**Unique contribution:**
AppleScript bulk-move batching (Phase 2) + Graph `$batch` (Phase 2) + Mac Archive folder detect-and-create (Phase 1).

---

## Consensus Matrix

| Question | Tim (SpaceX) | Rob (Roblox) | Fran (Meta) | Al (AWS) |
|---|---|---|---|---|
| Dry-run as default on CLI | YES | YES | YES | YES |
| Round-trip verify on every write | YES | YES | YES | YES |
| Unified `OperationResult` across backends | YES | YES | YES | YES |
| Classifier + executor as separate CLI verbs | YES | YES | YES | YES |
| Soft-delete to Deleted Items as default | YES | YES | YES | YES |
| Graph backend returns `unsupported` without scope | YES | YES | YES | YES |
| v1 serial, v2 batching | YES | YES | YES | YES |
| Mac Archive folder detect-and-create | YES (via Al) | YES | YES | YES (owner) |
| Rule provenance in triage report | YES | YES (proposed) | YES | YES |
| `--permanent` requires double-flag | YES (proposed) | YES | YES | YES |

**Unanimous.** No open disagreements.

---

## Will Larson's Decision

**Scope:** Ship in three phases. Phase 1 = write methods on `OutlookClient` (correctness). Phase 2 = `/inbox-triage` skill (workflow). Phase 3 = batching (throughput).

### Phase 1 — Write methods on OutlookClient (~1 dev day)

| Step | What | Why | Effort |
|---|---|---|---|
| 1.1 | Add `OperationResult` dataclass to `outlook_tool.py` — shared contract per Rob F1 | Backends return uniform shape | 15 min |
| 1.2 | Add `move(message_id, target_folder) → OperationResult` on `OutlookClient` as the dispatch method | Single public API | 15 min |
| 1.3 | Implement `_AppleScriptBackend.move()` with detect-and-create Archive folder per Al F3 + round-trip verify per Tim F1 | Mac path | 90 min |
| 1.4 | Implement `_GraphBackend.move()` — detects scope, returns `unsupported` without `Mail.ReadWrite` per Rob F2 | Graceful degrade | 45 min |
| 1.5 | Add COM backend `move()` (Windows) for parity — detect Archive folder, round-trip verify | Windows path | 60 min |
| 1.6 | Add `delete(message_id, permanent=False) → OperationResult` — soft-delete default, `permanent=True` requires `yes_really_permanent=True` per Tim | Dangerous op, guarded | 60 min |
| 1.7 | Add `mark_read(message_id, read=True) → OperationResult` — lowest-risk write; implement all three backends | Round out the surface | 45 min |
| 1.8 | Add CLI verbs: `outlook-tool move <msg_id> <folder>`, `outlook-tool delete <msg_id>`, `outlook-tool mark-read <msg_id>` — all dry-run default, `--apply` required for real write per Fran | CLI consistency | 60 min |
| 1.9 | Unit + integration tests covering all 3 backends × 3 ops. Tests use fixture mailbox / respx mock for Graph | Correctness guard | 120 min |
| 1.10 | Add CHANGELOG entry + README section documenting the new verbs + the `yes_really_permanent` safety flag | Operator docs | 30 min |

**Total Phase 1: ~8-10 hours of focused dev time**

### Phase 2 — `/inbox-triage` skill (~1 dev day)

| Step | What | Why | Effort |
|---|---|---|---|
| 2.1 | Create `~/.claude/skills/inbox-triage/SKILL.md` describing the two-step workflow | User-invocable skill | 30 min |
| 2.2 | Write `scripts/triage.py` — classifier. Reads inbox via `outlook_tool.search`, applies rules from `scripts/triage_rules.py`, outputs `triage-YYYYMMDD-HHMM.json` | Classification | 180 min |
| 2.3 | Write `scripts/triage_rules.py` — initial rule set with 5 buckets per Fran. Buckets: safe-delete / resolved-ticket / low-signal-client / needs-action / keep | Rule base | 90 min |
| 2.4 | Write `scripts/apply_triage.py` — reads a triage report, prints markdown preview, applies writes when `--apply` passed. Per-bucket filter via `--bucket`. Rejects reports > 1 hour old per Tim | Executor | 120 min |
| 2.5 | Add rule provenance to classifier output per Rob — every classified message has a `matched_rule` field with the rule id that flagged it | Debuggability | 45 min |
| 2.6 | Dry-run test on Chris's real inbox: produce a triage report, review the markdown, confirm the bucket distribution looks sane before shipping | Live validation | 60 min |
| 2.7 | Write short operator manual at `docs/manuals/inbox-triage.md` — the two-command workflow, how to tune rules, how to review a report | Operator docs | 45 min |

**Total Phase 2: ~9 hours**

### Phase 3 — Batching (~0.5 dev day, deferred)

| Step | What | Why | Effort |
|---|---|---|---|
| 3.1 | AppleScript bulk-move: single `move messages {id1, ...} to folder "Archive"` per target folder. Verify after. | 5-10× speedup | 90 min |
| 3.2 | Graph `$batch` POST: up to 20 ops per call, honor `Retry-After` on 429 | Avoid throttle | 90 min |
| 3.3 | Benchmark 500-message archive before + after | Prove the win | 30 min |

**Total Phase 3: ~3-4 hours. Deferred until Phase 1 + 2 are in use for ≥ 1 week.**

### Total: ~17 hours of focused dev (Phase 1 + 2 only) + ~3-4 hours deferred

### What's explicitly deferred

| Item | Rationale | Revisit When |
|---|---|---|
| ML classifier (sender clustering, subject topic model) | V1 heuristic rules are enough for Chris's inbox shape. ML adds training-set maintenance burden with marginal win. | Rule count exceeds ~40 OR Chris's bucket-override rate exceeds 10% of classifications |
| Recurring scheduled triage (daily cron) | Manual trigger is fine until classifier confidence is proven over weeks of use | After 4 consecutive weeks of manual use with < 5% misclassifications |
| Cross-mailbox delegation | Chris doesn't delegate today; not needed | First time Chris asks to triage someone else's mailbox |
| Graph `Mail.ReadWrite` scope acquisition | IT ticket needed; AppleScript / COM path is sufficient on Chris's primary machine | First time Chris needs to triage from a machine where Outlook isn't running |
| Hard-delete batch mode (Phase 3) | `--permanent --yes-really-permanent` is a per-message operation by design; batching it is a foot-gun | Never — explicit no-build |

### Key takeaways

> *"Classify first, act second. The JSON report in between is the audit trail."* — Fran

> *"Three backends, one contract. The triage layer doesn't care which one ran."* — Rob

> *"Round-trip verify is the only defense against AppleScript silent-success. Non-negotiable."* — Tim

> *"Mac Archive folder is per-account and may not exist. Detect-and-create or refuse loudly — never assume."* — Al

---

## F-Tag Graduation Ledger

| F-tag | Finding (one line) | Disposition | Backlog row / Reason |
|-------|--------------------|-------------|----------------------|
| F1 | Round-trip verify on every write operation (Tim) | graduated | `BACKLOG.md` → "Phase 1.3/1.5 — Round-trip verify in AppleScript + COM backends" (owner: Chris, trigger: Phase 1 kickoff, effort: S) |
| F2 | Graph backend refuses writes loudly when scope missing, doesn't auto-upgrade (Rob) | graduated | `BACKLOG.md` → "Phase 1.4 — Graph backend returns status=unsupported without Mail.ReadWrite" (owner: Chris, trigger: Phase 1 kickoff, effort: S) |
| F3 | Mac Archive folder detect-and-create on first use (Al) | graduated | `BACKLOG.md` → "Phase 1.3 — AppleScript backend detect-and-create Archive folder" (owner: Chris, trigger: Phase 1 kickoff, effort: S) |
| F4 | Unified `OperationResult` dataclass contract across backends (Rob) | graduated | `BACKLOG.md` → "Phase 1.1 — OperationResult dataclass" (owner: Chris, trigger: Phase 1 kickoff, effort: S) |
| F5 | Dry-run as default, `--apply` explicit opt-in (Panel unanimous) | graduated | `BACKLOG.md` → "Phase 1.8 — CLI dry-run default" + "Phase 2.4 — apply_triage.py dry-run default" (owner: Chris, trigger: Phase 1 + 2, effort: XS per row) |
| F6 | `--permanent` delete requires double-flag (Tim + unanimous) | graduated | `BACKLOG.md` → "Phase 1.6 — delete() requires yes_really_permanent=True for hard delete" (owner: Chris, trigger: Phase 1, effort: XS) |
| F7 | AppleScript bulk-move batching (Al) | graduated | `BACKLOG.md` → "Phase 3.1 — AppleScript bulk-move" (owner: Chris, trigger: Phase 1 + 2 live for ≥ 1 week, effort: S) |
| F8 | Graph `$batch` batching (Al) | graduated | `BACKLOG.md` → "Phase 3.2 — Graph $batch" (owner: Chris, trigger: Phase 1 + 2 live for ≥ 1 week AND Graph scope granted, effort: S) |
| F9 | Rule provenance (`matched_rule`) in triage report (Rob) | graduated | `BACKLOG.md` → "Phase 2.5 — rule provenance in classifier output" (owner: Chris, trigger: Phase 2 kickoff, effort: XS) |
| F10 | Classifier + executor as separate CLI verbs (Fran) | graduated | `BACKLOG.md` → "Phase 2.2 + 2.4 — triage and apply as distinct verbs" (owner: Chris, trigger: Phase 2 kickoff, effort: already in scope) |
| F11 | Reject triage reports older than 1 hour on `apply` (Tim) | graduated | `BACKLOG.md` → "Phase 2.4 — apply_triage.py rejects reports > 1h old" (owner: Chris, trigger: Phase 2 kickoff, effort: XS) |
| F12 | ML classifier for sender / subject clustering (deferred) | not-graduating | reason: v1 heuristic rules are sufficient for Chris's inbox shape; ML adds training-set maintenance with marginal win. Revisit trigger documented in "What's explicitly deferred." |
| F13 | Recurring scheduled triage via cron (deferred) | not-graduating | reason: manual trigger is correct for v1; schedule only after classifier is proven over weeks of use. Revisit trigger documented. |

---

## Files Referenced

| File | Role |
|---|---|
| `outlook_tool.py:131` (`_AppleScriptBackend`) | Mac backend — gets new `move`/`delete`/`mark_read` methods |
| `outlook_tool.py:596` (`_GraphBackend`) | Graph backend — gets new methods returning `unsupported` without `Mail.ReadWrite` |
| `outlook_tool.py:740` (`OutlookClient`) | Public API — gets new dispatch methods |
| `~/.claude/skills/inbox-triage/SKILL.md` (NEW) | Claude Code skill for the two-step workflow |
| `scripts/triage.py` (NEW) | Classifier |
| `scripts/triage_rules.py` (NEW) | Rule base |
| `scripts/apply_triage.py` (NEW) | Executor |
| `docs/manuals/inbox-triage.md` (NEW) | Operator manual |

## Implementation Plan

See `docs/pm/plans/20261116-1100-inbox-triage-plan.md` for the ordered task breakdown.

## Findings to Fix

Covered in F-Tag Graduation Ledger above. Each F-tag → `BACKLOG.md` row.
