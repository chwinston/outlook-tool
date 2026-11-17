---
name: inbox-triage
description: Classify and sweep the Outlook inbox into buckets (safe-delete, resolved-ticket, low-signal-client, needs-action, keep) using rule-based heuristics. Two-step workflow — classifier produces a JSON report, executor runs the actual writes behind a dry-run gate.
user-invocable: true
allowed-tools: Bash, Read
argument-hint: [mode] [args...] — classify | preview <bucket> | apply <bucket> (defaults to classify+preview safe-delete)
---

# /inbox-triage — Rule-based inbox triage

Two commands. First classify. Then preview or apply.

> Panel-approved shape (Tim/Rob/Fran/Al, 2026-11-16): rules not ML, dry-run default, soft-delete only, every classification carries the rule name that produced it.

## Trigger

User says any of:
- "triage my inbox"
- "clean up my inbox"
- "show me what's safe to delete"
- `/inbox-triage`
- `/inbox-triage apply safe-delete`

## The two-command workflow

### 1. Classify

```bash
python scripts/triage.py --window-days 30 --max-results 500
```

Writes `output/triage-YYYYMMDD-HHMM.json`. Prints the path on stdout and a bucket summary on stderr.

Flags:
- `--window-days N` — scan messages received in the last N days (default: 30)
- `--max-results N` — cap on classification (default: 500)
- `--folder NAME` — folder to scan (default: Inbox)
- `--backend applescript|win32com|graph` — force a backend (default: auto-detect)
- `--output PATH` — override the report path

### 2. Preview the bucket you want to act on

```bash
python scripts/apply_triage.py output/triage-YYYYMMDD-HHMM.json --bucket safe-delete
```

Prints a Markdown table (sender · subject · matched rule · received time). No writes happen without `--apply`.

### 3. Apply (optional)

```bash
python scripts/apply_triage.py output/triage-YYYYMMDD-HHMM.json --bucket safe-delete --apply
```

Fires the bucket's action:

| Bucket | Action with --apply |
|---|---|
| `safe-delete` | Soft-delete → Deleted Items |
| `resolved-ticket` | Move → Archive |
| `low-signal-client` | Move → Archive |
| `needs-action` | Preview-only — no auto-action |
| `keep` | Preview-only — no auto-action |

## Hard rules the executor enforces

- **Reports older than 1 hour are refused.** AppleScript message ids are session-scoped. Regenerate with step 1 and retry.
- **Verified, not status.** An `OperationResult` with `status="success"` but `verified=False` is counted as unverified-success and surfaces loudly. The exit code is non-zero.
- **No hard-delete through the executor.** `--yes-really-permanent` is refused. Use `python cli.py delete --permanent --yes-really-permanent <id>` per-message if you ever need it.
- **Fingerprint fallback on stale ids.** If the initial write fails with a "message not found" style error, the executor re-locates by `(sender_email, subject, received_date ± 1 day)` and retries once with the fresh id. Ambiguous matches abort for that one message.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Clean — every attempted op verified-success, or nothing to do |
| 1 | Some op failed, was skipped, or returned unverified-success |
| 2 | Report too old / backend unsupported / safety refusal |

## Example session

```text
$ python scripts/triage.py --window-days 7
Wrote /Users/chris/repos/outlook-tool/output/triage-20261116-1030.json
  backend: applescript
  classified: 142 messages
    safe-delete            58
    resolved-ticket         9
    low-signal-client       4
    needs-action           21
    keep                   50
/Users/chris/repos/outlook-tool/output/triage-20261116-1030.json

$ python scripts/apply_triage.py output/triage-20261116-1030.json --bucket safe-delete
## Triage Preview — safe-delete bucket (58 messages)
Report: output/triage-20261116-1030.json (generated: 2026-11-16T10:30:00, age: 3 min)
Backend: applescript
Action with --apply: soft-delete

| # | Received | Sender | Subject | Matched rule |
|---|---|---|---|---|
| 1 | 2026-11-16T09:12 | no-reply@sns.amazonaws.com | ALARM: "callhero-foo"… | cloudwatch-alarm-old |
…

DRY RUN — re-run with --apply to execute (safe-delete bucket).

$ python scripts/apply_triage.py output/triage-20261116-1030.json --bucket safe-delete --apply
Applying safe-delete → soft-delete via applescript backend…
  [1/58] no-reply@sns.amazonaws.com — ALARM: "callhero-foo"…
…
## Apply Result — safe-delete → soft-delete (58 attempted)

- verified-success: 56
- UNVERIFIED-success (silent-success risk): 2
- failed: 0
- unsupported: 0
- skipped (ambiguous fingerprint): 0

WARNING: unverified-success rows need manual spot-checking.
```

## Workflow conversational shape

When the user says "triage my inbox":

1. Run step 1 (classify).
2. Show the bucket summary.
3. Ask which bucket to preview. Default to `safe-delete`.
4. Run step 2 (preview) and show the table.
5. Ask whether to apply. Only run step 3 with explicit confirmation.
6. After an apply run, read the `verified-success` count and report it. If `UNVERIFIED-success > 0`, highlight those rows — they are the Phase 1 silent-success guard surfacing something.

## When to tune the rules

Before adding a rule, re-read `scripts/triage_rules.py`. Every rule lists in its docstring the cluster it came from in the 2026-11-16 inbox sample. Rules are ordered — first match wins — so new rules go in the right bucket section, not at the bottom.

Trigger to re-sample: rule count exceeds 40, or Chris overrides the bucket on more than 10% of a bucket's rows. Below that threshold, add a rule, do not retrain anything.

## Related files

- `scripts/triage.py` — classifier
- `scripts/apply_triage.py` — executor
- `scripts/triage_rules.py` — rule base (19 rules, 5 buckets)
- `docs/manuals/inbox-triage.md` — operator reference
- `outlook_tool.py` — OperationResult + backend write verbs (Phase 1)
