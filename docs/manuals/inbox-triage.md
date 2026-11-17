# Inbox triage operator manual

**Audience:** Chris + whoever onboards next.
**What it does:** Classifies recent inbox mail into five buckets (`safe-delete`, `resolved-ticket`, `low-signal-client`, `needs-action`, `keep`) using rule-based heuristics, then — behind a dry-run gate — fires the corresponding move / soft-delete via the Phase 1 write verbs on `OutlookClient`.

This is a two-command workflow by design. Classification and execution are separate steps so you always get a chance to look at the table before anything writes.

---

## TL;DR

```bash
# 1. Classify (produces output/triage-YYYYMMDD-HHMM.json)
python scripts/triage.py --window-days 30

# 2. Preview the bucket you care about
python scripts/apply_triage.py output/triage-20261116-1030.json --bucket safe-delete

# 3. When you're happy with the table — apply
python scripts/apply_triage.py output/triage-20261116-1030.json --bucket safe-delete --apply
```

If you forget step 1, step 2 can't run. If step 2 shows anything you didn't expect in the table, tune `scripts/triage_rules.py` and re-run. Only then step 3.

---

## The buckets

| Bucket | Meaning | Action with `--apply` |
|---|---|---|
| `safe-delete` | Automated noise with no live signal — CloudWatch alarms > 3d old, marketing drip, dependabot > 7d, generic no-reply > 14d | Soft-delete (move to Deleted Items) |
| `resolved-ticket` | Platform ticket activity that is archivable, not immediately destroyable (HubSpot ticket comments, Zendesk vendor updates, "ticket #X closed") | Move → `Archive` |
| `low-signal-client` | Routine client FYI mail with no explicit ask (MyStudio failed-payment reruns, HubSpot daily task digest, Workday workflow notifications) | Move → `Archive` |
| `needs-action` | A real human owes a reply. CI failures, HubSpot task assignments, subjects with action verbs or question marks, direct mail from internal coworkers | Preview-only |
| `keep` | Everything the rules didn't pattern-match. The default. | Preview-only |

The two preview-only buckets (`needs-action`, `keep`) deliberately have no auto-action. You read them, you decide. The executor doesn't move or delete them even with `--apply`.

---

## Safety invariants

### 1. Dry-run default

`python scripts/apply_triage.py <report> --bucket <b>` with no `--apply` prints the Markdown preview and exits. Zero writes. No exceptions.

### 2. Report freshness

Reports older than **1 hour** are refused outright with exit code 2:

```
REFUSING to apply: triage report is 84 min old (max 60 min). AppleScript message
ids go stale across sessions — regenerate with 'python scripts/triage.py' and retry.
```

Why: AppleScript message indices are **session-scoped**. Outlook reorders messages, incoming mail shifts indices, "message 7 of Inbox" means a different thing after 30 minutes. The executor re-reads the classification but trusts the ids inside — so stale ids are a correctness risk, not just a UX quirk.

If the refusal fires, run `python scripts/triage.py` again and point the executor at the new report.

### 3. Verified, not status

The Phase 1 contract (`OperationResult`) distinguishes two success signals:

- `status == "success"` — the write call returned without an exception
- `verified == True` — a post-op re-fetch confirmed the state change actually landed

A `status=success, verified=False` combination is the **silent-success** path — Outlook returned OK but the state didn't change. The executor counts those as `UNVERIFIED-success` in the apply summary and exits non-zero:

```
## Apply Result — safe-delete → soft-delete (58 attempted)

- verified-success: 56
- UNVERIFIED-success (silent-success risk): 2
- failed: 0
```

**You treat unverified-success as a soft failure.** Spot-check those two rows in Outlook before you trust the next run.

### 4. Fingerprint fallback on stale ids

If an individual write fails with a "message not found" style error (which happens when the inbox reordered between classify and apply), the executor automatically:

1. Searches the source folder for `(sender_email, subject, received_date ± 1 day)`
2. If exactly one match comes back, retries the write with that fresh id
3. If zero or multiple matches come back, records the row as skipped-ambiguous and moves on

Skipped-ambiguous rows appear in the apply summary. You go read them and decide whether to re-classify or hand-handle.

### 5. No hard-delete through the executor

`apply_triage.py` **only soft-deletes** (moves to Deleted Items). The `--yes-really-permanent` flag is explicitly refused:

```
REFUSED: --yes-really-permanent is not available via the triage executor.
Hard-delete is per-message only, via 'python cli.py delete --permanent --yes-really-permanent <id>'.
```

Hard-delete is a per-message, deliberate operation. Batch hard-delete is explicitly a no-build (panel F6 unanimous).

---

## Commands

### `scripts/triage.py` — classifier

```
usage: triage [-h] [--window-days WINDOW_DAYS] [--max-results MAX_RESULTS]
              [--folder FOLDER] [--backend {applescript,win32com,graph}]
              [--output PATH] [--quiet]
```

| Flag | Default | Notes |
|---|---|---|
| `--window-days N` | 30 | Scan messages received in the last N days |
| `--max-results N` | 500 | Cap on classified messages |
| `--folder NAME` | `Inbox` | Any folder name Outlook knows about |
| `--backend` | auto | Force `applescript`, `win32com`, or `graph` |
| `--output PATH` | `output/triage-YYYYMMDD-HHMM.json` | Report path override |
| `--quiet` | off | Skip the bucket summary on stderr |

**Report shape (schema_version=1):**

```json
{
  "schema_version": 1,
  "generated_at": "2026-11-16T16:42:30.580445",
  "backend": "applescript",
  "sample_folder": "Inbox",
  "sample_size": 266,
  "window_days": 30,
  "summary": {
    "safe-delete": 92, "resolved-ticket": 17,
    "low-signal-client": 6, "needs-action": 52, "keep": 99
  },
  "messages": [
    {
      "message_id": "as_msg_7",
      "sender_email": "no-reply@sns.amazonaws.com",
      "sender_name": "AWS Notifications",
      "subject": "ALARM: …",
      "received_at": "2026-10-28T03:14:22",
      "folder": "Inbox",
      "bucket": "safe-delete",
      "matched_rule": "cloudwatch-alarm-old"
    }
  ]
}
```

`matched_rule` is load-bearing: it tells you (and future-you) why a message landed in that bucket. When a bucket looks wrong, read the rule. Do not tune by changing the classification, tune by changing the rule.

### `scripts/apply_triage.py` — executor

```
usage: apply_triage [-h] [--report REPORT_FLAG] [--bucket {safe-delete,resolved-ticket,low-signal-client,needs-action,keep}]
                    [--source-folder SOURCE_FOLDER] [--apply] [--output PATH]
                    [report]
```

| Flag | Default | Notes |
|---|---|---|
| `report` (positional) | — | Path to the triage report JSON |
| `--report PATH` | — | Long form of the positional report path |
| `--bucket NAME` | `safe-delete` | Which bucket to preview or apply |
| `--source-folder NAME` | `Inbox` | Folder the messages currently live in |
| `--apply` | off | Fire the writes. Omit for Markdown preview |
| `--output PATH` | — | Write preview (+ apply summary) to PATH in addition to stdout |

Exit codes:

| Code | Meaning |
|---|---|
| 0 | Clean — every op verified-success, or nothing to do |
| 1 | At least one row: failed, skipped-ambiguous, or unverified-success |
| 2 | Report too old, unsupported-only path, or safety refusal |

---

## Tuning the rules

Rules live in `scripts/triage_rules.py` as a single ordered list (`RULES`). Each rule is a named callable. The classifier applies them top-to-bottom and the first match wins. This ordering matters:

1. `needs-action` rules run first so action-signalling mail doesn't get swept up by a broader auto-delete sweep (example: a `[Build failed]` from Azure DevOps lands in needs-action even though the sender is automated).
2. `resolved-ticket` / `low-signal-client` rules run next so archivable noise doesn't get soft-deleted.
3. `safe-delete` rules run last. Most of them require an age threshold (> 3 days, > 7 days, > 14 days) so fresh automated mail still gets a chance.
4. `keep` is the implicit fall-through.

**When to add a rule:** you ran the previews a couple of times and a dominant pattern keeps landing in `keep` or the wrong bucket. Add a specific rule, name it, put it in the right section, write a one-liner docstring that names the cluster.

**When to NOT add a rule:** when a single edge case jumped at you. Precision > recall — one weird promo landing in `keep` is cheaper than one real reply landing in `safe-delete`.

**Panel-named escape hatches:**

- Rule count exceeds 40 → sample again and consider consolidating or graduating to ML (backlog says ML stays deferred until that threshold)
- Chris overrides a bucket on > 10% of its rows → re-sample and tune the dominant rule, not individual messages

---

## Troubleshooting

**"REFUSING to apply: triage report is 84 min old"**
→ Regenerate with `python scripts/triage.py` and retry inside the 1-hour window.

**"UNVERIFIED-success" count > 0 in the apply summary**
→ Open Outlook, look at Deleted Items (or Archive, depending on bucket). Spot-check that the flagged rows really landed. If they didn't, you've hit a Phase 1 silent-success path and you should re-run the executor — the fingerprint fallback will pick it up on the retry. If they DID land, the verify just couldn't find them in the top 50 of the destination folder; benign but worth noting if it recurs.

**"skipped (ambiguous fingerprint)" count > 0**
→ The classify report had a message id that went stale AND the fingerprint lookup found either zero or multiple matches. Re-classify and retry; if it keeps happening on the same message, hand-handle it in Outlook.

**"unsupported" count > 0 (Graph backend)**
→ The Graph backend needs `Mail.ReadWrite` scope and you don't have it. See `docs/pm/BACKLOG.md § Deferred from Bolt 4` — "Graph `Mail.ReadWrite` scope acquisition (IT ticket)" is the open path.

**Report file isn't there when you try to apply**
→ `python scripts/triage.py` prints the path on stdout — `cp` it or pipe it. You can also just `ls output/triage-*.json` and grab the latest.

---

## Related

- Phase 1 write-verb contract: `OperationResult` in `outlook_tool.py`
- CLI write verbs: `python cli.py move|delete|mark-read --help`
- Spec: `docs/pm/plans/20261116-1100-inbox-triage-plan.md`
- Panel report: `docs/key_findings/20261116-1100-inbox-triage-write-verbs-Staff-Engineer-Panel.md`
- Backlog: `docs/pm/BACKLOG.md § Bolt 4`
