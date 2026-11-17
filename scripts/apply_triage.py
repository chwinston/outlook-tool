#!/usr/bin/env python3
"""Execute a triage report — Bolt 4 Phase 2 (B4.2.4).

Reads a report produced by ``scripts/triage.py``, prints a Markdown preview of
the messages in the selected bucket, and — when ``--apply`` is set — fires the
corresponding :class:`OutlookClient` write verb and surfaces the result.

Safety invariants (from the staff panel + Phase 1 handoff)
----------------------------------------------------------
1. **Dry-run by default.** No ``--apply``, no writes.
2. **Report age.** AppleScript message ids are session-scoped. We refuse
   reports more than 1 hour old (panel F11).
3. **Verified, not status.** A ``status=="success"`` with ``verified=False``
   is the silent-success path (panel F1 / Tim). The executor surfaces
   verified separately and counts it in the ship-gate metrics.
4. **Fingerprint fallback.** When a direct AppleScript lookup misses (the
   inbox reordered, the id went stale), we re-scan the source folder for a
   ``(sender_email, subject, received_date)`` match and retry with the fresh
   index. If still ambiguous, we abort for that one message — never guess.
5. **No hard-delete.** The executor runs soft-delete only. Hard-delete is
   per-message and lives behind
   ``cli.py delete --permanent --yes-really-permanent``.
6. **unsupported != failed.** If the backend says the op is unsupported
   (Graph without ``Mail.ReadWrite``), route past as a no-op with a clear
   message. Exit code 2.

Bucket → operation mapping
--------------------------
``safe-delete``       → soft-delete (move to ``Deleted Items``)
``resolved-ticket``   → move to ``Archive`` (created if missing)
``low-signal-client`` → move to ``Archive`` (created if missing)
``needs-action``      → preview only; no automated action
``keep``              → preview only; no automated action
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Bootstrap sys.path for both script-style and -m invocations.
_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "scripts"))

from outlook_tool import OperationResult, OutlookClient  # noqa: E402


# Max age of a triage report, in seconds. Beyond this we refuse to apply
# because AppleScript message ids become session-stale.
MAX_REPORT_AGE_SECONDS = 60 * 60  # 1 hour

# Bolt 6 Phase B.3: confidence floor for the unread-junk re-route gate.
# When ``apply_bucket`` runs with ``--unread-junk-rule`` on (default), any
# message whose LLM verdict is ``bucket=="junk"`` at confidence >= this
# threshold AND ``is_read == False`` will be re-routed to safe-delete
# from the softer archivable buckets (resolved-ticket / low-signal-client)
# or from keep. This mirrors Chris's explicit heuristic: "unread + junk
# = delete."
UNREAD_JUNK_CONFIDENCE_FLOOR = 0.85

# Buckets that are softened (archive-worthy) but can be upgraded to
# safe-delete when the unread-junk gate fires.
UNREAD_JUNK_SOURCE_BUCKETS = ("resolved-ticket", "low-signal-client", "keep")

# Per-bucket destination for move operations. ``None`` means the bucket is
# preview-only; the executor will not automatically act on it even with
# ``--apply``.
BUCKET_ACTIONS: Dict[str, Optional[str]] = {
    "safe-delete": "soft-delete",     # routes through OutlookClient.delete
    "resolved-ticket": "Archive",
    "low-signal-client": "Archive",
    "needs-action": None,
    "keep": None,
}


# ---------------------------------------------------------------------------
# Unread-junk re-route gate (Bolt 6 Phase B.3)
# ---------------------------------------------------------------------------

def apply_unread_junk_gate(
    report: Dict[str, Any],
    *,
    confidence_floor: float = UNREAD_JUNK_CONFIDENCE_FLOOR,
) -> Dict[str, Any]:
    """Re-route unread + high-confidence-junk messages to safe-delete.

    Chris's explicit heuristic: "unread + junk = delete." The rules path
    is deliberately conservative — a routine FYI from a client or a
    resolved-ticket notification goes to the archivable buckets, not to
    safe-delete, because pattern rules alone can't tell "junk I'll never
    read" from "FYI I may want later." The LLM reads the body and
    recovers that signal, and the gate here acts on it:

      - If ``bucket`` is in ``UNREAD_JUNK_SOURCE_BUCKETS``
      - AND ``is_read == False``
      - AND ``llm_verdict.bucket == "junk"``
      - AND ``llm_verdict.confidence >= confidence_floor``
      → route to ``safe-delete`` with
         ``matched_rule = "unread-junk-upgrade"`` and preserve the
         original rule in ``unread_junk_upgrade``.

    The function mutates a shallow copy of the report so callers can
    diff or log the before/after. The report's ``summary`` is recounted
    in place.
    """
    new_report = dict(report)
    new_messages: List[Dict[str, Any]] = []
    upgrades = 0
    for msg in report.get("messages", []):
        if _should_upgrade_unread_junk(msg, confidence_floor):
            upgraded = dict(msg)
            upgraded["unread_junk_upgrade"] = {
                "from_bucket": msg.get("bucket"),
                "from_rule": msg.get("matched_rule"),
                "llm_confidence": (
                    msg.get("llm_verdict") or {}
                ).get("confidence"),
            }
            upgraded["bucket"] = "safe-delete"
            upgraded["matched_rule"] = "unread-junk-upgrade"
            new_messages.append(upgraded)
            upgrades += 1
        else:
            new_messages.append(msg)

    new_report["messages"] = new_messages

    # Recount the summary — upgrades change bucket populations.
    summary = {b: 0 for b in (
        "safe-delete", "resolved-ticket", "low-signal-client",
        "needs-action", "keep",
    )}
    for m in new_messages:
        b = m.get("bucket")
        if b in summary:
            summary[b] += 1
    new_report["summary"] = summary
    new_report["unread_junk_upgrades"] = upgrades
    return new_report


def _should_upgrade_unread_junk(
    msg: Dict[str, Any],
    confidence_floor: float,
) -> bool:
    """Return True iff this message meets the unread-junk re-route gate.

    Guarded by: current bucket is archivable-or-keep, message is unread,
    LLM verdict is high-confidence junk.
    """
    if msg.get("bucket") not in UNREAD_JUNK_SOURCE_BUCKETS:
        return False
    if bool(msg.get("is_read", False)):
        return False
    verdict = msg.get("llm_verdict") or {}
    if verdict.get("bucket") != "junk":
        return False
    try:
        conf = float(verdict.get("confidence", 0.0))
    except (TypeError, ValueError):
        return False
    return conf >= confidence_floor


# ---------------------------------------------------------------------------
# Report loading + freshness
# ---------------------------------------------------------------------------

class ReportTooOldError(Exception):
    """Raised when the triage report is older than :data:`MAX_REPORT_AGE_SECONDS`."""


def load_report(path: Path, *, now: Optional[datetime] = None) -> Dict[str, Any]:
    """Load and validate a triage report.

    Raises :class:`ReportTooOldError` if the report is too old to safely
    apply. The age gate runs even in dry-run mode so the preview always
    surfaces the staleness warning before a human says "apply".
    """
    if not path.exists():
        raise FileNotFoundError(f"triage report not found: {path}")
    report = json.loads(path.read_text())
    generated_raw = report.get("generated_at")
    if not generated_raw:
        raise ValueError("triage report is missing 'generated_at'")
    try:
        generated = datetime.fromisoformat(generated_raw)
    except ValueError as e:
        raise ValueError(f"invalid 'generated_at' in report: {e}")

    reference = now or datetime.now()
    age_seconds = (reference - generated).total_seconds()
    if age_seconds > MAX_REPORT_AGE_SECONDS:
        age_minutes = int(age_seconds / 60)
        raise ReportTooOldError(
            f"REFUSING to apply: triage report is {age_minutes} min old "
            f"(max {MAX_REPORT_AGE_SECONDS // 60} min). AppleScript message "
            f"ids go stale across sessions — regenerate with "
            f"'python scripts/triage.py' and retry."
        )

    # Stash computed age so the preview can show it without re-doing math.
    report["_age_seconds"] = age_seconds
    return report


# ---------------------------------------------------------------------------
# Preview rendering
# ---------------------------------------------------------------------------

def _escape_md_cell(s: str, max_len: int = 80) -> str:
    """Make a string safe to embed in a Markdown table cell."""
    s = (s or "").replace("|", "\\|").replace("\n", " ").replace("\r", " ")
    if len(s) > max_len:
        s = s[: max_len - 1] + "…"
    return s


def render_preview(
    report: Dict[str, Any],
    *,
    bucket: str,
    report_path: Path,
    exclude_ids: Optional[set] = None,
) -> str:
    """Build the Markdown preview the operator reads before saying 'apply'."""
    exclude = exclude_ids or set()
    messages = [
        m for m in report.get("messages", [])
        if m.get("bucket") == bucket and m.get("message_id") not in exclude
    ]
    generated_at = report.get("generated_at", "unknown")
    age_seconds = report.get("_age_seconds", 0)
    age_minutes = int(age_seconds / 60)
    action = BUCKET_ACTIONS.get(bucket)
    action_note = (
        f"Action with --apply: {action}"
        if action is not None
        else "Preview-only bucket — --apply is a no-op for safety."
    )

    lines: List[str] = []
    lines.append(f"## Triage Preview — {bucket} bucket ({len(messages)} messages)")
    lines.append("")
    lines.append(
        f"Report: {report_path.as_posix()} "
        f"(generated: {generated_at}, age: {age_minutes} min)"
    )
    lines.append(f"Backend: {report.get('backend', 'unknown')}")
    lines.append(f"Sample folder: {report.get('sample_folder', 'Inbox')}, "
                 f"window: {report.get('window_days', '?')} days, "
                 f"sample size: {report.get('sample_size', len(messages))}")
    lines.append(f"{action_note}")
    lines.append("")
    if not messages:
        lines.append("*(no messages in this bucket)*")
        return "\n".join(lines) + "\n"

    lines.append("| # | Received | Sender | Subject | Matched rule |")
    lines.append("|---|---|---|---|---|")
    for i, m in enumerate(messages, 1):
        received = m.get("received_at", "")[:16]  # trim to minute
        sender = _escape_md_cell(m.get("sender_email", ""), 40)
        subject = _escape_md_cell(m.get("subject", ""), 72)
        rule = _escape_md_cell(m.get("matched_rule", ""), 32)
        lines.append(f"| {i} | {received} | {sender} | {subject} | {rule} |")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Fingerprint fallback — AppleScript session IDs go stale
# ---------------------------------------------------------------------------

def _fingerprint_lookup(
    client: OutlookClient,
    *,
    sender_email: str,
    subject: str,
    received_at: str,
    folder: str,
) -> Optional[str]:
    """Try to find a message by ``(sender_email, subject, received_date)``.

    Returns the backend-scoped message id if found uniquely, else ``None``.
    """
    # Narrow the search window to the message's received date (and +/- 1 day
    # to tolerate timezone drift).
    try:
        received_dt = datetime.fromisoformat(received_at)
    except ValueError:
        return None
    date_from = (received_dt - timedelta(days=1)).strftime("%Y-%m-%d")
    date_to = (received_dt + timedelta(days=1)).strftime("%Y-%m-%d")

    hits = client.search(
        date_from=date_from,
        date_to=date_to,
        sender_email=sender_email or None,
        folder=folder,
        max_results=50,
    )
    exact: List[Dict[str, Any]] = []
    for h in hits:
        if (h.get("subject") or "") != subject:
            continue
        if (h.get("sender_email") or "").lower() != (sender_email or "").lower():
            continue
        exact.append(h)
    if len(exact) == 1:
        return exact[0].get("id")
    # Zero or ambiguous → caller handles.
    return None


# ---------------------------------------------------------------------------
# Apply loop
# ---------------------------------------------------------------------------

class ApplyOutcome:
    """Aggregate result of an --apply run — used for the final report."""

    def __init__(self):
        self.verified_success: int = 0
        self.unverified_success: int = 0   # status==success but verified==False
        self.failed: int = 0
        self.unsupported: int = 0
        self.skipped_ambiguous: int = 0
        self.rows: List[Dict[str, Any]] = []   # per-message outcome

    def tally(self, msg: Dict[str, Any], result: Optional[OperationResult], note: str = "") -> None:
        if result is None:
            self.skipped_ambiguous += 1
            status_label = "skipped"
            verified = False
            backend = None
            error = note or "fingerprint lookup failed (ambiguous or missing)"
        elif result.status == "unsupported":
            self.unsupported += 1
            status_label = "unsupported"
            verified = False
            backend = result.backend
            error = result.error_detail or ""
        elif result.status == "success":
            if result.verified:
                self.verified_success += 1
                status_label = "OK"
            else:
                self.unverified_success += 1
                status_label = "UNVERIFIED-success"
            verified = result.verified
            backend = result.backend
            error = result.error_detail or ""
        else:
            self.failed += 1
            status_label = "failed"
            verified = result.verified
            backend = result.backend
            error = result.error_detail or ""

        self.rows.append({
            "message_id": msg.get("message_id", ""),
            "sender_email": msg.get("sender_email", ""),
            "subject": msg.get("subject", ""),
            "matched_rule": msg.get("matched_rule", ""),
            "status": status_label,
            "verified": verified,
            "backend": backend,
            "error": error,
            "elapsed_ms": result.elapsed_ms if result is not None else None,
        })


def _apply_one(
    client: OutlookClient,
    msg: Dict[str, Any],
    *,
    action: str,
    source_folder: str,
) -> Tuple[Optional[OperationResult], str]:
    """Run the appropriate write verb for a single message.

    Returns ``(OperationResult_or_None, note)``. ``None`` means the
    fingerprint fallback couldn't resolve the message uniquely — the apply
    loop records it as skipped-ambiguous.
    """
    message_id = msg.get("message_id") or ""

    def _do(mid: str) -> OperationResult:
        if action == "soft-delete":
            return client.delete(mid, permanent=False, folder_name=source_folder)
        # move-to-folder
        return client.move(mid, action, source_folder=source_folder)

    # First attempt — use the id from the report.
    try:
        result = _do(message_id)
    except Exception as e:
        return None, f"exception on first attempt: {e}"

    # If the first attempt failed due to a stale/missing id, try a
    # fingerprint lookup against the source folder.
    stale_signal = (
        result.status == "failed"
        and result.error_detail
        and (
            "message_not_found" in result.error_detail
            or "pre-fetch failed" in result.error_detail
            or "not resolvable" in result.error_detail
        )
    )
    if not stale_signal:
        return result, ""

    fresh_id = _fingerprint_lookup(
        client,
        sender_email=msg.get("sender_email", ""),
        subject=msg.get("subject", ""),
        received_at=msg.get("received_at", ""),
        folder=source_folder,
    )
    if not fresh_id:
        return None, (
            f"stale id {message_id!r}, fingerprint lookup ambiguous or empty"
        )

    try:
        return _do(fresh_id), f"recovered via fingerprint (was {message_id})"
    except Exception as e:
        return None, f"exception on fingerprint retry: {e}"


def _apply_bulk(
    client: OutlookClient,
    messages: List[Dict[str, Any]],
    *,
    action: str,
    source_folder: str,
    outcome: ApplyOutcome,
    progress: bool,
) -> None:
    """Fire one bulk-move/delete call for all ``messages`` and tally.

    Bolt 5 (B5.2) bulk path. All messages share ``(source_folder, action)``.
    AppleScript ids go stale within a single session only when another
    process mutates the Inbox underneath us; the bulk primitive captures
    all refs before any move fires inside the same tell block, so the
    stale-id fingerprint fallback is NOT needed here. If a row fails for
    any reason, the caller can retry it through the per-message path.
    """
    if progress:
        tgt_label = "Deleted Items" if action == "soft-delete" else action
        print(
            f"  Bulk {action} — {len(messages)} message(s) → {tgt_label}",
            file=sys.stderr,
        )

    message_ids = [m.get("message_id") or "" for m in messages]
    try:
        if action == "soft-delete":
            results = client.delete_bulk(
                message_ids,
                permanent=False,
                folder_name=source_folder,
            )
        else:
            moves = [(mid, action) for mid in message_ids]
            results = client.move_bulk(moves, source_folder=source_folder)
    except Exception as e:
        # Backend blew up — mark every row as skipped-ambiguous with the
        # exception message so the operator can investigate.
        note = f"bulk {action} raised: {e!r}"
        for m in messages:
            outcome.tally(m, None, note=note)
        return

    # Zip results back to the messages. Lengths should match (the primitive
    # contracts this), but guard anyway so a bug on either side doesn't
    # scramble the audit log.
    if len(results) != len(messages):
        note = (
            f"bulk {action} returned {len(results)} results for "
            f"{len(messages)} messages — tallying by position"
        )
        print(f"WARNING: {note}", file=sys.stderr)

    for m, r in zip(messages, results):
        outcome.tally(m, r, note="")


def apply_bucket(
    client: OutlookClient,
    report: Dict[str, Any],
    *,
    bucket: str,
    source_folder: str = "Inbox",
    progress: bool = True,
    exclude_ids: Optional[set] = None,
    bulk: bool = True,
) -> ApplyOutcome:
    """Run the appropriate action on every message in ``bucket``.

    ``safe-delete`` → soft-delete via :meth:`OutlookClient.delete`.
    ``resolved-ticket`` / ``low-signal-client`` → move into ``Archive``.
    Preview-only buckets (needs-action, keep) short-circuit cleanly.

    ``exclude_ids`` — optional set of ``message_id`` values to skip. Used by
    the human-reviewed walk-through to preserve rows flagged for keep even
    though the classifier bucketed them.

    ``bulk`` — when True (default under Bolt 5), uses the single-tell-block
    bulk primitive. When False, falls back to the Bolt 4 per-message loop
    with fingerprint-fallback recovery.
    """
    action = BUCKET_ACTIONS.get(bucket)
    outcome = ApplyOutcome()
    if action is None:
        print(
            f"Bucket {bucket!r} is preview-only — nothing to apply.",
            file=sys.stderr,
        )
        return outcome

    exclude = exclude_ids or set()
    messages = [
        m for m in report.get("messages", [])
        if m.get("bucket") == bucket and m.get("message_id") not in exclude
    ]
    if not messages:
        print(f"No messages in bucket {bucket!r}.", file=sys.stderr)
        return outcome
    if exclude:
        print(
            f"Excluding {len(exclude)} message id(s) per --exclude-ids.",
            file=sys.stderr,
        )

    if bulk and client.backend == "applescript":
        _apply_bulk(
            client, messages,
            action=action,
            source_folder=source_folder,
            outcome=outcome,
            progress=progress,
        )
        return outcome

    for i, m in enumerate(messages, 1):
        if progress:
            print(
                f"  [{i}/{len(messages)}] {m.get('sender_email', '')} — "
                f"{(m.get('subject') or '')[:60]}",
                file=sys.stderr,
            )
        result, note = _apply_one(
            client, m, action=action, source_folder=source_folder,
        )
        outcome.tally(m, result, note=note)

    return outcome


# ---------------------------------------------------------------------------
# Rendering the apply summary
# ---------------------------------------------------------------------------

def render_apply_summary(outcome: ApplyOutcome, *, bucket: str, action: str) -> str:
    lines: List[str] = []
    total = (
        outcome.verified_success
        + outcome.unverified_success
        + outcome.failed
        + outcome.unsupported
        + outcome.skipped_ambiguous
    )
    lines.append(f"## Apply Result — {bucket} → {action} ({total} attempted)")
    lines.append("")
    lines.append(f"- verified-success: {outcome.verified_success}")
    lines.append(f"- UNVERIFIED-success (silent-success risk): {outcome.unverified_success}")
    lines.append(f"- failed: {outcome.failed}")
    lines.append(f"- unsupported: {outcome.unsupported}")
    lines.append(f"- skipped (ambiguous fingerprint): {outcome.skipped_ambiguous}")
    lines.append("")
    if outcome.unverified_success:
        lines.append(
            "WARNING: unverified-success rows need manual spot-checking. "
            "The write returned OK but the post-fetch verify did not "
            "confirm the state change. This is Phase 1's silent-success "
            "guard surfacing noise — treat it as a soft failure."
        )
        lines.append("")
    if outcome.failed or outcome.unverified_success:
        lines.append("### Rows needing attention")
        lines.append("")
        lines.append("| # | Status | Sender | Subject | Error |")
        lines.append("|---|---|---|---|---|")
        i = 0
        for row in outcome.rows:
            if row["status"] in ("OK", "unsupported"):
                continue
            i += 1
            lines.append(
                f"| {i} | {row['status']}"
                f"{' (verified)' if row['verified'] else ''} "
                f"| {_escape_md_cell(row['sender_email'], 35)} "
                f"| {_escape_md_cell(row['subject'], 60)} "
                f"| {_escape_md_cell(row.get('error', ''), 60)} |"
            )
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="apply_triage",
        description=(
            "Apply a triage report. Preview by default; --apply to execute. "
            "Refuses reports older than 1 hour. Branches on verified, "
            "not status — silent-success is surfaced loudly."
        ),
    )
    parser.add_argument(
        "report", nargs="?",
        help="Path to the triage report JSON. If omitted, --report is required.",
    )
    parser.add_argument(
        "--report", dest="report_flag",
        help="Alternative long form of the positional report path.",
    )
    parser.add_argument(
        "--bucket", choices=list(BUCKET_ACTIONS.keys()),
        default="safe-delete",
        help="Which bucket to preview or apply (default: safe-delete).",
    )
    parser.add_argument(
        "--source-folder", default="Inbox",
        help="Folder messages currently live in (default: Inbox).",
    )
    parser.add_argument(
        "--apply", action="store_true",
        help="Actually perform the action. Omit for Markdown preview only.",
    )
    parser.add_argument(
        "--output", metavar="PATH",
        help="Write the preview / apply summary to PATH in addition to stdout.",
    )
    parser.add_argument(
        "--exclude-ids", default="",
        help=(
            "Comma-separated list of message_ids to skip (e.g. "
            "'as_msg_62,as_msg_104'). Used after a human walk-through "
            "to preserve classifier-bucketed messages that the operator "
            "wants to keep. Applies to both preview and --apply."
        ),
    )
    parser.add_argument(
        "--bulk", dest="bulk", action="store_true",
        default=True,
        help=(
            "Use the Bolt 5 bulk-move primitive (default). Issues ONE "
            "AppleScript tell block for the entire bucket — much faster "
            "and avoids the sync-lag false-failure mode that the "
            "one-at-a-time path hits."
        ),
    )
    parser.add_argument(
        "--no-bulk", dest="bulk", action="store_false",
        help=(
            "Opt out of the bulk primitive and fall back to the Bolt 4 "
            "per-message loop with fingerprint-fallback recovery."
        ),
    )
    parser.add_argument(
        "--unread-junk-rule", dest="unread_junk_rule", action="store_true",
        default=True,
        help=(
            "Enable the Bolt 6 Phase B.3 unread-junk re-route gate "
            "(default on). Any message in resolved-ticket / "
            "low-signal-client / keep that is unread AND carries an "
            "LLM verdict of 'junk' at confidence >= 0.85 is upgraded "
            "to safe-delete before the applier fires. Mirrors Chris's "
            "'unread + junk = delete' heuristic."
        ),
    )
    parser.add_argument(
        "--no-unread-junk-rule", dest="unread_junk_rule", action="store_false",
        help="Opt out of the unread-junk re-route gate.",
    )
    parser.add_argument(
        "--unread-junk-confidence-floor", type=float,
        default=UNREAD_JUNK_CONFIDENCE_FLOOR,
        help=(
            "LLM confidence floor for the unread-junk gate "
            f"(default: {UNREAD_JUNK_CONFIDENCE_FLOOR})."
        ),
    )
    parser.add_argument(
        "--yes-really-permanent", action="store_true",
        help=argparse.SUPPRESS,  # Documented NO. Catch & refuse.
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # Panel-mandated refusal — the executor never offers hard-delete.
    if args.yes_really_permanent:
        print(
            "REFUSED: --yes-really-permanent is not available via the "
            "triage executor. Hard-delete is per-message only, via "
            "'python cli.py delete --permanent --yes-really-permanent'.",
            file=sys.stderr,
        )
        return 2

    report_path = args.report or args.report_flag
    if not report_path:
        parser.error("report path is required (positional or --report)")
    report_path = Path(report_path).expanduser().resolve()

    try:
        report = load_report(report_path)
    except ReportTooOldError as e:
        print(str(e), file=sys.stderr)
        return 2
    except (FileNotFoundError, ValueError) as e:
        print(f"REFUSED: {e}", file=sys.stderr)
        return 2

    # Bolt 6 Phase B.3 — apply the unread-junk re-route gate BEFORE any
    # preview or write. The gate is a pure data transform; running it
    # pre-preview means the operator sees the upgraded buckets in the
    # Markdown preview and can spot-check before --apply.
    if args.unread_junk_rule:
        report = apply_unread_junk_gate(
            report,
            confidence_floor=args.unread_junk_confidence_floor,
        )
        upgrades = report.get("unread_junk_upgrades", 0)
        if upgrades:
            print(
                f"unread-junk gate: upgraded {upgrades} message(s) to "
                f"safe-delete (confidence floor "
                f"{args.unread_junk_confidence_floor}).",
                file=sys.stderr,
            )

    exclude_ids = {
        s.strip() for s in (args.exclude_ids or "").split(",") if s.strip()
    }
    preview = render_preview(
        report, bucket=args.bucket, report_path=report_path,
        exclude_ids=exclude_ids,
    )
    print(preview)
    if exclude_ids:
        print(
            f"\n(Excluding {len(exclude_ids)} message_id(s) from the apply: "
            f"{', '.join(sorted(exclude_ids))})",
            file=sys.stderr,
        )

    if not args.apply:
        print(
            f"\nDRY RUN — re-run with --apply to execute "
            f"({args.bucket} bucket).",
            file=sys.stderr,
        )
        if args.output:
            Path(args.output).write_text(preview)
        return 0

    action = BUCKET_ACTIONS.get(args.bucket)
    if action is None:
        print(
            f"Bucket {args.bucket!r} is preview-only. "
            f"No write will fire even with --apply.",
            file=sys.stderr,
        )
        if args.output:
            Path(args.output).write_text(preview)
        return 0

    client = OutlookClient()
    bulk_label = (
        "bulk (Bolt 5)" if args.bulk and client.backend == "applescript"
        else "per-message (Bolt 4)"
    )
    print(
        f"\nApplying {args.bucket} → {action} via {client.backend} "
        f"backend — {bulk_label}…",
        file=sys.stderr,
    )

    outcome = apply_bucket(
        client, report,
        bucket=args.bucket,
        source_folder=args.source_folder,
        exclude_ids=exclude_ids,
        bulk=args.bulk,
    )
    summary = render_apply_summary(outcome, bucket=args.bucket, action=action)
    print(summary)

    if args.output:
        Path(args.output).write_text(preview + "\n\n" + summary)

    # Exit policy mirrors the CLI verbs:
    #   verified-success only (zero other kinds) → 0
    #   unsupported + nothing else           → 2
    #   anything else                        → 1
    if outcome.failed or outcome.skipped_ambiguous or outcome.unverified_success:
        return 1
    if outcome.unsupported and not outcome.verified_success:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
