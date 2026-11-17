#!/usr/bin/env python3
"""Classify inbox messages into triage buckets — Bolt 4 Phase 2 (B4.2.2).

Reads the inbox via :class:`outlook_tool.OutlookClient` (no subprocess round-
trip to ``cli.py``), applies :mod:`scripts.triage_rules` in order, and writes
a JSON report to ``output/triage-YYYYMMDD-HHMM.json``.

Report shape
------------
::

    {
      "schema_version": 1,
      "generated_at": "2026-11-16T16:40:00",
      "backend": "applescript",
      "sample_folder": "Inbox",
      "sample_size": 500,
      "window_days": 30,
      "summary": {
        "safe-delete": 142,
        "resolved-ticket": 38,
        "low-signal-client": 29,
        "needs-action": 61,
        "keep": 230
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
        },
        ...
      ]
    }

The executor (``scripts/apply_triage.py``) consumes this report by path.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

# Make ``outlook_tool`` and ``scripts.triage_rules`` importable regardless of
# whether you invoke via ``python scripts/triage.py`` or ``python -m scripts.triage``.
_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "scripts"))

from outlook_tool import OutlookClient  # noqa: E402
from triage_rules import BUCKETS, build_sent_context, classify  # noqa: E402


def _default_report_path(now: datetime) -> Path:
    stamp = now.strftime("%Y%m%d-%H%M")
    return _REPO_ROOT / "output" / f"triage-{stamp}.json"


def _message_record(
    msg: Dict[str, Any],
    context: Dict[str, Any] | None = None,
) -> Dict[str, Any]:
    """Project an inbox message to the fields we persist in the report.

    Bolt 6 Phase B.1: added ``body_preview`` to the projection (the LLM
    gate needs it) and thread ``context`` into :func:`classify` so the
    behavioral rules (ignored-sender-cold, dead-thread-reference, etc.)
    can read the Sent Items scan.
    """
    bucket, matched_rule = classify(msg, context)
    received = msg.get("received_datetime")
    if isinstance(received, datetime):
        received_iso = received.isoformat()
    else:
        received_iso = str(received) if received else ""

    return {
        "message_id": msg.get("id", ""),
        "sender_email": msg.get("sender_email", ""),
        "sender_name": msg.get("sender_name", ""),
        "subject": msg.get("subject", ""),
        "received_at": received_iso,
        "folder": msg.get("folder") or "Inbox",
        "is_read": bool(msg.get("is_read", False)),
        "body_preview": msg.get("body_preview", "") or "",
        "bucket": bucket,
        "matched_rule": matched_rule,
    }


def classify_inbox(
    *,
    window_days: int = 30,
    max_results: int = 500,
    folder: str = "Inbox",
    backend: str | None = None,
    sent_lookback_days: int = 365,
    sent_max_results: int = 5000,
    use_llm: bool = False,
    llm_limit: int = 200,
    llm_model: str | None = None,
    llm_confidence_threshold: float | None = None,
    llm_api_key: str | None = None,
) -> Dict[str, Any]:
    """Pull inbox messages, classify them, return the full report dict.

    Args:
        window_days: Only classify messages received within the last N days.
        max_results: Cap on search results.
        folder: Inbox-like folder to scan. Defaults to ``Inbox``.
        backend: Pass through to :class:`OutlookClient` (test hook).
        sent_lookback_days: How far back to scan Sent Items when building
            the per-run context (Bolt 6). Lets behavioral rules such as
            ``ignored-sender-cold`` detect "did Chris ever reply."
        sent_max_results: Upper bound on Sent Items scan size.
        use_llm: When True, run the LLM gate over every message the rules
            left in the ``keep`` bucket (Bolt 6 Phase B.1).
        llm_limit: Max messages to send to the LLM per run (cost cap).
        llm_model: Anthropic model ID. ``None`` means use the triage_llm
            default (Haiku).
        llm_confidence_threshold: Threshold above which the LLM's bucket
            is trusted for auto-action. ``None`` uses the triage_llm default.
        llm_api_key: Overrides ``ANTHROPIC_API_KEY`` env var.
    """
    client = OutlookClient(backend=backend) if backend else OutlookClient()
    now = datetime.now()
    since = now - timedelta(days=window_days)

    # Bolt 6 Phase B.1: build the per-run context ONCE. Behavioral rules
    # (ignored-sender-cold, dead-thread-reference, cold-outreach-dormant)
    # read ``sent_to_counts`` from this dict to answer "did Chris ever
    # reply?" without re-querying Outlook per message. We also compute
    # ``sender_counts`` from the live inbox pull so the volume floor on
    # ignored-sender-cold has something to clamp against.
    try:
        context: Dict[str, Any] = build_sent_context(
            client,
            days_lookback=sent_lookback_days,
            max_results=sent_max_results,
            now=now,
        )
    except Exception as exc:  # pragma: no cover — defensive
        print(
            f"warning: sent-items scan failed ({exc}) — "
            "behavioral rules will skip",
            file=sys.stderr,
        )
        context = {}

    messages = client.search(
        date_from=since.strftime("%Y-%m-%d"),
        date_to=now.strftime("%Y-%m-%d"),
        folder=folder,
        max_results=max_results,
    )

    # Build sender_counts from the inbox sample so ``ignored-sender-cold``
    # can enforce its volume floor. This is cheap (dict walk, no I/O).
    sender_counts: Dict[str, int] = {}
    for m in messages:
        s = (m.get("sender_email") or "").lower().strip()
        if not s:
            continue
        sender_counts[s] = sender_counts.get(s, 0) + 1
    context["sender_counts"] = sender_counts

    records: List[Dict[str, Any]] = [_message_record(m, context) for m in messages]

    # Optional LLM pass (Bolt 6 Phase B.1). Runs only on keep-bucket
    # messages via the triage_llm module. Additive — never overwrites
    # bucket / matched_rule, only attaches ``llm_verdict``.
    llm_summary: Dict[str, Any] | None = None
    if use_llm:
        # Lazy import so running triage without --llm doesn't require
        # the anthropic SDK installed.
        from triage_llm import (  # noqa: WPS433
            DEFAULT_CONFIDENCE_THRESHOLD,
            DEFAULT_MODEL,
            classify_messages_with_llm,
        )

        model = llm_model or DEFAULT_MODEL
        threshold = (
            llm_confidence_threshold
            if llm_confidence_threshold is not None
            else DEFAULT_CONFIDENCE_THRESHOLD
        )

        # Only the keep bucket needs LLM review — the rules already
        # bucketed everything else. We filter here (not inside the
        # classifier) so the summary line can report "classified N of M
        # keep messages" without the classifier having to know about
        # upstream counts.
        keep_msgs = [r for r in records if r.get("bucket") == "keep"]
        other_msgs = [r for r in records if r.get("bucket") != "keep"]

        augmented_keep = classify_messages_with_llm(
            keep_msgs,
            limit=llm_limit,
            model=model,
            confidence_threshold=threshold,
            api_key=llm_api_key,
        )
        verdicts_added = sum(1 for m in augmented_keep if "llm_verdict" in m)
        # Rebuild records preserving original order.
        aug_by_id = {m.get("message_id"): m for m in augmented_keep}
        records = [
            aug_by_id.get(r.get("message_id"), r) if r.get("bucket") == "keep"
            else r
            for r in records
        ]
        # Fallback: if any keep message's id was blank (shouldn't happen,
        # but defensive), fall back to the augmented list for keep rows.
        if len([r for r in records if r.get("bucket") == "keep"]) != len(augmented_keep):
            records = other_msgs + augmented_keep

        llm_summary = {
            "model": model,
            "limit": llm_limit,
            "confidence_threshold": threshold,
            "keep_candidates": len(keep_msgs),
            "verdicts_added": verdicts_added,
        }

    summary = {b: 0 for b in BUCKETS}
    for r in records:
        summary[r["bucket"]] = summary.get(r["bucket"], 0) + 1

    out: Dict[str, Any] = {
        "schema_version": 1,
        "generated_at": now.isoformat(),
        "backend": client.backend,
        "sample_folder": folder,
        "sample_size": len(records),
        "window_days": window_days,
        "sent_context": {
            "scanned_messages": context.get("scanned_messages", 0),
            "days_lookback": context.get("days_lookback", sent_lookback_days),
            "unique_recipients": len(context.get("sent_to_counts") or {}),
        },
        "summary": summary,
        "messages": records,
    }
    if llm_summary is not None:
        out["llm_summary"] = llm_summary
    return out


def _write_report(report: Dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2, default=str))


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="triage",
        description=(
            "Classify inbox messages into triage buckets. "
            "Produces output/triage-YYYYMMDD-HHMM.json for the executor."
        ),
    )
    parser.add_argument(
        "--window-days", type=int, default=30,
        help="Scan messages received in the last N days (default: 30)",
    )
    parser.add_argument(
        "--max-results", type=int, default=500,
        help="Maximum messages to classify (default: 500)",
    )
    parser.add_argument(
        "--folder", default="Inbox",
        help="Folder to scan (default: Inbox)",
    )
    parser.add_argument(
        "--backend", choices=["applescript", "win32com", "graph"],
        default=None,
        help="Force a specific backend (default: auto-detect)",
    )
    parser.add_argument(
        "--output", metavar="PATH",
        help=(
            "Report output path. Defaults to "
            "output/triage-YYYYMMDD-HHMM.json inside the repo."
        ),
    )
    parser.add_argument(
        "--quiet", action="store_true",
        help="Suppress the summary line on stdout.",
    )
    parser.add_argument(
        "--sent-lookback-days", type=int, default=365,
        help=(
            "How far back to scan Sent Items for the behavioral rules "
            "context (default: 365). The classifier reads this to decide "
            "'did Chris ever reply to this sender?'."
        ),
    )
    parser.add_argument(
        "--sent-max-results", type=int, default=5000,
        help="Upper bound on the Sent Items scan (default: 5000).",
    )
    parser.add_argument(
        "--llm", action="store_true",
        help=(
            "Run the Claude Haiku LLM gate over every message the rules "
            "left in 'keep'. Adds an 'llm_verdict' field to each "
            "classified message. Requires ANTHROPIC_API_KEY."
        ),
    )
    parser.add_argument(
        "--llm-limit", type=int, default=200,
        help="Max messages to send to the LLM per run (default: 200).",
    )
    parser.add_argument(
        "--llm-model", default=None,
        help="Anthropic model ID (default: triage_llm.DEFAULT_MODEL).",
    )
    parser.add_argument(
        "--llm-confidence-threshold", type=float, default=None,
        help=(
            "Confidence threshold above which the LLM bucket is trusted "
            "for auto-action (default: triage_llm.DEFAULT_CONFIDENCE_THRESHOLD)."
        ),
    )
    parser.add_argument(
        "--llm-api-key", default=None,
        help="Override ANTHROPIC_API_KEY env var for the LLM pass.",
    )
    args = parser.parse_args(argv)

    report = classify_inbox(
        window_days=args.window_days,
        max_results=args.max_results,
        folder=args.folder,
        backend=args.backend,
        sent_lookback_days=args.sent_lookback_days,
        sent_max_results=args.sent_max_results,
        use_llm=args.llm,
        llm_limit=args.llm_limit,
        llm_model=args.llm_model,
        llm_confidence_threshold=args.llm_confidence_threshold,
        llm_api_key=args.llm_api_key,
    )

    out_path = Path(args.output) if args.output else _default_report_path(
        datetime.fromisoformat(report["generated_at"])
    )
    _write_report(report, out_path)

    if not args.quiet:
        print(f"Wrote {out_path}", file=sys.stderr)
        print(f"  backend: {report['backend']}", file=sys.stderr)
        print(f"  classified: {report['sample_size']} messages", file=sys.stderr)
        sent_ctx = report.get("sent_context", {})
        if sent_ctx:
            print(
                f"  sent-context: {sent_ctx.get('scanned_messages', 0)} sent "
                f"messages scanned, {sent_ctx.get('unique_recipients', 0)} "
                f"unique recipients (lookback "
                f"{sent_ctx.get('days_lookback', '?')} days)",
                file=sys.stderr,
            )
        for bucket, count in report["summary"].items():
            print(f"    {bucket:<20s} {count:>4d}", file=sys.stderr)
        llm_summary = report.get("llm_summary")
        if llm_summary:
            print(
                f"  llm: {llm_summary.get('verdicts_added', 0)} verdicts "
                f"added over {llm_summary.get('keep_candidates', 0)} "
                f"keep candidates (model={llm_summary.get('model', '?')}, "
                f"limit={llm_summary.get('limit', '?')})",
                file=sys.stderr,
            )

    # Print report path on stdout so it's scriptable:
    print(out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
