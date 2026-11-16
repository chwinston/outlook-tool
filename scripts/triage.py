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
from triage_rules import BUCKETS, classify  # noqa: E402


def _default_report_path(now: datetime) -> Path:
    stamp = now.strftime("%Y%m%d-%H%M")
    return _REPO_ROOT / "output" / f"triage-{stamp}.json"


def _message_record(msg: Dict[str, Any]) -> Dict[str, Any]:
    """Project an inbox message to the fields we persist in the report."""
    bucket, matched_rule = classify(msg)
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
        "bucket": bucket,
        "matched_rule": matched_rule,
    }


def classify_inbox(
    *,
    window_days: int = 30,
    max_results: int = 500,
    folder: str = "Inbox",
    backend: str | None = None,
) -> Dict[str, Any]:
    """Pull inbox messages, classify them, return the full report dict.

    Args:
        window_days: Only classify messages received within the last N days.
        max_results: Cap on search results.
        folder: Inbox-like folder to scan. Defaults to ``Inbox``.
        backend: Pass through to :class:`OutlookClient` (test hook).
    """
    client = OutlookClient(backend=backend) if backend else OutlookClient()
    now = datetime.now()
    since = now - timedelta(days=window_days)

    messages = client.search(
        date_from=since.strftime("%Y-%m-%d"),
        date_to=now.strftime("%Y-%m-%d"),
        folder=folder,
        max_results=max_results,
    )

    records: List[Dict[str, Any]] = [_message_record(m) for m in messages]
    summary = {b: 0 for b in BUCKETS}
    for r in records:
        summary[r["bucket"]] = summary.get(r["bucket"], 0) + 1

    return {
        "schema_version": 1,
        "generated_at": now.isoformat(),
        "backend": client.backend,
        "sample_folder": folder,
        "sample_size": len(records),
        "window_days": window_days,
        "summary": summary,
        "messages": records,
    }


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
    args = parser.parse_args(argv)

    report = classify_inbox(
        window_days=args.window_days,
        max_results=args.max_results,
        folder=args.folder,
        backend=args.backend,
    )

    out_path = Path(args.output) if args.output else _default_report_path(
        datetime.fromisoformat(report["generated_at"])
    )
    _write_report(report, out_path)

    if not args.quiet:
        print(f"Wrote {out_path}", file=sys.stderr)
        print(f"  backend: {report['backend']}", file=sys.stderr)
        print(f"  classified: {report['sample_size']} messages", file=sys.stderr)
        for bucket, count in report["summary"].items():
            print(f"    {bucket:<20s} {count:>4d}", file=sys.stderr)

    # Print report path on stdout so it's scriptable:
    print(out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
