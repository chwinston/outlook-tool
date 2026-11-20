#!/usr/bin/env python3
"""Merge in-session LLM verdicts back into a triage report.

Takes a verdicts JSON (shape produced by the assistant mid-session) and merges
``llm_verdict`` onto each matching message in the original triage report, so
``apply_triage.py --unread-junk-rule`` can consume it exactly like a Haiku run.

Verdict JSON shape::

    {
      "verdicts": [
        {"id": "as_msg_42", "bucket": "junk", "confidence": 0.9, "reason": "..."},
        ...
      ]
    }

``bucket`` ∈ {"junk", "reference", "action", "unclear"}
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_VALID_BUCKETS = {"junk", "reference", "action", "unclear"}


def main() -> int:
    p = argparse.ArgumentParser(prog="apply_llm_verdicts")
    p.add_argument("--report", required=True, help="Triage report path to update.")
    p.add_argument(
        "--verdicts",
        required=True,
        nargs="+",
        help="One or more verdicts JSON files (merged in order).",
    )
    p.add_argument(
        "--output",
        default=None,
        help="Write merged report here. Defaults to overwriting --report.",
    )
    p.add_argument(
        "--route-floor",
        type=float,
        default=0.85,
        help="Verdicts with confidence below this get route='needs-action'.",
    )
    args = p.parse_args()

    report_path = Path(args.report).expanduser().resolve()
    report = json.loads(report_path.read_text())
    msg_by_id = {m.get("id") or m.get("message_id"): m for m in report.get("messages", [])}

    merged = 0
    missing = 0
    bad_bucket = 0

    for v_path in args.verdicts:
        data = json.loads(Path(v_path).read_text())
        for v in data.get("verdicts", []):
            mid = v.get("id")
            if not mid or mid not in msg_by_id:
                missing += 1
                continue
            bucket = v.get("bucket")
            if bucket not in _VALID_BUCKETS:
                bad_bucket += 1
                continue
            confidence = float(v.get("confidence", 0.0))
            reason = v.get("reason", "")
            route = "auto" if confidence >= args.route_floor else "needs-action"
            msg_by_id[mid]["llm_verdict"] = {
                "bucket": bucket,
                "confidence": confidence,
                "reason": reason,
                "route": route,
            }
            merged += 1

    out_path = Path(args.output).expanduser().resolve() if args.output else report_path
    out_path.write_text(json.dumps(report, indent=2, default=str))

    print(f"Merged verdicts: {merged}", file=sys.stderr)
    if missing:
        print(f"Verdicts for unknown message_ids: {missing}", file=sys.stderr)
    if bad_bucket:
        print(f"Verdicts with invalid bucket: {bad_bucket}", file=sys.stderr)
    print(f"Wrote: {out_path}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
