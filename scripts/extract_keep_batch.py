#!/usr/bin/env python3
"""Extract a batch of `keep`-bucket messages for in-session LLM classification.

The /inbox-triage flow originally expected the Haiku API path. This script is
the "classify inside the Claude Code session" alternative: it emits a slim JSON
batch of N messages (id, sender, subject, body_preview) so Chris's co-pilot can
produce verdicts and feed them back via ``apply_llm_verdicts.py``.

Usage::

    python scripts/extract_keep_batch.py \\
        --report output/triage-YYYYMMDD-HHMM.json \\
        --batch-size 100 \\
        --offset 0 \\
        --output output/llm-batch-001.json

Then the assistant classifies by writing ``output/llm-verdicts-001.json``
in the same shape, and ``apply_llm_verdicts.py`` merges verdicts back into
the triage report before apply.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def main() -> int:
    p = argparse.ArgumentParser(prog="extract_keep_batch")
    p.add_argument("--report", required=True)
    p.add_argument("--batch-size", type=int, default=100)
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--output", required=True)
    p.add_argument(
        "--max-body",
        type=int,
        default=500,
        help="Truncate body_preview to this length (default 500 chars).",
    )
    args = p.parse_args()

    report_path = Path(args.report).expanduser().resolve()
    data = json.loads(report_path.read_text())
    keep = [m for m in data.get("messages", []) if m.get("bucket") == "keep"]
    # Only include messages with a non-trivial body preview — the LLM needs
    # content to judge intent on.
    eligible = [m for m in keep if len((m.get("body_preview") or "").strip()) >= 40]

    start = args.offset
    end = min(start + args.batch_size, len(eligible))
    batch = eligible[start:end]

    slim = []
    for m in batch:
        slim.append({
            "id": m.get("id") or m.get("message_id"),
            "sender_email": m.get("sender_email", ""),
            "sender_name": m.get("sender_name", ""),
            "subject": m.get("subject", ""),
            "body_preview": (m.get("body_preview") or "")[: args.max_body],
            "received_at": m.get("received_at", "") or m.get("received_datetime", ""),
            "is_read": m.get("is_read", True),
        })

    out = {
        "report_path": str(report_path),
        "offset": start,
        "batch_size": len(slim),
        "total_eligible": len(eligible),
        "messages": slim,
    }
    Path(args.output).write_text(json.dumps(out, indent=2, default=str))
    print(f"Wrote batch {start}-{end} ({len(slim)} messages) → {args.output}",
          file=sys.stderr)
    print(f"Eligible total: {len(eligible)}", file=sys.stderr)
    print(f"Next offset: {end}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
