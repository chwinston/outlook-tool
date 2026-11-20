#!/usr/bin/env python3
"""Find Inbox clusters: identical subject templates from the same sender, N>=5."""
from __future__ import annotations
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from outlook_tool import OutlookClient


def normalize_subject(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"^((re|fwd?|aw):\s*)+", "", s, flags=re.I)
    s = re.sub(r"#\d+", "#N", s)
    s = re.sub(r"\[\d+\]", "[N]", s)
    s = re.sub(r"\$[\d,]+\.?\d*", "$N", s)
    s = re.sub(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", "DATE", s)
    s = re.sub(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", "DATE", s)
    s = re.sub(r"\b\d+\b", "N", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s[:80]


def main() -> int:
    c = OutlookClient()
    print("Scanning full Inbox...", file=sys.stderr)
    inbox = c.search(max_results=5000, folder="Inbox",
                     date_from="2010-01-01", date_to="2026-11-19")
    print(f"Total: {len(inbox)}", file=sys.stderr)

    groups: dict[tuple[str, str], list] = defaultdict(list)
    for m in inbox:
        sender = (m.get("sender_email", "") or "").lower()
        subj = normalize_subject(m.get("subject", ""))
        if not sender or not subj:
            continue
        groups[(sender, subj)].append(m)

    big = [(k, v) for k, v in groups.items() if len(v) >= 5]
    big.sort(key=lambda x: -len(x[1]))

    total_in_big = sum(len(v) for _, v in big)
    print(f"\nGroups with >= 5 identical msgs: {len(big)}")
    print(f"Messages in those groups: {total_in_big} "
          f"({100*total_in_big/len(inbox):.0f}% of Inbox)")
    print()
    print(f"{'N':>4}  {'sender':<40}  template")
    print("-" * 120)
    for (sender, template), msgs in big:
        print(f"{len(msgs):4}  {sender[:40]:<40}  {template[:70]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
