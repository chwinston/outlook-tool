#!/usr/bin/env python3
"""What's actually in the Inbox — age, read-state, sender-count, internal-vs-external."""
from __future__ import annotations
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from outlook_tool import OutlookClient


def main() -> int:
    c = OutlookClient()
    print("Scanning full Inbox...", file=sys.stderr)
    inbox = c.search(max_results=5000, folder="Inbox",
                     date_from="2010-01-01", date_to="2026-11-19")
    print(f"Total: {len(inbox)}", file=sys.stderr)

    now = datetime.now()
    age_buckets = defaultdict(int)
    read_by_age = defaultdict(lambda: {"read": 0, "unread": 0})

    for m in inbox:
        rdt = m.get("received_datetime")
        if isinstance(rdt, str):
            try:
                dt = datetime.fromisoformat(rdt.replace("Z", "+00:00")).replace(tzinfo=None)
            except Exception:
                continue
        elif isinstance(rdt, datetime):
            dt = rdt
        else:
            continue

        days = (now - dt).days
        if days < 0:
            bucket = "future"
        elif days < 7:
            bucket = "<7d"
        elif days < 30:
            bucket = "7-30d"
        elif days < 90:
            bucket = "30-90d"
        elif days < 365:
            bucket = "90-365d"
        elif days < 730:
            bucket = "1-2yr"
        elif days < 1825:
            bucket = "2-5yr"
        else:
            bucket = ">5yr"

        age_buckets[bucket] += 1
        is_read = m.get("is_read", True)
        read_by_age[bucket]["read" if is_read else "unread"] += 1

    print("\n=== Age distribution ===")
    order = ["<7d", "7-30d", "30-90d", "90-365d", "1-2yr", "2-5yr", ">5yr", "future"]
    for k in order:
        n = age_buckets.get(k, 0)
        if n == 0:
            continue
        read = read_by_age[k]["read"]
        unread = read_by_age[k]["unread"]
        print(f"  {k:10}  {n:4}  (read: {read}, unread: {unread})")

    # Sender concentration
    sender_counts = Counter((m.get("sender_email", "") or "").lower() for m in inbox)
    top20 = sender_counts.most_common(20)
    total_top20 = sum(n for _, n in top20)

    print(f"\n=== Top 20 senders ({total_top20} msgs = {100*total_top20/len(inbox):.0f}% of Inbox) ===")
    for s, n in top20:
        print(f"  {n:4}  {s}")

    # Long tail
    one_offs = sum(1 for _, n in sender_counts.items() if n == 1)
    two_to_four = sum(1 for _, n in sender_counts.items() if 2 <= n <= 4)
    five_plus = sum(1 for _, n in sender_counts.items() if n >= 5)
    print(f"\n=== Sender concentration ===")
    print(f"  Unique senders: {len(sender_counts)}")
    print(f"  1-message senders: {one_offs}")
    print(f"  2-4 message senders: {two_to_four}")
    print(f"  5+ message senders: {five_plus}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
