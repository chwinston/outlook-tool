#!/usr/bin/env python3
"""Sweep a target folder by sender email — bulk delete or archive.

Fast lane for whole-sender cleanup. Bypasses the classifier: you specify
which senders go to which action, this script bulk-moves them.

Safety invariants (same as /inbox-triage):
1. Dry-run by default. --apply required for writes.
2. Archive uses move-to-folder; delete uses soft-delete (Deleted Items).
3. Uses Bolt 5 bulk primitive (OutlookClient.move_bulk).
4. Exemption rules from scripts/triage_rules.py are honored — any message
   that classifies to needs-action is EXCLUDED from the sweep.
5. --exclude-ids for message-level overrides.

Usage:
    # Preview
    python scripts/sweep_senders.py --config sweep-config.json

    # Execute
    python scripts/sweep_senders.py --config sweep-config.json --apply

Config format (JSON):
    {
      "delete": ["no-reply@sns.amazonaws.com", "sa@ahrefs.com", ...],
      "archive": ["noreply@notifications.hubspot.com", ...]
    }
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple

_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "scripts"))

from outlook_tool import OutlookClient  # noqa: E402
import triage_rules  # noqa: E402


def load_config(path: Path) -> Dict[str, List[str]]:
    if not path.exists():
        raise FileNotFoundError(f"config not found: {path}")
    data = json.loads(path.read_text())
    out = {
        "delete": [s.lower() for s in data.get("delete", [])],
        "archive": [s.lower() for s in data.get("archive", [])],
    }
    overlap = set(out["delete"]) & set(out["archive"])
    if overlap:
        raise ValueError(f"senders in both delete and archive: {sorted(overlap)}")
    return out


def scan_matches(
    client: OutlookClient,
    *,
    source_folder: str,
    senders: Set[str],
    date_from: str,
    date_to: str,
) -> List[Dict[str, Any]]:
    """Return all messages whose sender_email is in ``senders``."""
    if not senders:
        return []
    hits = client.search(
        max_results=10000,
        folder=source_folder,
        date_from=date_from,
        date_to=date_to,
    )
    out = []
    for h in hits:
        s = (h.get("sender_email") or "").lower()
        if s in senders:
            out.append(h)
    return out


def classify_as_needs_action(msg: Dict[str, Any]) -> bool:
    """Check if the message classifies to needs-action via triage_rules.

    Exemption rules (GCP at-risk, Google Account recoverability, MSI billing,
    SharePoint permission changes) route to needs-action and must not be
    swept.
    """
    # triage_rules.classify expects received_datetime; search returns received_at
    m = dict(msg)
    m.setdefault("received_datetime", msg.get("received_at"))
    bucket, _ = triage_rules.classify(m)
    return bucket == "needs-action"


def render_preview(
    matches: Dict[str, List[Dict[str, Any]]],
    excluded_exempt: List[Dict[str, Any]],
    excluded_ids: Set[str],
) -> str:
    """Render markdown preview grouped by action and sender."""
    lines: List[str] = []
    lines.append("# Sender Sweep — Preview")
    lines.append("")
    for action in ("delete", "archive"):
        bucket = matches.get(action, [])
        action_label = "Soft-delete (Deleted Items)" if action == "delete" else "Archive"
        lines.append(f"## {action_label} — {len(bucket)} messages")
        lines.append("")
        if not bucket:
            lines.append("*(none)*")
            lines.append("")
            continue
        # group by sender
        by_sender: Dict[str, List[Dict[str, Any]]] = {}
        for m in bucket:
            s = (m.get("sender_email") or "").lower()
            by_sender.setdefault(s, []).append(m)
        for s in sorted(by_sender, key=lambda k: -len(by_sender[k])):
            group = by_sender[s]
            lines.append(f"### `{s}` — {len(group)}")
            lines.append("")
        lines.append("")
    if excluded_exempt:
        lines.append(f"## Excluded — {len(excluded_exempt)} msgs matched exemption rules (needs-action)")
        lines.append("")
        for m in excluded_exempt[:20]:
            s = (m.get("sender_email") or "").lower()
            lines.append(f"- `{s}` — {m.get('subject', '')[:80]}")
        if len(excluded_exempt) > 20:
            lines.append(f"- *(+{len(excluded_exempt) - 20} more)*")
        lines.append("")
    if excluded_ids:
        lines.append(f"## Excluded — {len(excluded_ids)} via --exclude-ids")
        lines.append("")
    return "\n".join(lines)


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="sweep_senders",
        description="Bulk delete or archive messages by sender email.",
    )
    parser.add_argument("--config", required=True, help="Path to sweep config JSON.")
    parser.add_argument(
        "--source-folder", default="Inbox",
        help="Folder to sweep (default: Inbox).",
    )
    parser.add_argument(
        "--date-from", default="2010-01-01",
        help="Earliest message date (default: 2010-01-01).",
    )
    parser.add_argument(
        "--date-to", default=None,
        help="Latest message date (default: today).",
    )
    parser.add_argument(
        "--archive-folder", default="Archive",
        help="Folder for archive action (default: Archive).",
    )
    parser.add_argument(
        "--apply", action="store_true",
        help="Execute. Omit for preview only.",
    )
    parser.add_argument(
        "--exclude-ids", default="",
        help="Comma-separated message_ids to skip.",
    )
    parser.add_argument(
        "--output", default=None,
        help="Write preview/summary to PATH in addition to stdout.",
    )
    args = parser.parse_args(argv)

    config_path = Path(args.config).expanduser().resolve()
    config = load_config(config_path)

    # Default date_to = today
    import datetime
    date_to = args.date_to or datetime.date.today().isoformat()

    client = OutlookClient()
    print(f"Scanning {args.source_folder} {args.date_from} → {date_to}…", file=sys.stderr)

    delete_senders = set(config["delete"])
    archive_senders = set(config["archive"])
    all_senders = delete_senders | archive_senders

    hits = scan_matches(
        client,
        source_folder=args.source_folder,
        senders=all_senders,
        date_from=args.date_from,
        date_to=date_to,
    )
    print(f"Found {len(hits)} total matches for configured senders.", file=sys.stderr)

    exclude_ids = {
        s.strip() for s in (args.exclude_ids or "").split(",") if s.strip()
    }

    # Separate exempt messages (needs-action per triage rules)
    delete_bucket: List[Dict[str, Any]] = []
    archive_bucket: List[Dict[str, Any]] = []
    excluded_exempt: List[Dict[str, Any]] = []

    for m in hits:
        mid = m.get("id") or ""
        if mid in exclude_ids:
            continue
        sender = (m.get("sender_email") or "").lower()
        if classify_as_needs_action(m):
            excluded_exempt.append(m)
            continue
        if sender in delete_senders:
            delete_bucket.append(m)
        elif sender in archive_senders:
            archive_bucket.append(m)

    matches = {"delete": delete_bucket, "archive": archive_bucket}
    preview = render_preview(matches, excluded_exempt, exclude_ids)
    print(preview)

    if args.output:
        Path(args.output).write_text(preview)

    if not args.apply:
        print(f"\nDRY RUN — re-run with --apply to execute.", file=sys.stderr)
        return 0

    # ================ APPLY ================
    print(file=sys.stderr)
    print(f"Executing sweep via {client.backend} backend…", file=sys.stderr)

    summary_lines: List[str] = []
    rc = 0

    # CRITICAL: Run delete BEFORE archive (or any 2nd action) because
    # AppleScript message indices shift after each move. If both actions
    # are present, the second one needs a fresh scan.
    #
    # Simple fix: process delete first, then RE-SCAN for archive.
    action_order = ("delete", "archive")

    for action_idx, action in enumerate(action_order):
        bucket = matches.get(action, [])
        if not bucket:
            continue

        # Re-scan for any action after the first — prior actions shifted indices.
        if action_idx > 0:
            print(f"\nRe-scanning Inbox for {action} action (prior action shifted indices)…", file=sys.stderr)
            fresh_hits = scan_matches(
                client,
                source_folder=args.source_folder,
                senders=(delete_senders if action == "delete" else archive_senders),
                date_from=args.date_from,
                date_to=date_to,
            )
            # Re-filter exemptions
            bucket = []
            for m in fresh_hits:
                mid = m.get("id") or ""
                if mid in exclude_ids:
                    continue
                if classify_as_needs_action(m):
                    continue
                bucket.append(m)
            print(f"  Fresh bucket: {len(bucket)} messages", file=sys.stderr)

        action_label = "soft-delete" if action == "delete" else f"archive → {args.archive_folder}"
        print(f"\n{action_label}: {len(bucket)} messages", file=sys.stderr)

        # Resolve message indices (AppleScript index-based)
        ids_or_indices: List[int] = []
        for m in bucket:
            mid = m.get("id") or ""
            if not mid.startswith("as_msg_"):
                print(f"  WARN: skipping non-AppleScript id {mid}", file=sys.stderr)
                continue
            try:
                ids_or_indices.append(int(mid.split("_")[-1]))
            except ValueError:
                print(f"  WARN: skipping unparseable id {mid}", file=sys.stderr)

        if not ids_or_indices:
            continue

        # Use bulk dispatch.
        # delete_bulk expects list of message_id strings + folder_name kwarg.
        # move_bulk expects list of (int_index, target_folder) tuples + source_folder kwarg.
        if action == "delete":
            ids_as_strings = [f"as_msg_{idx}" for idx in ids_or_indices]
            results = client.delete_bulk(
                ids_as_strings,
                folder_name=args.source_folder,
                permanent=False,
            )
        else:
            moves = [(idx, args.archive_folder) for idx in ids_or_indices]
            results = client.move_bulk(
                moves,
                source_folder=args.source_folder,
            )

        verified = sum(1 for r in results if r.status == "success" and r.verified)
        unverified = sum(1 for r in results if r.status == "success" and not r.verified)
        failed = sum(1 for r in results if r.status == "failed")
        unsupported = sum(1 for r in results if r.status == "unsupported")

        line = (
            f"  {action_label}: "
            f"verified-success={verified}, "
            f"UNVERIFIED-success={unverified}, "
            f"failed={failed}, "
            f"unsupported={unsupported}"
        )
        print(line, file=sys.stderr)
        summary_lines.append(line)

        if failed:
            rc = 1

    if args.output and summary_lines:
        Path(args.output).write_text(
            preview + "\n\n## Apply Result\n\n" + "\n".join(summary_lines) + "\n"
        )

    return rc


if __name__ == "__main__":
    sys.exit(main())
