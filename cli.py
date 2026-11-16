#!/usr/bin/env python3
"""
Outlook Tool CLI — Search, download attachments, and send emails from the terminal.

Usage:
    python cli.py search --from 2026-03-01 --to 2026-03-15 --subject "report"
    python cli.py search --domain example.com --has-attachments --download ./out
    python cli.py send --to user@example.com --subject "Hello" --body "Hi there"
    python cli.py send --to user@example.com --subject "Notes" --body "See attached" --attach notes.pdf
"""

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from outlook_tool import OperationResult, OutlookClient


def cmd_search(args):
    """Execute a search and optionally download attachments."""
    client = OutlookClient()

    kwargs = {}
    if args.date_from:
        kwargs["date_from"] = args.date_from
    if args.date_to:
        kwargs["date_to"] = args.date_to
    if args.subject:
        kwargs["subject_contains"] = args.subject
    if args.subject_regex:
        kwargs["subject_matches"] = args.subject_regex
    if args.sender_name:
        kwargs["sender_name"] = args.sender_name
    if args.sender_email:
        kwargs["sender_email"] = args.sender_email
    if args.domain:
        kwargs["sender_domain"] = args.domain
    if args.domains:
        kwargs["sender_domains"] = args.domains
    if args.has_attachments:
        kwargs["has_attachments"] = True
    if args.folder:
        kwargs["folder"] = args.folder
    if args.folders:
        kwargs["folders"] = args.folders
    if args.unread:
        kwargs["is_read"] = False
    if args.read:
        kwargs["is_read"] = True
    if args.body:
        kwargs["body_contains"] = args.body
    if args.to:
        kwargs["to_contains"] = args.to
    if args.importance:
        kwargs["importance"] = args.importance
    if args.max_results:
        kwargs["max_results"] = args.max_results

    results = client.search(**kwargs)

    if args.json:
        serializable = []
        for e in results:
            entry = {k: v for k, v in e.items() if not k.startswith("_")}
            entry["received_datetime"] = entry["received_datetime"].isoformat()
            serializable.append(entry)
        print(json.dumps(serializable, indent=2))
    else:
        print(f"\nFound {len(results)} email(s)\n")

        for i, email in enumerate(results, 1):
            att_count = len(email["attachments"])
            att_info = f" [{att_count} attachment(s)]" if att_count > 0 else ""
            print(f"  {i}. {email['received_date']} | {email['sender_name']} <{email['sender_email']}>")
            print(f"     Subject: {email['subject']}{att_info}")
            if email["attachments"]:
                for att in email["attachments"]:
                    size_kb = att.get("size", 0) / 1024
                    print(f"       -> {att['name']} ({size_kb:.0f} KB)")
            print()

        # Download attachments if requested
        if args.download and results:
            download_dir = Path(args.download)
            download_dir.mkdir(parents=True, exist_ok=True)
            count = 0
            for email in results:
                for att in email["attachments"]:
                    dest = client.download_attachment(email, att, output_dir=download_dir)
                    print(f"  Downloaded: {dest}")
                    count += 1
            print(f"\n  {count} attachment(s) downloaded to {download_dir}")

    return results


def cmd_events(args):
    """List calendar events."""
    client = OutlookClient()

    kwargs = {}
    if args.today:
        kwargs["date_from"] = datetime.now().strftime("%Y-%m-%d")
        kwargs["date_to"] = datetime.now().strftime("%Y-%m-%d")
    elif args.week:
        kwargs["date_from"] = datetime.now().strftime("%Y-%m-%d")
        kwargs["date_to"] = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
    else:
        if args.date_from:
            kwargs["date_from"] = args.date_from
        if args.date_to:
            kwargs["date_to"] = args.date_to
    if args.subject:
        kwargs["subject_contains"] = args.subject
    if args.max_results:
        kwargs["max_results"] = args.max_results

    results = client.get_events(**kwargs)

    if args.json:
        serializable = []
        for e in results:
            entry = {k: v for k, v in e.items() if not k.startswith("_")}
            entry["start_datetime"] = entry["start_datetime"].isoformat()
            entry["end_datetime"] = entry["end_datetime"].isoformat()
            serializable.append(entry)
        print(json.dumps(serializable, indent=2))
    else:
        print(f"\nFound {len(results)} event(s)\n")

        for i, evt in enumerate(results, 1):
            time_str = f"{evt['start_datetime'].strftime('%H:%M')}–{evt['end_datetime'].strftime('%H:%M')}"
            if evt["is_all_day"]:
                time_str = "All day"
            loc = f" @ {evt['location']}" if evt["location"] else ""
            org = f" (organized by {evt['organizer_name']})" if evt["organizer_name"] else ""

            print(f"  {i}. {evt['start_date']} {time_str}{loc}")
            print(f"     {evt['subject']}{org}")

            if evt["attendees"]:
                att_summary = ", ".join(
                    f"{a['name'] or a['email']} ({a['status']})"
                    for a in evt["attendees"][:5]
                )
                extra = f" +{len(evt['attendees']) - 5} more" if len(evt["attendees"]) > 5 else ""
                print(f"     Attendees: {att_summary}{extra}")

            if args.show_body and evt.get("body_preview", "").strip():
                body = evt["body_preview"].strip()
                # Collapse excessive blank lines
                while "\n\n\n" in body:
                    body = body.replace("\n\n\n", "\n\n")
                # Indent each line under the event
                for line in body.split("\n"):
                    print(f"     | {line}")

            print()

    return results


def cmd_summary(args):
    """Generate a structured chronological ledger of emails and calendar events."""
    client = OutlookClient()

    date_from = args.date_from or datetime.now().strftime("%Y-%m-%d")
    date_to = args.date_to or date_from

    # Fetch emails
    email_kwargs = {"date_from": date_from, "date_to": date_to, "max_results": 500}
    if args.folders:
        email_kwargs["folders"] = args.folders

    emails = client.search(**email_kwargs)

    # Fetch calendar events (unless --no-calendar)
    events = []
    if not args.no_calendar:
        events = client.get_events(date_from=date_from, date_to=date_to, max_results=200)

    # Build unified timeline
    entries = []
    for e in emails:
        entries.append({
            "kind": "email",
            "id": e.get("id", ""),
            "datetime": e["received_datetime"],
            "date": e["received_date"],
            "time": e["received_datetime"].strftime("%I:%M %p"),
            "day": e["received_datetime"].strftime("%a %b %d"),
            "subject": e["subject"],
            "from_name": e["sender_name"],
            "from_email": e["sender_email"],
            "to": e.get("to", ""),
            "folder": next(
                (a["_as_folder_name"] for a in e.get("attachments", []) if "_as_folder_name" in a),
                "Inbox",
            ),
            "has_attachments": e["has_attachments"],
            "body_preview": e.get("body_preview", "")[:500],
        })

    for ev in events:
        entries.append({
            "kind": "event",
            "id": ev.get("id", ""),
            "datetime": ev["start_datetime"],
            "date": ev["start_date"],
            "time": ev["start_datetime"].strftime("%I:%M %p"),
            "end_time": ev["end_datetime"].strftime("%I:%M %p"),
            "day": ev["start_datetime"].strftime("%a %b %d"),
            "subject": ev["subject"],
            "from_name": ev.get("organizer_name", ""),
            "from_email": ev.get("organizer_email", ""),
            "location": ev.get("location", ""),
            "is_all_day": ev.get("is_all_day", False),
            "status": ev.get("status", ""),
            "attendees": ev.get("attendees", []),
            "body_preview": ev.get("body_preview", "")[:500],
        })

    # Sort chronologically
    entries.sort(key=lambda x: x["datetime"])

    # Assign display IDs
    email_idx = 0
    event_idx = 0
    for entry in entries:
        if entry["kind"] == "email":
            email_idx += 1
            entry["display_id"] = f"E{email_idx:03d}"
        else:
            event_idx += 1
            entry["display_id"] = f"C{event_idx:03d}"

    # Output
    if args.format == "json":
        serializable = []
        for entry in entries:
            row = {k: v for k, v in entry.items() if k != "datetime"}
            row["datetime"] = entry["datetime"].isoformat()
            serializable.append(row)
        output = json.dumps({
            "schema_version": 1,
            "date_from": date_from,
            "date_to": date_to,
            "email_count": email_idx,
            "event_count": event_idx,
            "entries": serializable,
        }, indent=2)
    else:
        # Markdown ledger
        lines = [
            f"# Chronological Ledger: {date_from} to {date_to}",
            "",
            f"**Emails:** {email_idx} | **Calendar events:** {event_idx} | **Total:** {len(entries)}",
            "",
            "---",
            "",
            "| ID | Type | Date | Time | From | Subject |",
            "|----|------|------|------|------|---------|",
        ]
        for entry in entries:
            eid = entry["display_id"]
            kind = entry["kind"]
            day = entry["day"]
            time = entry["time"]
            if kind == "event" and entry.get("is_all_day"):
                time = "All day"
            elif kind == "event":
                time = f"{time}–{entry.get('end_time', '')}"
            from_name = entry["from_name"] or entry.get("from_email", "")
            subject = entry["subject"][:80]
            icon = "📧" if kind == "email" else "📅"
            lines.append(
                f"| <a id=\"{eid}\"></a>{eid} | {icon} | {day} | {time} | {from_name} | {subject} |"
            )
        lines.append("")
        output = "\n".join(lines)

    if args.output:
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(output)
        print(f"Summary written to {out_path}", file=sys.stderr)
    else:
        print(output)

    return entries


def _print_result(result: OperationResult, as_json: bool) -> None:
    """Render an OperationResult as JSON or a one-line human summary."""
    if as_json:
        print(json.dumps(result.to_dict(), indent=2, default=str))
        return
    marker = {
        "success": "OK",
        "failed": "FAIL",
        "unsupported": "UNSUPPORTED",
        "dry_run": "DRY-RUN",
    }.get(result.status, result.status)
    verified = "verified" if result.verified else "UNVERIFIED"
    detail = f" — {result.error_detail}" if result.error_detail else ""
    elapsed = f" ({result.elapsed_ms} ms)" if result.elapsed_ms is not None else ""
    print(
        f"[{marker}] {result.operation} msg={result.message_id} "
        f"backend={result.backend} ({verified}){elapsed}{detail}"
    )
    if result.before_state:
        print(f"  before: {json.dumps(result.before_state, default=str)}")
    if result.after_state:
        print(f"  after:  {json.dumps(result.after_state, default=str)}")


def _dry_run_result(operation: str, message_id: str, plan: dict) -> OperationResult:
    """Build a DRY-RUN OperationResult used when --apply is not set."""
    return OperationResult(
        status="dry_run",
        operation=operation,
        message_id=message_id,
        before_state={"plan": plan},
        backend="cli",
    )


def _exit_code_from(result: OperationResult) -> int:
    """CLI exit code policy: success/dry_run → 0, unsupported → 2, failed → 1."""
    if result.status in ("success", "dry_run"):
        return 0
    if result.status == "unsupported":
        return 2
    return 1


def cmd_move(args):
    """Move a message to ``target_folder``. Dry-run by default; --apply to fire."""
    plan = {
        "operation": "move",
        "message_id": args.message_id,
        "target_folder": args.target_folder,
        "source_folder": args.source_folder,
    }
    if not args.apply:
        result = _dry_run_result("move", args.message_id, plan)
        print(
            f"DRY RUN: would move message {args.message_id} to folder "
            f"{args.target_folder!r} (source={args.source_folder or 'Inbox'}). "
            f"Re-run with --apply to execute.",
            file=sys.stderr,
        )
        _print_result(result, args.json)
        return _exit_code_from(result)

    client = OutlookClient()
    result = client.move(
        args.message_id,
        args.target_folder,
        source_folder=args.source_folder,
    )
    _print_result(result, args.json)
    return _exit_code_from(result)


def cmd_delete(args):
    """Delete a message. Soft-delete by default; --permanent needs --yes-really-permanent."""
    plan = {
        "operation": "delete",
        "message_id": args.message_id,
        "permanent": args.permanent,
        "folder_name": args.folder,
    }
    if args.permanent and not args.yes_really_permanent:
        print(
            "REFUSED: --permanent requires --yes-really-permanent in the "
            "same invocation. Soft-delete (default) is almost always what "
            "you want.",
            file=sys.stderr,
        )
        return 2

    if not args.apply:
        kind = "HARD DELETE" if args.permanent else "SOFT DELETE"
        print(
            f"DRY RUN: would {kind} message {args.message_id} "
            f"(folder={args.folder or 'Inbox'}). Re-run with --apply to execute.",
            file=sys.stderr,
        )
        result = _dry_run_result("delete", args.message_id, plan)
        _print_result(result, args.json)
        return _exit_code_from(result)

    client = OutlookClient()
    try:
        result = client.delete(
            args.message_id,
            permanent=args.permanent,
            yes_really_permanent=args.yes_really_permanent,
            folder_name=args.folder,
        )
    except ValueError as e:
        # Belt-and-suspenders: client re-asserts the double-flag invariant.
        print(f"REFUSED: {e}", file=sys.stderr)
        return 2
    _print_result(result, args.json)
    return _exit_code_from(result)


def cmd_mark_read(args):
    """Mark a message read (default) or unread."""
    read = not args.unread  # default: mark as read
    plan = {
        "operation": "mark_read",
        "message_id": args.message_id,
        "read": read,
        "folder_name": args.folder,
    }
    if not args.apply:
        verb = "mark READ" if read else "mark UNREAD"
        print(
            f"DRY RUN: would {verb} message {args.message_id} "
            f"(folder={args.folder or 'Inbox'}). Re-run with --apply to execute.",
            file=sys.stderr,
        )
        result = _dry_run_result("mark_read", args.message_id, plan)
        _print_result(result, args.json)
        return _exit_code_from(result)

    client = OutlookClient()
    result = client.mark_read(
        args.message_id, read=read, folder_name=args.folder,
    )
    _print_result(result, args.json)
    return _exit_code_from(result)


def cmd_send(args):
    """Send an email."""
    client = OutlookClient()

    kwargs = {
        "to": args.to,
        "subject": args.subject,
        "body": args.body,
    }
    if args.cc:
        kwargs["cc"] = args.cc
    if args.bcc:
        kwargs["bcc"] = args.bcc
    if args.attach:
        kwargs["attachments"] = args.attach
    if args.html:
        kwargs["html"] = True
    if args.importance:
        kwargs["importance"] = args.importance

    client.send(**kwargs)
    print(f"  Email sent to {', '.join(args.to)}")


def main():
    parser = argparse.ArgumentParser(
        prog="outlook-tool",
        description="Search, download, and send Outlook emails from the terminal.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # ---- search ----
    sp = subparsers.add_parser("search", help="Search emails with flexible filters")
    sp.add_argument("--from", dest="date_from", help="Start date (YYYY-MM-DD)")
    sp.add_argument("--to-date", dest="date_to", help="End date (YYYY-MM-DD)")
    sp.add_argument("--subject", help="Subject contains (case-insensitive)")
    sp.add_argument("--subject-regex", help="Subject regex pattern")
    sp.add_argument("--sender-name", help="Sender display name contains")
    sp.add_argument("--sender-email", help="Sender email exact match")
    sp.add_argument("--domain", help="Sender domain filter")
    sp.add_argument("--domains", nargs="+", help="Multiple sender domains")
    sp.add_argument("--has-attachments", action="store_true", help="Only emails with attachments")
    sp.add_argument("--folder", help="Folder to search (default: Inbox)")
    sp.add_argument("--folders", nargs="+", help="Multiple folders to search (e.g., Inbox Archive Snoozed)")
    sp.add_argument("--unread", action="store_true", help="Only unread emails")
    sp.add_argument("--read", action="store_true", help="Only read emails")
    sp.add_argument("--body", help="Body contains (case-insensitive)")
    sp.add_argument("--to", dest="to", help="To field contains")
    sp.add_argument("--importance", choices=["high", "normal", "low"])
    sp.add_argument("--max-results", type=int, default=50, help="Max results (default 50)")
    sp.add_argument("--download", metavar="DIR", help="Download attachments to this directory")
    sp.add_argument("--json", action="store_true", help="Output results as JSON")
    sp.set_defaults(func=cmd_search)

    # ---- events ----
    sp = subparsers.add_parser("events", help="List calendar events")
    sp.add_argument("--from", dest="date_from", help="Start date (YYYY-MM-DD, default: today)")
    sp.add_argument("--to-date", dest="date_to", help="End date (YYYY-MM-DD, default: +7 days)")
    range_group = sp.add_mutually_exclusive_group()
    range_group.add_argument("--today", action="store_true", help="Show today's events only")
    range_group.add_argument("--week", action="store_true", help="Show this week's events")
    sp.add_argument("--subject", help="Subject contains (case-insensitive)")
    sp.add_argument("--max-results", type=int, default=50, help="Max results (default 50)")
    sp.add_argument("--body", dest="show_body", action="store_true", help="Show meeting notes/description")
    sp.add_argument("--json", action="store_true", help="Output results as JSON")
    sp.set_defaults(func=cmd_events)

    # ---- summary ----
    sp = subparsers.add_parser("summary", help="Generate chronological ledger of emails and events")
    sp.add_argument("--from", dest="date_from", help="Start date (YYYY-MM-DD, default: today)")
    sp.add_argument("--to-date", dest="date_to", help="End date (YYYY-MM-DD, default: same as --from)")
    sp.add_argument("--folders", nargs="+", help="Email folders to search (default: Inbox)")
    sp.add_argument("--no-calendar", action="store_true", help="Exclude calendar events")
    sp.add_argument("--format", choices=["json", "markdown"], default="json", help="Output format (default: json)")
    sp.add_argument("--output", metavar="FILE", help="Write to file instead of stdout")
    sp.set_defaults(func=cmd_summary)

    # ---- move ----
    sp = subparsers.add_parser(
        "move",
        help="Move a message to a folder (dry-run default; --apply to execute)",
    )
    sp.add_argument("message_id", help="Backend-scoped message id (e.g. 'as_msg_7' or an Outlook EntryID)")
    sp.add_argument("target_folder", help="Destination folder (created if missing on AppleScript/COM)")
    sp.add_argument("--source-folder", help="Source folder for AppleScript (default: Inbox)")
    sp.add_argument("--apply", action="store_true", help="Actually perform the move (otherwise dry-run)")
    sp.add_argument("--json", action="store_true", help="Emit the OperationResult as JSON")
    sp.set_defaults(func=cmd_move)

    # ---- delete ----
    sp = subparsers.add_parser(
        "delete",
        help="Delete a message — soft-delete default; --permanent + --yes-really-permanent for hard-delete",
    )
    sp.add_argument("message_id", help="Backend-scoped message id")
    sp.add_argument("--folder", help="Source folder for AppleScript (default: Inbox)")
    sp.add_argument(
        "--permanent", action="store_true",
        help="Hard-delete (bypass Deleted Items). Requires --yes-really-permanent.",
    )
    sp.add_argument(
        "--yes-really-permanent", action="store_true",
        help="Required alongside --permanent — the belt to the braces.",
    )
    sp.add_argument("--apply", action="store_true", help="Actually perform the delete (otherwise dry-run)")
    sp.add_argument("--json", action="store_true", help="Emit the OperationResult as JSON")
    sp.set_defaults(func=cmd_delete)

    # ---- mark-read ----
    sp = subparsers.add_parser(
        "mark-read",
        help="Mark a message as read (default) or unread (--unread). Dry-run default; --apply to execute.",
    )
    sp.add_argument("message_id", help="Backend-scoped message id")
    sp.add_argument("--folder", help="Folder for AppleScript (default: Inbox)")
    sp.add_argument("--unread", action="store_true", help="Mark as unread instead of read")
    sp.add_argument("--apply", action="store_true", help="Actually perform the write (otherwise dry-run)")
    sp.add_argument("--json", action="store_true", help="Emit the OperationResult as JSON")
    sp.set_defaults(func=cmd_mark_read)

    # ---- send ----
    sp = subparsers.add_parser("send", help="Send an email")
    sp.add_argument("--to", nargs="+", required=True, help="Recipient(s)")
    sp.add_argument("--subject", required=True, help="Email subject")
    sp.add_argument("--body", required=True, help="Email body")
    sp.add_argument("--cc", nargs="+", help="CC recipient(s)")
    sp.add_argument("--bcc", nargs="+", help="BCC recipient(s)")
    sp.add_argument("--attach", nargs="+", help="File(s) to attach")
    sp.add_argument("--html", action="store_true", help="Treat body as HTML")
    sp.add_argument("--importance", choices=["high", "normal", "low"])
    sp.set_defaults(func=cmd_send)

    args = parser.parse_args()
    result = args.func(args)
    if isinstance(result, int):
        sys.exit(result)


if __name__ == "__main__":
    main()
