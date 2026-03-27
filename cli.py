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

from outlook_tool import OutlookClient


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

    # JSON output if requested
    if args.json:
        serializable = []
        for e in results:
            entry = {k: v for k, v in e.items() if not k.startswith("_")}
            entry["received_datetime"] = entry["received_datetime"].isoformat()
            serializable.append(entry)
        print(json.dumps(serializable, indent=2))

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
        print()

    if args.json:
        serializable = []
        for e in results:
            entry = {k: v for k, v in e.items() if not k.startswith("_")}
            entry["start_datetime"] = entry["start_datetime"].isoformat()
            entry["end_datetime"] = entry["end_datetime"].isoformat()
            serializable.append(entry)
        print(json.dumps(serializable, indent=2))

    return results


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
    sp.add_argument("--json", action="store_true", help="Output results as JSON")
    sp.set_defaults(func=cmd_events)

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
    args.func(args)


if __name__ == "__main__":
    main()
