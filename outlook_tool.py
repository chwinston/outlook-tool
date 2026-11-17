#!/usr/bin/env python3
"""
Outlook Tool — General-purpose Outlook email client for Python.

Cross-platform wrapper around Microsoft Outlook that supports:
  - Searching/filtering emails by date, subject, sender, domain, folder, etc.
  - Downloading attachments from matched emails
  - Sending emails with optional attachments

Platform detection (priority order):
  - Windows: win32com (Outlook COM automation) — talks directly to desktop Outlook
  - Mac: AppleScript/JXA — talks directly to Outlook for Mac (Legacy/Classic mode)
  - Fallback: Microsoft Graph API via MSAL device code auth

Usage:
    from outlook_tool import OutlookClient

    client = OutlookClient()

    # Search emails
    results = client.search(
        date_from="2026-03-01",
        date_to="2026-03-15",
        sender_domain="example.com",
        subject_contains="quarterly report",
        has_attachments=True,
    )

    # Download attachments
    for email in results:
        for att in email["attachments"]:
            client.download_attachment(email, att, output_dir="./downloads")

    # Send an email
    client.send(
        to=["colleague@example.com"],
        subject="Meeting notes",
        body="Please find the notes attached.",
        attachments=["notes.pdf"],
    )
"""

import hashlib
import os
import re
import sys
import time as _time_module
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from urllib.parse import quote as _url_quote


# =============================================================================
# OPERATION RESULT (Bolt 4 — write verbs contract)
# =============================================================================

@dataclass
class OperationResult:
    """Unified return contract for write operations across every backend.

    Produced by :meth:`OutlookClient.move`, :meth:`OutlookClient.delete`,
    :meth:`OutlookClient.mark_read` and the backend implementations they
    dispatch to.

    Fields
    ------
    status:
        One of ``"success"``, ``"failed"``, ``"unsupported"``, ``"dry_run"``.
        ``success`` means the op was attempted and returned without an error;
        callers MUST also check ``verified`` before trusting the state change.
        ``unsupported`` is reserved for cases the backend cannot execute —
        e.g. Graph without ``Mail.ReadWrite`` scope.
    verified:
        ``True`` only when a post-op re-fetch confirms the target state.
        This defends against the AppleScript silent-success failure mode where
        the op returns no error but Outlook refused the change (Tim F1).
    message_id:
        Backend-scoped identifier of the message that was operated on.
    operation:
        The verb executed: ``"move"``, ``"delete"`` or ``"mark_read"``.
    before_state:
        Snapshot of the salient fields before the op (e.g. ``{"folder": "Inbox",
        "is_read": False}``). Used to verify no-op safety and for audit logs.
    after_state:
        Snapshot of the salient fields after the op, when the backend could
        re-fetch successfully.
    error_detail:
        Human-readable detail when ``status`` is ``failed`` or ``unsupported``.
    backend:
        Which backend handled the op — ``"applescript"``, ``"win32com"`` or
        ``"graph"`` — so callers can reason about provenance.
    elapsed_ms:
        Wall-clock duration of the call in milliseconds. Useful for the
        batching decision Phase 3 will make later.
    """

    status: str  # "success" | "failed" | "unsupported" | "dry_run"
    operation: str  # "move" | "delete" | "mark_read"
    message_id: Optional[str] = None
    verified: bool = False
    before_state: Dict[str, Any] = field(default_factory=dict)
    after_state: Dict[str, Any] = field(default_factory=dict)
    error_detail: Optional[str] = None
    backend: Optional[str] = None
    elapsed_ms: Optional[int] = None

    _VALID_STATUSES = frozenset({"success", "failed", "unsupported", "dry_run"})
    _VALID_OPERATIONS = frozenset({"move", "delete", "mark_read"})

    def __post_init__(self) -> None:
        if self.status not in self._VALID_STATUSES:
            raise ValueError(
                f"Invalid OperationResult.status={self.status!r}; "
                f"must be one of {sorted(self._VALID_STATUSES)}"
            )
        if self.operation not in self._VALID_OPERATIONS:
            raise ValueError(
                f"Invalid OperationResult.operation={self.operation!r}; "
                f"must be one of {sorted(self._VALID_OPERATIONS)}"
            )

    def to_dict(self) -> Dict[str, Any]:
        """Serialize to a plain dict (JSON-safe for CLI output)."""
        return asdict(self)

# =============================================================================
# PLATFORM DETECTION
# =============================================================================

HAS_WIN32 = False
HAS_APPLESCRIPT = False
HAS_GRAPH = False

if sys.platform == "win32":
    try:
        import win32com.client
        HAS_WIN32 = True
    except ImportError:
        pass

if not HAS_WIN32 and sys.platform == "darwin":
    # Check if osascript exists (always does on macOS)
    import shutil
    HAS_APPLESCRIPT = shutil.which("osascript") is not None

try:
    import requests as _requests
except ImportError:
    _requests = None  # type: ignore[assignment]

if not HAS_WIN32 and not HAS_APPLESCRIPT:
    try:
        import msal
        if _requests is not None:
            HAS_GRAPH = True
    except ImportError:
        HAS_GRAPH = False


# =============================================================================
# APPLESCRIPT/JXA BACKEND (Mac — Legacy/Classic Outlook)
# =============================================================================

import json
import subprocess


def _escape_applescript(s: str) -> str:
    """Escape a string for safe interpolation into AppleScript double-quoted strings.

    Handles backslashes, double quotes, newlines, carriage returns, and tabs
    to prevent command injection via crafted email addresses, subjects, or paths.
    """
    s = s.replace("\0", "")  # Strip null bytes
    s = s.replace("\\", "\\\\")
    s = s.replace('"', '\\"')
    s = s.replace("\n", "\\n")
    s = s.replace("\r", "\\r")
    s = s.replace("\t", "\\t")
    return s


def _run_jxa(script: str, timeout: int = 120) -> str:
    """Run a JXA script via osascript and return stdout."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(f"JXA error: {result.stderr.strip()}")
    return result.stdout.strip()


def _run_applescript(script: str, timeout: int = 60) -> str:
    """Run an AppleScript via osascript and return stdout."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
    return result.stdout.strip()


class _AppleScriptBackend:
    """Outlook for Mac email backend using JXA/AppleScript.

    Requires macOS with Microsoft Outlook running in Legacy/Classic mode.
    No API keys or Azure registration needed — talks directly to the local app.

    Bolt 5 — verify hardening
    -------------------------
    The AppleScript ``move`` command returns before the Exchange→client sync
    settles, so a verify scan that fires immediately sees the pre-move cached
    state and reports "still in source" for a message that actually did move.
    See ``docs/pm/plans/20261117-0740-bolt5-batching-plan.md`` for the pick-off
    experiment that confirmed this.

    Two mitigations, both configurable via constructor args (both honour env
    overrides so the apply-triage CLI can tune without API change):

    * ``verify_delay_ms`` (``OUTLOOK_TOOL_VERIFY_DELAY_MS``, default 1500ms) —
      sleep between the AppleScript write and the fingerprint verify scan.
    * ``verify_use_count_check`` (``OUTLOOK_TOOL_VERIFY_USE_COUNT``, default
      True) — gate verify on a source-folder count delta. If the source count
      dropped by one, that is strong evidence of success even when the
      fingerprint scan can't find the row (large Deleted Items / Archive).
    """

    _DEFAULT_VERIFY_DELAY_MS = 1500

    def __init__(
        self,
        verify_delay_ms: Optional[int] = None,
        verify_use_count_check: Optional[bool] = None,
    ):
        if sys.platform != "darwin":
            raise RuntimeError("AppleScript backend only works on macOS")

        try:
            version = _run_applescript(
                'tell application "Microsoft Outlook" to get version'
            )
            self._version = version
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to Outlook for Mac. "
                f"Make sure Outlook is running in Legacy/Classic mode.\n{e}"
            )

        if verify_delay_ms is None:
            env = os.environ.get("OUTLOOK_TOOL_VERIFY_DELAY_MS")
            if env is not None:
                try:
                    verify_delay_ms = max(0, int(env))
                except ValueError:
                    verify_delay_ms = self._DEFAULT_VERIFY_DELAY_MS
            else:
                verify_delay_ms = self._DEFAULT_VERIFY_DELAY_MS
        self.verify_delay_ms = max(0, int(verify_delay_ms))

        if verify_use_count_check is None:
            env = os.environ.get("OUTLOOK_TOOL_VERIFY_USE_COUNT")
            if env is not None:
                verify_use_count_check = env.strip().lower() not in {
                    "0", "false", "no", "off",
                }
            else:
                verify_use_count_check = True
        self.verify_use_count_check = bool(verify_use_count_check)

    def scan_emails(
        self,
        start_date: datetime,
        end_date: datetime,
        sender_domains: Optional[List[str]] = None,
        only_with_attachments: Optional[bool] = None,
        max_results: int = 250,
        folder_name: Optional[str] = None,
    ) -> List[Dict]:
        """Scan a mail folder for emails within date range via JXA.

        Args:
            folder_name: Folder to scan (default: Inbox). Must match the exact
                folder name as it appears in Outlook (e.g., "Archive", "Snoozed",
                "Sent Items", "Deleted Items", "Junk Email", "Drafts").
        """
        domain_filter_js = "null"
        if sender_domains:
            domain_filter_js = json.dumps([d.lower() for d in sender_domains])

        att_filter = "true" if only_with_attachments else "false"
        safe_folder = json.dumps(folder_name) if folder_name else "null"

        jxa_script = f"""
        (function() {{
            var outlook = Application("Microsoft Outlook");
            var folderName = {safe_folder};
            var folder;
            if (!folderName || folderName.toLowerCase() === "inbox") {{
                folder = outlook.inbox;
            }} else {{
                var accts = outlook.exchangeAccounts();
                if (accts.length === 0) {{
                    return JSON.stringify({{error: "No Exchange account found"}});
                }}
                var acct = accts[0];
                var allFolders = acct.mailFolders();
                var found = false;
                for (var fi = 0; fi < allFolders.length; fi++) {{
                    try {{
                        if (allFolders[fi].name() === folderName) {{
                            folder = allFolders[fi];
                            found = true;
                            break;
                        }}
                    }} catch(e) {{}}
                }}
                if (!found) {{
                    return JSON.stringify({{error: "Folder not found: " + folderName}});
                }}
            }}
            var msgCount = folder.messages.length;

            var startMs = {int(start_date.timestamp() * 1000)};
            var endMs = {int((end_date + timedelta(days=1)).timestamp() * 1000)};
            var domainFilter = {domain_filter_js};
            var onlyWithAtt = {att_filter};
            var maxResults = {max_results};

            var results = [];
            var scanned = 0;

            for (var i = 0; i < msgCount && i < 10000; i++) {{
                if (results.length >= maxResults) break;
                try {{
                    var msg = folder.messages[i];
                    var recvDate = msg.timeReceived();
                    if (!recvDate) continue;

                    var recvMs = recvDate.getTime();

                    if (recvMs < startMs) continue;
                    if (recvMs >= endMs) continue;

                    scanned++;

                    var attCount = msg.attachments.length;
                    if (onlyWithAtt && attCount === 0) continue;

                    var senderName = "";
                    var senderAddr = "";
                    try {{
                        var s = msg.sender();
                        senderName = s.name || "";
                        senderAddr = s.address || "";
                    }} catch(e) {{}}

                    if (domainFilter && senderAddr) {{
                        var domain = senderAddr.split("@").pop().toLowerCase();
                        if (domainFilter.indexOf(domain) === -1) continue;
                    }}

                    var atts = [];
                    for (var j = 0; j < attCount; j++) {{
                        try {{
                            var att = msg.attachments[j];
                            atts.push({{
                                name: att.name(),
                                size: att.fileSize(),
                                index: j + 1
                            }});
                        }} catch(e) {{}}
                    }}

                    if (onlyWithAtt && atts.length === 0) continue;

                    var subject = "";
                    try {{ subject = msg.subject() || "(No Subject)"; }} catch(e) {{ subject = "(No Subject)"; }}

                    var isRead = false;
                    try {{ isRead = msg.isRead(); }} catch(e) {{}}

                    var toStr = "";
                    try {{
                        var recipients = msg.toRecipients();
                        var addrs = [];
                        for (var k = 0; k < recipients.length; k++) {{
                            try {{ addrs.push(recipients[k].address()); }} catch(e) {{}}
                        }}
                        toStr = addrs.join(", ");
                    }} catch(e) {{}}

                    var bodyPreview = "";
                    try {{ bodyPreview = (msg.plainTextContent() || "").substring(0, 5000); }} catch(e) {{}}

                    results.push({{
                        _msg_index: i + 1,
                        _folder_name: folderName || "Inbox",
                        subject: subject,
                        sender_name: senderName,
                        sender_email: senderAddr,
                        received_datetime: recvDate.toISOString(),
                        is_read: isRead,
                        has_attachments: atts.length > 0,
                        to: toStr,
                        body_preview: bodyPreview,
                        attachments: atts
                    }});

                }} catch(e) {{
                    continue;
                }}
            }}

            return JSON.stringify({{
                scanned: scanned,
                matched: results.length,
                results: results
            }});
        }})();
        """

        display_folder = folder_name or "Inbox"
        print(f"Scanning Outlook {display_folder}: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}...", file=sys.stderr)

        raw = _run_jxa(jxa_script, timeout=180)
        data = json.loads(raw)

        if "error" in data:
            raise ValueError(data["error"])

        print(f"  Scanned {data['scanned']} messages, found {data['matched']} matches", file=sys.stderr)

        results = []
        for email in data["results"]:
            try:
                received_dt = datetime.fromisoformat(
                    email["received_datetime"].replace("Z", "+00:00")
                ).replace(tzinfo=None)
            except (ValueError, AttributeError):
                received_dt = datetime.now()

            sender_email_addr = email.get("sender_email", "")

            email_folder = email.get("_folder_name", "Inbox")

            att_list = []
            for att in email.get("attachments", []):
                att_list.append({
                    "id": f"as_{email['_msg_index']}_{att['index']}",
                    "name": att["name"],
                    "size": att.get("size", 0),
                    "_as_msg_index": email["_msg_index"],
                    "_as_att_index": att["index"],
                    "_as_folder_name": email_folder,
                })

            results.append({
                "id": f"as_msg_{email['_msg_index']}",
                "subject": email.get("subject", "(No Subject)"),
                "sender_name": email.get("sender_name", ""),
                "sender_email": sender_email_addr,
                "received_datetime": received_dt,
                "received_date": received_dt.strftime("%Y-%m-%d"),
                "day_of_week": received_dt.strftime("%A"),
                "is_read": email.get("is_read", False),
                "has_attachments": email.get("has_attachments", False),
                "importance": "normal",  # AppleScript doesn't expose importance easily
                "body_preview": email.get("body_preview", ""),
                "to": email.get("to", ""),
                "attachments": att_list,
            })

        return results

    def save_attachment(
        self, msg_index: int, att_index: int, output_path: Path,
        folder_name: Optional[str] = None,
    ) -> bool:
        """Save an attachment to disk via AppleScript."""
        msg_index = int(msg_index)
        att_index = int(att_index)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        save_path = _escape_applescript(str(output_path.resolve()))

        if not folder_name or folder_name.lower() == "inbox":
            folder_ref = "inbox"
        else:
            safe_name = _escape_applescript(folder_name)
            folder_ref = f'mail folder "{safe_name}" of (first exchange account)'

        script = f'''
tell application "Microsoft Outlook"
    set m to message {msg_index} of {folder_ref}
    set att to attachment {att_index} of m
    save att in POSIX file "{save_path}"
end tell
'''
        _run_applescript(script)
        return True

    def send_email(
        self,
        to: List[str],
        subject: str,
        body: str,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        attachments: Optional[List[Path]] = None,
        html: bool = False,
    ) -> bool:
        """Send an email via AppleScript."""
        to_lines = "\n".join(
            f'        make new to recipient at end of to recipients with properties {{email address:{{address:"{_escape_applescript(addr)}"}}}}'
            for addr in to
        )
        cc_lines = ""
        if cc:
            cc_lines = "\n".join(
                f'        make new cc recipient at end of cc recipients with properties {{email address:{{address:"{_escape_applescript(addr)}"}}}}'
                for addr in cc
            )
        bcc_lines = ""
        if bcc:
            bcc_lines = "\n".join(
                f'        make new bcc recipient at end of bcc recipients with properties {{email address:{{address:"{_escape_applescript(addr)}"}}}}'
                for addr in bcc
            )

        safe_subject = _escape_applescript(subject)
        safe_body = _escape_applescript(body)

        content_prop = "content" if html else "plain text content"

        att_lines = ""
        if attachments:
            att_lines = "\n".join(
                f'        make new attachment at end of attachments with properties {{file:POSIX file "{_escape_applescript(str(p.resolve()))}"}}'
                for p in attachments
            )

        script = f'''
tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{safe_subject}", {content_prop}:"{safe_body}"}}
    tell newMsg
{to_lines}
{cc_lines}
{bcc_lines}
{att_lines}
    end tell
    send newMsg
end tell
'''
        _run_applescript(script, timeout=30)
        return True

    def get_calendar_events(
        self,
        start_date: datetime,
        end_date: datetime,
        max_results: int = 50,
    ) -> List[Dict]:
        """Fetch calendar events within a date range via JXA."""
        jxa_script = f"""
        (function() {{
            var outlook = Application("Microsoft Outlook");
            var startMs = {int(start_date.timestamp() * 1000)};
            var endMs = {int((end_date + timedelta(days=1)).timestamp() * 1000)};
            var maxResults = {max_results};
            var results = [];

            var accts = outlook.exchangeAccounts();
            if (accts.length === 0) {{
                return JSON.stringify({{error: "No Exchange account found"}});
            }}
            var acct = accts[0];
            var calendar = acct.defaultCalendar();
            var events = calendar.calendarEvents();
            var evtCount = events.length;

            for (var i = 0; i < evtCount && i < 10000; i++) {{
                if (results.length >= maxResults) break;
                try {{
                    var evt = events[i];
                    var startTime = evt.startTime();
                    var endTime = evt.endTime();
                    if (!startTime) continue;

                    var startMs2 = startTime.getTime();
                    var endMs2 = endTime ? endTime.getTime() : startMs2;

                    // Event overlaps range if it starts before range end AND ends after range start
                    if (startMs2 >= endMs) continue;
                    if (endMs2 < startMs) continue;

                    var subject = "";
                    try {{ subject = evt.subject() || "(No Subject)"; }} catch(e) {{ subject = "(No Subject)"; }}

                    var location = "";
                    try {{ location = evt.location() || ""; }} catch(e) {{}}

                    var organizer = "";
                    try {{ organizer = evt.organizer() || ""; }} catch(e) {{}}

                    var isAllDay = false;
                    try {{ isAllDay = evt.allDayFlag(); }} catch(e) {{}}

                    var status = "busy";
                    try {{
                        var fbs = evt.freeBusyStatus();
                        if (fbs === "free") status = "free";
                        else if (fbs === "tentative") status = "tentative";
                        else if (fbs === "out of office" || fbs === "working elsewhere") status = fbs;
                        else status = "busy";
                    }} catch(e) {{}}

                    var bodyPreview = "";
                    try {{ bodyPreview = (evt.plainTextContent() || "").substring(0, 2000); }} catch(e) {{}}

                    var attendees = [];
                    try {{
                        var attList = evt.attendees();
                        for (var j = 0; j < attList.length; j++) {{
                            try {{
                                var att = attList[j];
                                attendees.push({{
                                    name: att.name() || "",
                                    email: att.emailAddress ? (att.emailAddress().address || "") : "",
                                    status: att.acceptedStatus ? att.acceptedStatus() : "unknown"
                                }});
                            }} catch(e2) {{}}
                        }}
                    }} catch(e) {{}}

                    results.push({{
                        subject: subject,
                        start_datetime: startTime.toISOString(),
                        end_datetime: endTime ? endTime.toISOString() : startTime.toISOString(),
                        location: location,
                        organizer: organizer,
                        is_all_day: isAllDay,
                        status: status,
                        body_preview: bodyPreview,
                        attendees: attendees
                    }});
                }} catch(e) {{ continue; }}
            }}

            return JSON.stringify({{
                matched: results.length,
                results: results
            }});
        }})();
        """

        print(f"Scanning Outlook calendar: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}...", file=sys.stderr)

        raw = _run_jxa(jxa_script, timeout=180)
        data = json.loads(raw)

        if "error" in data:
            raise ValueError(data["error"])

        print(f"  Found {data['matched']} event(s)", file=sys.stderr)

        results = []
        for evt in data["results"]:
            try:
                start_dt = datetime.fromisoformat(
                    evt["start_datetime"].replace("Z", "+00:00")
                ).replace(tzinfo=None)
            except (ValueError, AttributeError):
                start_dt = datetime.now()

            try:
                end_dt = datetime.fromisoformat(
                    evt["end_datetime"].replace("Z", "+00:00")
                ).replace(tzinfo=None)
            except (ValueError, AttributeError):
                end_dt = start_dt

            results.append({
                "id": f"as_evt_{hashlib.md5((evt['subject'] + evt['start_datetime']).encode()).hexdigest()[:12]}",
                "subject": evt.get("subject", "(No Subject)"),
                "start_datetime": start_dt,
                "end_datetime": end_dt,
                "start_date": start_dt.strftime("%Y-%m-%d"),
                "end_date": end_dt.strftime("%Y-%m-%d"),
                "location": evt.get("location", ""),
                "organizer_name": evt.get("organizer", ""),
                "organizer_email": "",
                "is_all_day": evt.get("is_all_day", False),
                "status": evt.get("status", "busy"),
                "body_preview": evt.get("body_preview", ""),
                "attendees": evt.get("attendees", []),
            })

        return results

    # =========================================================================
    # WRITE OPERATIONS (Bolt 4) — mark_read / move / delete
    # =========================================================================
    #
    # Round-trip verify is mandatory on every write. The dominant AppleScript
    # risk is silent-success (Tim F1): the op returns 0 bytes to stderr but
    # Outlook didn't actually perform the change. Every public write method
    # below re-queries the message after the op and sets
    # ``OperationResult.verified`` only when the observed post-state matches
    # the intent.
    #
    # NOTE FOR PHASE 2: AppleScript message indices are session-scoped. A
    # triage report produced via ``scan_emails`` and acted on more than a
    # few minutes later may encounter stale indices (Outlook re-sorted the
    # folder, user deleted something, etc.). The executor in Phase 2 MUST
    # refuse reports > 1h old and SHOULD fall back to (sender, subject,
    # received_date) identity if the raw index misses.

    @staticmethod
    def _folder_ref(folder_name: Optional[str]) -> str:
        """Return the AppleScript expression resolving to the right mail folder.

        Inbox is a special-cased convenience; anything else goes through the
        first exchange account.
        """
        if not folder_name or folder_name.lower() == "inbox":
            return "inbox"
        safe = _escape_applescript(folder_name)
        return f'mail folder "{safe}" of (first exchange account)'

    def _fetch_message_state(
        self, msg_index: int, folder_name: Optional[str]
    ) -> Dict[str, Any]:
        """Re-fetch a message's salient state for round-trip verification.

        Returns a dict with ``folder``, ``is_read``, ``subject``, ``sender_email``.
        Raises ``RuntimeError`` if the message cannot be found in ``folder_name``
        (used by move/delete verify to confirm the message actually left).
        """
        msg_index = int(msg_index)
        safe_folder = json.dumps(folder_name) if folder_name else "null"
        script = f"""
        (function() {{
            var outlook = Application("Microsoft Outlook");
            var folderName = {safe_folder};
            var folder;
            if (!folderName || folderName.toLowerCase() === "inbox") {{
                folder = outlook.inbox;
            }} else {{
                var accts = outlook.exchangeAccounts();
                if (accts.length === 0) return JSON.stringify({{error: "no_account"}});
                var allFolders = accts[0].mailFolders();
                for (var fi = 0; fi < allFolders.length; fi++) {{
                    try {{
                        if (allFolders[fi].name() === folderName) {{
                            folder = allFolders[fi];
                            break;
                        }}
                    }} catch(e) {{}}
                }}
                if (!folder) return JSON.stringify({{error: "folder_not_found"}});
            }}
            try {{
                var m = folder.messages[{msg_index - 1}];
                var subj = ""; try {{ subj = m.subject() || ""; }} catch(e) {{}}
                var isRead = false; try {{ isRead = m.isRead(); }} catch(e) {{}}
                var addr = ""; try {{ addr = m.sender().address || ""; }} catch(e) {{}}
                return JSON.stringify({{subject: subj, is_read: isRead, sender_email: addr}});
            }} catch(e) {{
                return JSON.stringify({{error: "message_not_found"}});
            }}
        }})();
        """
        raw = _run_jxa(script, timeout=60)
        try:
            data = json.loads(raw)
        except Exception as e:
            raise RuntimeError(f"_fetch_message_state: bad JSON: {e}")
        if isinstance(data, dict) and "error" in data:
            raise RuntimeError(
                f"Message index {msg_index} not resolvable in folder "
                f"{folder_name or 'Inbox'!r}: {data['error']}"
            )
        return {
            "folder": folder_name or "Inbox",
            "is_read": bool(data.get("is_read", False)),
            "subject": data.get("subject", ""),
            "sender_email": data.get("sender_email", ""),
        }

    def mark_read(
        self,
        msg_index: int,
        folder_name: Optional[str] = None,
        read: bool = True,
    ) -> OperationResult:
        """Set the read/unread flag on a message and verify.

        Args:
            msg_index: AppleScript message index inside ``folder_name``.
                Comes from a ``scan_emails`` result's ``_as_msg_index``.
            folder_name: Folder the message lives in. ``None``/``"Inbox"``
                targets the default Inbox.
            read: ``True`` to mark read, ``False`` to mark unread.

        Returns:
            ``OperationResult`` with ``verified=True`` iff the post-op state
            matches the requested ``read`` value.
        """
        start_ms = _time_module.perf_counter()
        msg_index = int(msg_index)
        folder_ref = self._folder_ref(folder_name)
        operation = "mark_read"
        message_id = f"as_msg_{msg_index}"

        try:
            before = self._fetch_message_state(msg_index, folder_name)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                error_detail=f"pre-fetch failed: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        target_bool = "true" if read else "false"
        script = f'''
tell application "Microsoft Outlook"
    set m to message {msg_index} of {folder_ref}
    set read status of m to {target_bool}
end tell
'''
        try:
            _run_applescript(script)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                before_state=before,
                error_detail=f"applescript error: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Bolt 5 (B5.3) — settle delay before verify so Exchange→client
        # sync has caught up before we re-read the message state.
        self._sleep_verify_delay()

        # Round-trip verify.
        try:
            after = self._fetch_message_state(msg_index, folder_name)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                before_state=before,
                error_detail=f"post-verify fetch failed: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        verified = after["is_read"] == bool(read)
        return OperationResult(
            status="success" if verified else "failed",
            operation=operation,
            message_id=message_id,
            verified=verified,
            before_state=before,
            after_state=after,
            error_detail=None if verified else "post-state did not match request (silent success?)",
            backend="applescript",
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    def _folder_exists(self, folder_name: str) -> bool:
        """Does a top-level mail folder with this exact name exist?"""
        safe = _escape_applescript(folder_name)
        script = f'''
tell application "Microsoft Outlook"
    try
        set f to mail folder "{safe}" of (first exchange account)
        return "yes"
    on error
        return "no"
    end try
end tell
'''
        return _run_applescript(script).strip() == "yes"

    def ensure_folder(self, folder_name: str) -> None:
        """Detect-and-create a top-level mail folder on the primary account.

        No-op if the folder already exists. Raises ``RuntimeError`` on failure.
        """
        if not folder_name or folder_name.lower() == "inbox":
            return  # inbox is always present
        if self._folder_exists(folder_name):
            return
        safe = _escape_applescript(folder_name)
        script = f'''
tell application "Microsoft Outlook"
    set acct to first exchange account
    make new mail folder at acct with properties {{name:"{safe}"}}
end tell
'''
        _run_applescript(script)
        if not self._folder_exists(folder_name):
            raise RuntimeError(
                f"ensure_folder: creation of {folder_name!r} did not take effect"
            )

    def move(
        self,
        msg_index: int,
        target_folder: str,
        source_folder: Optional[str] = None,
    ) -> OperationResult:
        """Move a message from ``source_folder`` to ``target_folder``.

        Creates ``target_folder`` on the primary Exchange account if it doesn't
        yet exist (detect-and-create — Al F3). Round-trip verify re-queries the
        TARGET folder to confirm the message actually arrived (Tim F1 —
        silent-success defence).

        Note on verification: we re-query the target folder by scanning the last
        50 messages for a matching ``(subject, sender_email)`` fingerprint. This
        is more robust than trusting the post-move index would stay valid, since
        AppleScript indices reorder after a move.

        Args:
            msg_index: AppleScript index in the source folder.
            target_folder: Destination folder name. Created if missing.
            source_folder: Where the message lives now. ``None`` = Inbox.

        Returns:
            ``OperationResult`` with ``verified=True`` only when a fingerprint
            match is found in ``target_folder`` within the last 50 messages.
        """
        start_ms = _time_module.perf_counter()
        msg_index = int(msg_index)
        operation = "move"
        message_id = f"as_msg_{msg_index}"
        source_folder = source_folder or "Inbox"
        src_ref = self._folder_ref(source_folder)

        if not target_folder or target_folder.strip() == "":
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                error_detail="target_folder must be a non-empty folder name",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        try:
            before = self._fetch_message_state(msg_index, source_folder)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                error_detail=f"pre-fetch failed: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Detect-and-create target folder (Al F3).
        try:
            self.ensure_folder(target_folder)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                before_state=before,
                error_detail=f"ensure_folder({target_folder!r}) failed: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Bolt 5 (B5.3) — capture source folder count BEFORE the move so we can
        # use count-delta as a verify signal (robust against sync-lag).
        pre_count = (
            self._folder_message_count(source_folder)
            if self.verify_use_count_check
            else -1
        )

        safe_target = _escape_applescript(target_folder)
        script = f'''
tell application "Microsoft Outlook"
    set m to message {msg_index} of {src_ref}
    move m to mail folder "{safe_target}" of (first exchange account)
end tell
'''
        try:
            _run_applescript(script)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                before_state=before,
                error_detail=f"applescript error: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Bolt 5 (B5.3) — settle delay before verify. The move returns before
        # Exchange→client sync has caught up; without this delay the verify
        # scan sees cached state and reports false-failure in 29%+ of ops.
        self._sleep_verify_delay()

        # Round-trip verify (Tim F1) uses THREE signals now (Bolt 5):
        # (1) count-delta on source: fast + robust against fingerprint miss
        # (2) source-absence via fingerprint scan (legacy signal)
        # (3) target-presence via fingerprint scan (weakest — target may be
        #     huge and the match can fall off the scan limit)
        target_subject = before.get("subject", "")
        target_sender = before.get("sender_email", "")

        post_count = (
            self._folder_message_count(source_folder)
            if self.verify_use_count_check
            else -1
        )
        count_check_available = pre_count >= 0 and post_count >= 0
        count_dropped = count_check_available and (pre_count - post_count) >= 1

        still_in_source = self._verify_present_in_folder(
            folder_name=source_folder,
            subject=target_subject,
            sender_email=target_sender,
            scan_limit=200,
        )
        present_in_target = self._verify_present_in_folder(
            folder_name=target_folder,
            subject=target_subject,
            sender_email=target_sender,
            scan_limit=500,
        )

        # The move succeeded if EITHER the count dropped OR the fingerprint
        # scan says the message is gone from source. Count-delta is the new
        # primary signal under Bolt 5 — it addresses the sync-lag false-failure
        # mode directly. Target-side presence is confirmation icing.
        move_succeeded = count_dropped or (not still_in_source)
        verified = move_succeeded and present_in_target

        after: Dict[str, Any] = {
            "folder": target_folder,
            "subject": target_subject,
            "sender_email": target_sender,
        }

        if move_succeeded:
            # The write worked. verified distinguishes strong-confirmed vs
            # source-absence-only (Phase 1's UNVERIFIED-success tier).
            return OperationResult(
                status="success",
                operation=operation,
                message_id=message_id,
                verified=verified,
                before_state=before,
                after_state=after,
                error_detail=None if verified else (
                    f"source-absence confirms move; target-presence unconfirmed "
                    f"(scanned first 500 of {target_folder!r})"
                ),
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Source still has the message → the AppleScript move actually
        # didn't take effect. This is the real silent-success case.
        return OperationResult(
            status="failed",
            operation=operation,
            message_id=message_id,
            verified=False,
            before_state=before,
            after_state=after,
            error_detail=(
                f"silent-success: AppleScript returned OK but message is "
                f"still in {source_folder!r}"
            ),
            backend="applescript",
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    def _sleep_verify_delay(self) -> None:
        """Sleep for ``verify_delay_ms`` to let Exchange→client sync settle.

        Called between an AppleScript write and the verify scan that follows.
        Diagnoses in ``docs/pm/plans/20261117-0740-bolt5-batching-plan.md``
        pin false-failure reports on this sync lag.
        """
        if self.verify_delay_ms <= 0:
            return
        _time_module.sleep(self.verify_delay_ms / 1000.0)

    def _folder_message_count(self, folder_name: Optional[str]) -> int:
        """Return ``messages.length`` for ``folder_name`` on the primary account.

        Used as a count-delta verify signal — if the source folder dropped by
        the expected number of messages after a (bulk) move, treat it as
        strong evidence of success even when the fingerprint scan misses.

        Returns ``-1`` if the folder cannot be resolved; callers treat that
        as "count-check unavailable" (fall back to fingerprint only).
        """
        safe_folder = json.dumps(folder_name or "Inbox")
        script = f"""
        (function() {{
            var outlook = Application("Microsoft Outlook");
            var folderName = {safe_folder};
            var folder;
            if (!folderName || folderName.toLowerCase() === "inbox") {{
                folder = outlook.inbox;
            }} else {{
                var accts = outlook.exchangeAccounts();
                if (accts.length === 0) return "-1";
                var allFolders = accts[0].mailFolders();
                for (var fi = 0; fi < allFolders.length; fi++) {{
                    try {{
                        if (allFolders[fi].name() === folderName) {{
                            folder = allFolders[fi];
                            break;
                        }}
                    }} catch(e) {{}}
                }}
                if (!folder) return "-1";
            }}
            try {{
                return String(folder.messages.length);
            }} catch(e) {{
                return "-1";
            }}
        }})();
        """
        try:
            raw = _run_jxa(script, timeout=30)
            return int(raw.strip() or "-1")
        except Exception:
            return -1

    def _verify_present_in_folder(
        self,
        folder_name: str,
        subject: str,
        sender_email: str,
        scan_limit: int = 50,
    ) -> bool:
        """Scan the top ``scan_limit`` messages of ``folder_name`` for a match.

        Returns True if any message's ``(subject, sender_email)`` matches.
        ``sender_email`` comparison is case-insensitive and tolerates the
        empty string (some drafts and in-house messages lack a resolvable
        SMTP sender on AppleScript).
        """
        folder_ref = self._folder_ref(folder_name)
        # Build a JXA-side fetch so we can compare in Python and avoid the
        # AppleScript/tell-block handler scoping quirks.
        safe_folder = json.dumps(folder_name)
        script = f"""
        (function() {{
            var outlook = Application("Microsoft Outlook");
            var folderName = {safe_folder};
            var folder;
            if (!folderName || folderName.toLowerCase() === "inbox") {{
                folder = outlook.inbox;
            }} else {{
                var accts = outlook.exchangeAccounts();
                if (accts.length === 0) return JSON.stringify([]);
                var allFolders = accts[0].mailFolders();
                for (var fi = 0; fi < allFolders.length; fi++) {{
                    try {{
                        if (allFolders[fi].name() === folderName) {{
                            folder = allFolders[fi];
                            break;
                        }}
                    }} catch(e) {{}}
                }}
                if (!folder) return JSON.stringify([]);
            }}
            var msgCount = folder.messages.length;
            var limit = Math.min(msgCount, {int(scan_limit)});
            var out = [];
            for (var i = 0; i < limit; i++) {{
                try {{
                    var m = folder.messages[i];
                    var subj = "";
                    try {{ subj = m.subject() || ""; }} catch(e) {{}}
                    var addr = "";
                    try {{ addr = m.sender().address || ""; }} catch(e) {{}}
                    out.push({{subject: subj, sender: addr}});
                }} catch(e) {{}}
            }}
            return JSON.stringify(out);
        }})();
        """
        raw = _run_jxa(script, timeout=60)
        try:
            entries = json.loads(raw or "[]")
        except Exception:
            return False
        target_subject = subject or ""
        target_sender = (sender_email or "").lower()
        for entry in entries:
            if entry.get("subject", "") != target_subject:
                continue
            if not target_sender:
                return True
            if (entry.get("sender", "") or "").lower() == target_sender:
                return True
        return False

    def delete_message(
        self,
        msg_index: int,
        folder_name: Optional[str] = None,
        permanent: bool = False,
    ) -> OperationResult:
        """Delete a message — soft-delete (move to Deleted Items) by default.

        Args:
            msg_index: AppleScript index in ``folder_name``.
            folder_name: Source folder. ``None`` = Inbox.
            permanent: If True, hard-delete (remove without trash). The
                public ``OutlookClient.delete`` enforces the extra
                ``yes_really_permanent`` gate on top of this flag.

        Returns:
            ``OperationResult``. For soft-delete, verified=True when the
            message appears in Deleted Items. For hard-delete, verified=True
            when the message is no longer present in the source folder.
        """
        start_ms = _time_module.perf_counter()
        msg_index = int(msg_index)
        operation = "delete"
        message_id = f"as_msg_{msg_index}"
        source_folder = folder_name or "Inbox"

        if not permanent:
            # Soft-delete is just a move into Deleted Items. Use the move path
            # so verification semantics (fingerprint hit in destination) match.
            result = self.move(
                msg_index=msg_index,
                target_folder="Deleted Items",
                source_folder=source_folder,
            )
            # Rewrite op to reflect intent — move internally, delete externally.
            return OperationResult(
                status=result.status,
                operation=operation,
                message_id=result.message_id,
                verified=result.verified,
                before_state=result.before_state,
                after_state={**result.after_state, "permanent": False},
                error_detail=result.error_detail,
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Permanent delete path.
        src_ref = self._folder_ref(source_folder)
        try:
            before = self._fetch_message_state(msg_index, source_folder)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                error_detail=f"pre-fetch failed: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Bolt 5 — capture pre-count for count-delta verify (B5.3).
        pre_count = (
            self._folder_message_count(source_folder)
            if self.verify_use_count_check
            else -1
        )

        script = f'''
tell application "Microsoft Outlook"
    set m to message {msg_index} of {src_ref}
    delete m
end tell
'''
        try:
            _run_applescript(script)
        except Exception as e:
            return OperationResult(
                status="failed",
                operation=operation,
                message_id=message_id,
                before_state=before,
                error_detail=f"applescript error: {e}",
                backend="applescript",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

        # Bolt 5 (B5.3) — settle delay before verify, then count-check + fingerprint.
        self._sleep_verify_delay()

        post_count = (
            self._folder_message_count(source_folder)
            if self.verify_use_count_check
            else -1
        )
        count_check_available = pre_count >= 0 and post_count >= 0
        count_dropped = count_check_available and (pre_count - post_count) >= 1

        still_there = self._verify_present_in_folder(
            folder_name=source_folder,
            subject=before.get("subject", ""),
            sender_email=before.get("sender_email", ""),
            scan_limit=500,
        )
        verified = count_dropped or (not still_there)
        return OperationResult(
            status="success" if verified else "failed",
            operation=operation,
            message_id=message_id,
            verified=verified,
            before_state=before,
            after_state={"folder": None, "permanent": True},
            error_detail=None if verified else (
                "post-delete verify: message still present in source folder"
            ),
            backend="applescript",
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    # ------------------------------------------------------------------
    # Test harness helpers (used by tests/test_write_ops.py only)
    # ------------------------------------------------------------------

    def delete_folder_if_empty(self, folder_name: str) -> bool:
        """Delete ``folder_name`` on the primary account iff it is empty.

        Returns True if the folder was deleted, False otherwise. Never raises
        for a missing folder.
        """
        if not folder_name or folder_name.lower() == "inbox":
            return False
        if not self._folder_exists(folder_name):
            return False
        safe = _escape_applescript(folder_name)
        script = f'''
tell application "Microsoft Outlook"
    set f to mail folder "{safe}" of (first exchange account)
    if (count of messages of f) is 0 then
        delete f
        return "deleted"
    else
        return "not_empty"
    end if
end tell
'''
        return _run_applescript(script).strip() == "deleted"

    def seed_draft_in_folder(
        self,
        folder_name: str,
        subject: str,
        body: str,
    ) -> Dict[str, Any]:
        """Create a draft message directly inside ``folder_name`` and return
        its state fingerprint + a newly-fetched AppleScript index.

        The draft is created, moved into the folder, and then re-queried so
        tests get the ``_as_msg_index`` they need to exercise write verbs.
        """
        self.ensure_folder(folder_name)
        safe_subject = _escape_applescript(subject)
        safe_body = _escape_applescript(body)
        safe_folder = _escape_applescript(folder_name)
        script = f'''
tell application "Microsoft Outlook"
    set draftMsg to make new outgoing message with properties {{subject:"{safe_subject}", plain text content:"{safe_body}"}}
    move draftMsg to mail folder "{safe_folder}" of (first exchange account)
end tell
'''
        _run_applescript(script)
        # Find the just-created message by subject.
        folder_ref = self._folder_ref(folder_name)
        find_script = f'''
tell application "Microsoft Outlook"
    set f to {folder_ref}
    set msgCount to count of messages of f
    set i to 1
    repeat while i <= msgCount and i <= 50
        try
            set m to message i of f
            if subject of m is equal to "{safe_subject}" then
                return i as text
            end if
        end try
        set i to i + 1
    end repeat
    return "not_found"
end tell
'''
        idx_raw = _run_applescript(find_script).strip()
        if idx_raw == "not_found":
            raise RuntimeError(
                f"seed_draft_in_folder: could not locate seeded message "
                f"{subject!r} in {folder_name!r}"
            )
        return {
            "msg_index": int(idx_raw),
            "folder_name": folder_name,
            "subject": subject,
        }


# =============================================================================
# GRAPH API BACKEND
# =============================================================================

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES_READ = ["Mail.Read"]
GRAPH_SCOPES_SEND = ["Mail.Read", "Mail.Send"]
GRAPH_SCOPES_READWRITE = ["Mail.ReadWrite"]
TOKEN_CACHE_PATH = Path.home() / ".outlook_tool_token_cache.bin"

# Microsoft Office desktop app — public client ID, no registration needed.
MS_OFFICE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"


class _GraphBackend:
    """Microsoft Graph API backend for Mac/Linux."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        token_cache_path: Optional[Path] = None,
        scopes: Optional[List[str]] = None,
    ):
        if not HAS_GRAPH:
            raise RuntimeError(
                "Graph API dependencies not installed. Run: pip install msal requests"
            )

        self.client_id = (
            client_id
            or os.environ.get("KPI_GRAPH_CLIENT_ID")
            or MS_OFFICE_CLIENT_ID
        )
        self.tenant_id = (
            tenant_id
            or os.environ.get("KPI_GRAPH_TENANT_ID")
            or "common"
        )
        self.token_cache_path = token_cache_path or TOKEN_CACHE_PATH
        self.scopes = scopes or GRAPH_SCOPES_READ

        self._cache = msal.SerializableTokenCache()
        self._load_cache()

        self._app = msal.PublicClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self._cache,
        )

        self._token = None

    def _load_cache(self):
        if self.token_cache_path.exists():
            self._cache.deserialize(self.token_cache_path.read_text())

    def _save_cache(self):
        if self._cache.has_state_changed:
            import tempfile
            cache_dir = str(self.token_cache_path.parent)
            fd = None
            tmp_path = None
            try:
                fd, tmp_path = tempfile.mkstemp(dir=cache_dir, prefix=".token_cache_")
                os.fchmod(fd, 0o600)
                os.write(fd, self._cache.serialize().encode())
                os.close(fd)
                fd = None
                os.rename(tmp_path, str(self.token_cache_path))
            except Exception:
                if fd is not None:
                    os.close(fd)
                if tmp_path and os.path.exists(tmp_path):
                    os.unlink(tmp_path)
                raise

    def _acquire_token(self) -> str:
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]

        flow = self._app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow initiation failed: {flow}")

        print("\n" + "=" * 60)
        print("  MICROSOFT GRAPH API AUTHENTICATION")
        print("=" * 60)
        print(f"\n  {flow['message']}")
        print("=" * 60 + "\n")

        result = self._app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Unknown"))
            raise RuntimeError(f"Authentication failed: {error}")

        self._save_cache()
        print("  Authentication successful!\n")
        return result["access_token"]

    def _get_token(self) -> str:
        if not self._token:
            self._token = self._acquire_token()
        return self._token

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _api_request(self, method: str, url: str, **kwargs) -> _requests.Response:
        """Make a Graph API request with timeout, 401 retry, and 429 backoff."""
        import time as _time
        kwargs.setdefault("timeout", 30)
        kwargs["headers"] = self._headers()

        for attempt in range(3):
            resp = getattr(_requests, method)(url, **kwargs)
            if resp.status_code == 401 and attempt == 0:
                self._token = None
                kwargs["headers"] = self._headers()
                continue
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", 2 ** attempt))
                _time.sleep(min(retry_after, 30))
                continue
            break
        resp.raise_for_status()
        return resp

    def _api_get(self, url: str, params: Optional[dict] = None) -> dict:
        return self._api_request("get", url, params=params).json()

    def _api_post(self, url: str, json_body: dict):
        return self._api_request("post", url, json=json_body)

    def upgrade_scopes(self, scopes: List[str]):
        """Re-authenticate with broader scopes if needed."""
        if set(scopes) - set(self.scopes):
            self.scopes = list(set(self.scopes) | set(scopes))
            self._token = None
            self._app = msal.PublicClientApplication(
                self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                token_cache=self._cache,
            )

    # ------------------------------------------------------------------
    # Granted-scope detection (Bolt 4)
    # ------------------------------------------------------------------
    #
    # CRITICAL: Phase 1 does not auto-upgrade write scopes. The scope cliff
    # is surfaced as status="unsupported" so callers know the backend
    # literally cannot do the write — distinct from a 403 at runtime where
    # the scope is granted but a specific operation was refused.

    def has_write_scope(self) -> bool:
        """Return True iff the current client's scopes include Mail.ReadWrite.

        We consult ``self.scopes`` (which reflects the scopes the caller
        intended to request + any explicit ``upgrade_scopes`` calls) rather
        than cracking the access token. That keeps the check cheap and
        avoids pulling a JWT library in just for this.

        Mail.ReadWrite implies Mail.Read and Mail.Send on the Graph
        permission model, so any of ``Mail.ReadWrite`` or ``Mail.ReadWrite.All``
        counts as "can write."
        """
        write_markers = {"mail.readwrite", "mail.readwrite.all"}
        normalized = {str(s).lower() for s in (self.scopes or [])}
        return bool(write_markers & normalized)


# =============================================================================
# OUTLOOK CLIENT
# =============================================================================

class OutlookClient:
    """
    General-purpose Outlook email client.

    Automatically selects the best available backend:
      - Windows: win32com (Outlook COM) — talks to desktop Outlook
      - Mac: AppleScript/JXA — talks to Outlook for Mac (Legacy/Classic mode)
      - Fallback: Microsoft Graph API (REST, needs auth)

    All methods return plain dicts/lists — no COM objects or opaque handles leak out.
    """

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        token_cache_path: Optional[Path] = None,
        backend: Optional[str] = None,
    ):
        """
        Initialize the Outlook client.

        Args:
            client_id: Azure AD app client ID (Graph API only, optional).
            tenant_id: Azure AD tenant ID (Graph API only, optional).
            token_cache_path: Path for Graph API token cache file.
            backend: Force a specific backend: "win32com", "applescript", or "graph".
                     If None, auto-detects based on platform.
        """
        if backend:
            self.backend = backend
        elif HAS_WIN32:
            self.backend = "win32com"
        elif HAS_APPLESCRIPT:
            self.backend = "applescript"
        elif HAS_GRAPH:
            self.backend = "graph"
        else:
            self.backend = None

        if self.backend is None:
            platform_hint = (
                "Install pywin32: pip install pywin32"
                if sys.platform == "win32"
                else "Run Outlook for Mac in Legacy/Classic mode, or install msal + requests"
                if sys.platform == "darwin"
                else "Install msal + requests: pip install msal requests"
            )
            raise RuntimeError(f"No email backend available. {platform_hint}")

        self._graph: Optional[_GraphBackend] = None
        self._applescript: Optional[_AppleScriptBackend] = None

        if self.backend == "graph":
            self._graph = _GraphBackend(
                client_id=client_id,
                tenant_id=tenant_id,
                token_cache_path=token_cache_path,
            )
        elif self.backend == "applescript":
            self._applescript = _AppleScriptBackend()

        # For win32com, we store COM references keyed by email ID for attachment download
        self._win32_msg_cache: Dict[str, Any] = {}

    # -------------------------------------------------------------------------
    # WRITE VERBS (Bolt 4) — move / delete / mark_read
    # -------------------------------------------------------------------------
    #
    # Thin dispatch methods. Each returns an :class:`OperationResult`. The
    # backend-specific interpretation of ``message_id`` matters:
    #
    # * AppleScript: ``"as_msg_<N>"`` style — the numeric index inside
    #   ``source_folder`` (or Inbox). Session-scoped — callers that persist
    #   ids across time must refresh via ``search()``. The optional
    #   ``source_folder`` / ``folder_name`` kwargs below carry the needed
    #   folder context.
    # * COM: the Outlook ``EntryID`` string. Valid across sessions but
    #   invalidated on move (we rewrite to the moved item's new EntryID in
    #   the returned OperationResult).
    # * Graph: the Graph message id. Stable across sessions.

    def move(
        self,
        message_id: str,
        target_folder: str,
        *,
        source_folder: Optional[str] = None,
    ) -> OperationResult:
        """Move a message to ``target_folder``.

        Args:
            message_id: Backend-scoped message identifier (see class notes).
            target_folder: Destination folder name. Will be created if the
                active backend supports detect-and-create (AppleScript/COM).
            source_folder: (AppleScript only) Folder the message lives in.
                Defaults to Inbox. Ignored for COM and Graph where the
                message_id is self-locating.

        Returns:
            :class:`OperationResult` with ``verified=True`` only after a
            round-trip confirms the message landed in ``target_folder``.
        """
        if self.backend == "applescript":
            idx = _parse_as_msg_index(message_id)
            return self._applescript.move(
                msg_index=idx,
                target_folder=target_folder,
                source_folder=source_folder,
            )
        if self.backend == "win32com":
            return self._move_win32(message_id, target_folder)
        if self.backend == "graph":
            return self._move_graph(message_id, target_folder)
        raise RuntimeError(f"Unknown backend: {self.backend}")

    def delete(
        self,
        message_id: str,
        *,
        permanent: bool = False,
        yes_really_permanent: bool = False,
        folder_name: Optional[str] = None,
    ) -> OperationResult:
        """Delete a message — soft-delete (move to Deleted Items) by default.

        Args:
            message_id: Backend-scoped message identifier.
            permanent: If True, bypass Deleted Items and hard-delete the
                message. MUST be paired with ``yes_really_permanent=True`` —
                otherwise this method raises ``ValueError``. The double-flag
                is per-call on purpose (panel F6, unanimous): no amount of
                ``--yes`` global flag survives surviving an operator slip.
            yes_really_permanent: Explicit acknowledgement that hard-delete
                is intentional. Only consulted when ``permanent=True``.
            folder_name: (AppleScript only) Source folder — defaults to Inbox.

        Returns:
            :class:`OperationResult` with ``after_state["permanent"]``
            reflecting which path ran.
        """
        if permanent and not yes_really_permanent:
            raise ValueError(
                "Refusing to hard-delete: permanent=True requires "
                "yes_really_permanent=True to be set in the same call. "
                "This double-flag is intentional — soft-delete (default) "
                "is almost always what you want."
            )
        if self.backend == "applescript":
            idx = _parse_as_msg_index(message_id)
            return self._applescript.delete_message(
                msg_index=idx,
                folder_name=folder_name,
                permanent=permanent,
            )
        if self.backend == "win32com":
            return self._delete_win32(message_id, permanent=permanent)
        if self.backend == "graph":
            return self._delete_graph(message_id, permanent=permanent)
        raise RuntimeError(f"Unknown backend: {self.backend}")

    def mark_read(
        self,
        message_id: str,
        *,
        read: bool = True,
        folder_name: Optional[str] = None,
    ) -> OperationResult:
        """Set the read/unread flag on a message.

        Args:
            message_id: Backend-scoped message identifier.
            read: True to mark read, False to mark unread.
            folder_name: (AppleScript only) Folder the message lives in.
                Defaults to Inbox.

        Returns:
            :class:`OperationResult` with ``verified=True`` only when a
            re-fetch confirms ``after_state["is_read"] == read``.
        """
        if self.backend == "applescript":
            idx = _parse_as_msg_index(message_id)
            return self._applescript.mark_read(
                msg_index=idx,
                folder_name=folder_name,
                read=read,
            )
        if self.backend == "win32com":
            return self._mark_read_win32(message_id, read=read)
        if self.backend == "graph":
            return self._mark_read_graph(message_id, read=read)
        raise RuntimeError(f"Unknown backend: {self.backend}")

    # -------------------------------------------------------------------------
    # SEARCH
    # -------------------------------------------------------------------------

    def search(
        self,
        date_from: Optional[Union[str, datetime]] = None,
        date_to: Optional[Union[str, datetime]] = None,
        subject_contains: Optional[str] = None,
        subject_matches: Optional[str] = None,
        sender_name: Optional[str] = None,
        sender_email: Optional[str] = None,
        sender_domain: Optional[str] = None,
        sender_domains: Optional[List[str]] = None,
        has_attachments: Optional[bool] = None,
        folder: Optional[str] = None,
        folders: Optional[List[str]] = None,
        is_read: Optional[bool] = None,
        body_contains: Optional[str] = None,
        to_contains: Optional[str] = None,
        importance: Optional[str] = None,
        max_results: int = 250,
    ) -> List[Dict]:
        """
        Search Outlook emails with flexible filters.

        All filters are optional and combined with AND logic.

        Args:
            date_from: Start date (inclusive). String "YYYY-MM-DD" or datetime.
            date_to: End date (inclusive). String "YYYY-MM-DD" or datetime.
            subject_contains: Case-insensitive substring match on subject.
            subject_matches: Regex pattern to match against subject.
            sender_name: Case-insensitive substring match on sender display name.
            sender_email: Exact match on sender email address (case-insensitive).
            sender_domain: Filter to emails from this domain.
            sender_domains: Filter to emails from any of these domains.
            has_attachments: If True, only emails with attachments. If False, only without.
            folder: Folder name to search (default: Inbox). E.g., "Sent Items", "Drafts".
            folders: List of folder names to search. Scans each folder and merges
                results sorted by date. E.g., ["Inbox", "Archive", "Snoozed"].
                If both folder and folders are provided, they are merged.
            is_read: If True, only read emails. If False, only unread.
            body_contains: Case-insensitive substring match on email body.
            to_contains: Case-insensitive substring match on To recipients.
            importance: Filter by importance: "high", "normal", "low".
            max_results: Maximum number of results to return (default 250).

        Returns:
            List of email dicts, each containing:
                - id: Unique email identifier
                - subject: Email subject
                - sender_name: Display name of sender
                - sender_email: Email address of sender
                - received_datetime: datetime object
                - received_date: "YYYY-MM-DD" string
                - day_of_week: e.g., "Friday"
                - is_read: bool
                - has_attachments: bool
                - importance: "high", "normal", or "low"
                - body_preview: First 5000 chars of body
                - attachments: List of attachment dicts (id, name, size)
        """
        date_from = _parse_date(date_from) if date_from else None
        date_to = _parse_date(date_to) if date_to else None

        # Merge domain filters
        all_domains = set()
        if sender_domain:
            all_domains.add(sender_domain.lower())
        if sender_domains:
            all_domains.update(d.lower() for d in sender_domains)

        # Merge folder and folders parameters (like sender_domain/sender_domains)
        all_folders: Optional[List[str]] = None
        if folder or folders:
            folder_set: List[str] = []
            if folder:
                folder_set.append(folder)
            if folders:
                for f in folders:
                    if f not in folder_set:
                        folder_set.append(f)
            all_folders = folder_set

        def _search_single_folder(folder_name):
            """Dispatch a single-folder search to the active backend."""
            if self.backend == "win32com":
                return self._search_win32(
                    date_from=date_from, date_to=date_to,
                    subject_contains=subject_contains,
                    has_attachments=has_attachments, folder=folder_name,
                    is_read=is_read, importance=importance,
                    max_results=max_results,
                )
            elif self.backend == "applescript":
                return self._search_applescript(
                    date_from=date_from, date_to=date_to,
                    has_attachments=has_attachments,
                    sender_domains=all_domains if all_domains else None,
                    max_results=max_results, folder=folder_name,
                )
            else:
                return self._search_graph(
                    date_from=date_from, date_to=date_to,
                    subject_contains=subject_contains,
                    has_attachments=has_attachments, folder=folder_name,
                    is_read=is_read, body_contains=body_contains,
                    importance=importance, max_results=max_results,
                )

        if all_folders and len(all_folders) > 1:
            results = []
            for fn in all_folders:
                results.extend(_search_single_folder(fn))
            results.sort(key=lambda e: e["received_datetime"], reverse=True)
        else:
            results = _search_single_folder(
                all_folders[0] if all_folders else None
            )

        # Apply post-filters uniformly in Python.
        # Note: subject_contains is already applied server-side by Graph backend,
        # but not by win32/applescript — apply it here for consistency.
        # subject_matches (regex) is always post-filtered.
        results = self._apply_post_filters(
            results,
            subject_contains=subject_contains if self.backend != "graph" else None,
            subject_matches=subject_matches,
            sender_name=sender_name,
            sender_email=sender_email,
            sender_domains=all_domains if all_domains else None,
            body_contains=body_contains if self.backend != "graph" else None,
            to_contains=to_contains,
        )

        # Truncate AFTER post-filters so filters don't reduce below max_results
        return results[:max_results]

    # -------------------------------------------------------------------------
    # DOWNLOAD ATTACHMENT
    # -------------------------------------------------------------------------

    def download_attachment(
        self,
        email: Dict,
        attachment: Dict,
        output_dir: Optional[Union[str, Path]] = None,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Path:
        """
        Download an attachment from a search result.

        Args:
            email: An email dict returned by search().
            attachment: An attachment dict from email["attachments"].
            output_dir: Directory to save to (uses attachment filename). Mutually
                        exclusive with output_path.
            output_path: Exact file path to save to. Mutually exclusive with output_dir.

        Returns:
            Path to the downloaded file.

        Raises:
            ValueError: If neither output_dir nor output_path is specified.
            RuntimeError: If the download fails.
        """
        if output_path:
            dest = Path(output_path)
        elif output_dir:
            # Sanitize attachment filename to prevent path traversal attacks
            # (e.g., "../../.ssh/authorized_keys" from a malicious sender)
            safe_name = Path(attachment["name"]).name
            if not safe_name:
                safe_name = "unnamed_attachment"
            dest = Path(output_dir) / safe_name
        else:
            raise ValueError("Specify either output_dir or output_path")

        dest.parent.mkdir(parents=True, exist_ok=True)

        if self.backend == "win32com":
            self._download_win32(email, attachment, dest)
        elif self.backend == "applescript":
            self._download_applescript(email, attachment, dest)
        else:
            self._download_graph(email, attachment, dest)

        return dest

    # -------------------------------------------------------------------------
    # SEND
    # -------------------------------------------------------------------------

    def send(
        self,
        to: Union[str, List[str]],
        subject: str,
        body: str,
        cc: Optional[Union[str, List[str]]] = None,
        bcc: Optional[Union[str, List[str]]] = None,
        attachments: Optional[List[Union[str, Path]]] = None,
        html: bool = False,
        importance: Optional[str] = None,
    ) -> bool:
        """
        Send an email via Outlook.

        Args:
            to: Recipient(s). Single email string or list of emails.
            subject: Email subject.
            body: Email body (plain text or HTML depending on html flag).
            cc: CC recipient(s).
            bcc: BCC recipient(s).
            attachments: List of file paths to attach.
            html: If True, body is treated as HTML. Default: plain text.
            importance: "high", "normal", or "low".

        Returns:
            True on success.

        Raises:
            RuntimeError: If sending fails.
            FileNotFoundError: If an attachment file doesn't exist.
        """
        to = [to] if isinstance(to, str) else to
        cc = [cc] if isinstance(cc, str) else (cc or [])
        bcc = [bcc] if isinstance(bcc, str) else (bcc or [])

        # Validate attachment files exist
        att_paths = []
        for att in (attachments or []):
            p = Path(att)
            if not p.exists():
                raise FileNotFoundError(f"Attachment not found: {p}")
            att_paths.append(p)

        if self.backend == "win32com":
            return self._send_win32(to, subject, body, cc, bcc, att_paths, html, importance)
        elif self.backend == "applescript":
            if importance and importance.lower() != "normal":
                import warnings
                warnings.warn(
                    f"Importance '{importance}' is not supported by the AppleScript backend "
                    "and will be ignored. The email will be sent with normal importance.",
                    stacklevel=2,
                )
            return self._send_applescript(to, subject, body, cc, bcc, att_paths, html)
        else:
            return self._send_graph(to, subject, body, cc, bcc, att_paths, html, importance)

    # -------------------------------------------------------------------------
    # CALENDAR EVENTS
    # -------------------------------------------------------------------------

    def get_events(
        self,
        date_from: Optional[Union[str, datetime]] = None,
        date_to: Optional[Union[str, datetime]] = None,
        subject_contains: Optional[str] = None,
        max_results: int = 50,
    ) -> List[Dict]:
        """
        Get calendar events within a date range.

        Args:
            date_from: Start date (inclusive). String "YYYY-MM-DD" or datetime.
                Defaults to today.
            date_to: End date (inclusive). String "YYYY-MM-DD" or datetime.
                Defaults to 7 days from date_from.
            subject_contains: Case-insensitive substring match on event subject.
            max_results: Maximum number of results to return (default 50).

        Returns:
            List of event dicts, each containing:
                - id: Unique event identifier
                - subject: Event title
                - start_datetime: datetime object
                - end_datetime: datetime object
                - start_date: "YYYY-MM-DD" string
                - end_date: "YYYY-MM-DD" string
                - location: Event location
                - organizer_name: Organizer display name
                - organizer_email: Organizer email address
                - is_all_day: bool
                - status: "busy", "free", "tentative", "out of office"
                - body_preview: First 2000 chars of event body
                - attendees: List of dicts (name, email, status)
        """
        start = _parse_date(date_from) if date_from else datetime.now().replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        end = _parse_date(date_to) if date_to else (start + timedelta(days=7))

        if self.backend == "win32com":
            results = self._get_events_win32(start, end, max_results)
        elif self.backend == "applescript":
            results = self._get_events_applescript(start, end, max_results)
        else:
            results = self._get_events_graph(start, end, max_results)

        # Post-filter subject in Python (like search() does)
        if subject_contains:
            term = subject_contains.lower()
            results = [e for e in results if term in e.get("subject", "").lower()]

        return results[:max_results]

    # =========================================================================
    # APPLESCRIPT IMPLEMENTATION
    # =========================================================================

    def _search_applescript(
        self, date_from, date_to, has_attachments, sender_domains, max_results,
        folder=None,
    ) -> List[Dict]:
        """Search emails via AppleScript/JXA (Outlook for Mac Legacy)."""
        start = date_from or datetime(2000, 1, 1)
        end = date_to or datetime.now()
        domain_list = list(sender_domains) if sender_domains else None

        return self._applescript.scan_emails(
            start_date=start,
            end_date=end,
            sender_domains=domain_list,
            only_with_attachments=has_attachments,
            max_results=max_results,
            folder_name=folder,
        )

    def _download_applescript(self, email: Dict, attachment: Dict, dest: Path):
        """Download attachment via AppleScript."""
        msg_idx = attachment.get("_as_msg_index")
        att_idx = attachment.get("_as_att_index")
        folder_name = attachment.get("_as_folder_name")
        if msg_idx is None or att_idx is None:
            raise RuntimeError(
                f"Attachment missing AppleScript indices: {attachment}. "
                "Make sure this attachment came from a search() result."
            )
        self._applescript.save_attachment(msg_idx, att_idx, dest, folder_name=folder_name)

    def _send_applescript(self, to, subject, body, cc, bcc, attachments, html) -> bool:
        """Send email via AppleScript."""
        return self._applescript.send_email(
            to=to, subject=subject, body=body,
            cc=cc, bcc=bcc, attachments=attachments, html=html,
        )

    def _get_events_applescript(self, date_from, date_to, max_results) -> List[Dict]:
        """Get calendar events via AppleScript/JXA."""
        return self._applescript.get_calendar_events(
            start_date=date_from,
            end_date=date_to,
            max_results=max_results,
        )

    # =========================================================================
    # WIN32COM IMPLEMENTATION
    # =========================================================================

    def _get_win32_folder(self, folder_name: Optional[str]):
        """Get an Outlook folder by name via win32com."""
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if not folder_name or folder_name.lower() == "inbox":
            return namespace.GetDefaultFolder(6)  # olFolderInbox

        FOLDER_MAP = {
            "sent items": 5,
            "sent": 5,
            "drafts": 16,
            "deleted items": 3,
            "deleted": 3,
            "outbox": 4,
            "junk": 23,
            "junk email": 23,
        }

        folder_id = FOLDER_MAP.get(folder_name.lower())
        if folder_id:
            return namespace.GetDefaultFolder(folder_id)

        # Try to find by name in top-level folders
        for store in namespace.Folders:
            for f in store.Folders:
                if f.Name.lower() == folder_name.lower():
                    return f

        raise ValueError(f"Folder not found: {folder_name}")

    def _search_win32(
        self,
        date_from, date_to, subject_contains, has_attachments,
        folder, is_read, importance, max_results,
    ) -> List[Dict]:
        # Clear COM object cache from previous searches to prevent memory leaks
        self._win32_msg_cache.clear()
        """Search emails via win32com Outlook COM."""
        outlook_folder = self._get_win32_folder(folder)

        messages = outlook_folder.Items
        messages.Sort("[ReceivedTime]", True)

        # Build Outlook restriction filter
        restrictions = []
        if date_from:
            restrictions.append(f"[ReceivedTime] >= '{date_from.strftime('%m/%d/%Y')}'")
        if date_to:
            end = date_to + timedelta(days=1)
            restrictions.append(f"[ReceivedTime] < '{end.strftime('%m/%d/%Y')}'")
        if is_read is not None:
            restrictions.append(f"[UnRead] = {'False' if is_read else 'True'}")
        if importance:
            imp_map = {"low": 0, "normal": 1, "high": 2}
            if importance.lower() in imp_map:
                restrictions.append(f"[Importance] = {imp_map[importance.lower()]}")

        if restrictions:
            messages = messages.Restrict(" AND ".join(restrictions))

        results = []
        for i, msg in enumerate(messages):
            if len(results) >= max_results:
                break

            try:
                if msg.Class != 43:  # olMail
                    continue

                if has_attachments is True and msg.Attachments.Count == 0:
                    continue
                if has_attachments is False and msg.Attachments.Count > 0:
                    continue

                # Extract sender email, handling Exchange addresses
                sender_email_addr = ""
                try:
                    if msg.SenderEmailType == "EX":
                        sender = msg.Sender
                        if sender:
                            ex_user = sender.GetExchangeUser()
                            if ex_user:
                                sender_email_addr = ex_user.PrimarySmtpAddress
                            else:
                                sender_email_addr = msg.SenderEmailAddress
                    else:
                        sender_email_addr = msg.SenderEmailAddress
                except Exception:
                    sender_email_addr = msg.SenderEmailAddress or ""

                received = msg.ReceivedTime
                received_dt = datetime(
                    received.year, received.month, received.day,
                    received.hour, received.minute, received.second,
                )

                # Extract attachments metadata
                attachments = []
                for j in range(1, msg.Attachments.Count + 1):
                    att = msg.Attachments.Item(j)
                    # Skip embedded/inline items
                    try:
                        if att.Type == 5:  # olEmbeddedItem
                            continue
                    except Exception:
                        pass
                    attachments.append({
                        "id": f"win32_{j}",
                        "name": att.FileName,
                        "size": att.Size,
                        "_win32_index": j,
                    })

                # Extract To recipients
                to_addrs = ""
                try:
                    to_addrs = msg.To or ""
                except Exception:
                    pass

                email_id = f"win32_{id(msg)}"

                email_dict = {
                    "id": email_id,
                    "subject": msg.Subject or "(No Subject)",
                    "sender_name": msg.SenderName or "",
                    "sender_email": sender_email_addr,
                    "received_datetime": received_dt,
                    "received_date": received_dt.strftime("%Y-%m-%d"),
                    "day_of_week": received_dt.strftime("%A"),
                    "is_read": not msg.UnRead,
                    "has_attachments": msg.Attachments.Count > 0,
                    "importance": {0: "low", 1: "normal", 2: "high"}.get(msg.Importance, "normal"),
                    "body_preview": (msg.Body or "")[:5000],
                    "to": to_addrs,
                    "attachments": attachments,
                }

                # Cache the COM object for later download
                self._win32_msg_cache[email_id] = msg

                results.append(email_dict)

                if (i + 1) % 100 == 0:
                    print(f"  Scanned {i + 1} emails...", file=sys.stderr)

            except Exception as e:
                print(f"  Warning: Error processing email {i}: {e}", file=sys.stderr)
                continue

        return results

    def _download_win32(self, email: Dict, attachment: Dict, dest: Path):
        """Download attachment via win32com."""
        msg = self._win32_msg_cache.get(email["id"])
        if msg is None:
            raise RuntimeError(
                f"COM reference expired for email '{email['subject']}'. "
                "Re-run search() to refresh references."
            )

        idx = attachment.get("_win32_index")
        if idx is None:
            raise RuntimeError(f"Attachment missing _win32_index: {attachment}")

        att = msg.Attachments.Item(idx)
        att.SaveAsFile(str(dest))

    def _send_win32(
        self, to, subject, body, cc, bcc, attachments, html, importance,
    ) -> bool:
        """Send email via win32com."""
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem

        mail.To = "; ".join(to)
        mail.Subject = subject

        if cc:
            mail.CC = "; ".join(cc)
        if bcc:
            mail.BCC = "; ".join(bcc)

        if html:
            mail.HTMLBody = body
        else:
            mail.Body = body

        if importance:
            imp_map = {"low": 0, "normal": 1, "high": 2}
            mail.Importance = imp_map.get(importance.lower(), 1)

        for att_path in attachments:
            mail.Attachments.Add(str(att_path.resolve()))

        mail.Send()
        return True

    def _get_events_win32(self, date_from, date_to, max_results) -> List[Dict]:
        """Get calendar events via win32com."""
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar_folder = namespace.GetDefaultFolder(9)  # olFolderCalendar

        end_inclusive = date_to + timedelta(days=1)
        restriction = (
            f"[Start] >= '{date_from.strftime('%m/%d/%Y')}' AND "
            f"[Start] < '{end_inclusive.strftime('%m/%d/%Y')}'"
        )

        items = calendar_folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        items = items.Restrict(restriction)

        print(f"Scanning Outlook calendar: {date_from.strftime('%Y-%m-%d')} to {date_to.strftime('%Y-%m-%d')}...", file=sys.stderr)

        results = []
        for i, item in enumerate(items):
            if len(results) >= max_results:
                break
            try:
                start_dt = datetime(
                    item.Start.year, item.Start.month, item.Start.day,
                    item.Start.hour, item.Start.minute, item.Start.second,
                )
                end_dt = datetime(
                    item.End.year, item.End.month, item.End.day,
                    item.End.hour, item.End.minute, item.End.second,
                )

                organizer = ""
                try:
                    organizer = item.Organizer or ""
                except Exception:
                    pass

                attendees = []
                try:
                    for r in range(1, item.Recipients.Count + 1):
                        recip = item.Recipients.Item(r)
                        response_map = {0: "none", 1: "organizer", 2: "tentative",
                                        3: "accepted", 4: "declined"}
                        attendees.append({
                            "name": recip.Name or "",
                            "email": recip.Address or "",
                            "status": response_map.get(recip.MeetingResponseStatus, "unknown"),
                        })
                except Exception:
                    pass

                status_map = {0: "free", 1: "tentative", 2: "busy", 3: "out of office"}

                results.append({
                    "id": f"win32_evt_{item.EntryID[:16] if item.EntryID else i}",
                    "subject": item.Subject or "(No Subject)",
                    "start_datetime": start_dt,
                    "end_datetime": end_dt,
                    "start_date": start_dt.strftime("%Y-%m-%d"),
                    "end_date": end_dt.strftime("%Y-%m-%d"),
                    "location": item.Location or "",
                    "organizer_name": organizer,
                    "organizer_email": "",
                    "is_all_day": bool(item.AllDayEvent),
                    "status": status_map.get(item.BusyStatus, "busy"),
                    "body_preview": (item.Body or "")[:2000],
                    "attendees": attendees,
                })
            except Exception:
                continue

        print(f"  Found {len(results)} event(s)", file=sys.stderr)
        return results

    # =========================================================================
    # WIN32COM WRITE OPERATIONS (Bolt 4)
    # =========================================================================
    #
    # CRITICAL: Outlook's MailItem.Move() returns the moved item. After the
    # call, the ORIGINAL handle (by EntryID in the source store) is invalid
    # — the message now lives at a different EntryID in the destination
    # store. We MUST capture the returned item for verify. Never re-query
    # the old EntryID.

    def _get_win32_message_by_entry_id(self, entry_id: str):
        """Resolve an Outlook MailItem by EntryID via the MAPI namespace."""
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace.GetItemFromID(entry_id)

    def _win32_message_snapshot(self, msg) -> Dict[str, Any]:
        """Return a dict snapshot of the salient fields on a MailItem."""
        try:
            parent_name = msg.Parent.Name if msg.Parent is not None else ""
        except Exception:
            parent_name = ""
        return {
            "folder": parent_name or "Inbox",
            "is_read": not bool(getattr(msg, "UnRead", False)),
            "subject": getattr(msg, "Subject", "") or "",
            "sender_email": getattr(msg, "SenderEmailAddress", "") or "",
            "entry_id": getattr(msg, "EntryID", "") or "",
        }

    def _mark_read_win32(
        self, entry_id: str, read: bool = True,
    ) -> OperationResult:
        """Mark a message read/unread via COM and verify."""
        start_ms = _time_module.perf_counter()
        operation = "mark_read"
        try:
            msg = self._get_win32_message_by_entry_id(entry_id)
            before = self._win32_message_snapshot(msg)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                error_detail=f"pre-fetch failed: {e}", backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        try:
            msg.UnRead = not bool(read)
            msg.Save()
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail=f"COM error during write: {e}",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        # Round-trip verify. EntryID is still valid (mark_read doesn't move).
        try:
            msg_refetched = self._get_win32_message_by_entry_id(entry_id)
            after = self._win32_message_snapshot(msg_refetched)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail=f"post-verify fetch failed: {e}",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        verified = after["is_read"] == bool(read)
        return OperationResult(
            status="success" if verified else "failed",
            operation=operation, message_id=entry_id, verified=verified,
            before_state=before, after_state=after,
            error_detail=None if verified else "post-state did not match request",
            backend="win32com",
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    def _resolve_win32_folder(self, folder_name: str):
        """Find a target folder by name, creating it if missing on the default store."""
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Default folders handled by _get_win32_folder
        try:
            return self._get_win32_folder(folder_name)
        except Exception:
            pass
        # Otherwise: ensure a top-level folder on the default store.
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
        root = inbox.Parent
        for f in root.Folders:
            if f.Name.lower() == folder_name.lower():
                return f
        # Create it
        return root.Folders.Add(folder_name)

    def _move_win32(
        self, entry_id: str, target_folder: str,
    ) -> OperationResult:
        """Move a message to ``target_folder`` and verify via the Move() return."""
        start_ms = _time_module.perf_counter()
        operation = "move"
        if not target_folder:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                error_detail="target_folder must be a non-empty folder name",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        try:
            msg = self._get_win32_message_by_entry_id(entry_id)
            before = self._win32_message_snapshot(msg)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                error_detail=f"pre-fetch failed: {e}", backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        try:
            target = self._resolve_win32_folder(target_folder)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail=f"folder resolution failed: {e}",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        # CRITICAL (panel-flagged): capture the returned moved item.
        try:
            moved = msg.Move(target)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail=f"COM Move() error: {e}",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        if moved is None:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail="COM Move() returned None (move did not take effect)",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        after = self._win32_message_snapshot(moved)
        verified = after["folder"].lower() == target_folder.lower() or (
            # Resolved folder name may be the canonical one (e.g. "Archive" via
            # a default-folder constant); accept a match on the resolved folder.
            after["folder"].lower() == getattr(target, "Name", target_folder).lower()
        )
        new_entry_id = getattr(moved, "EntryID", None)
        return OperationResult(
            status="success" if verified else "failed",
            operation=operation,
            # Report the new EntryID — the old one is invalidated.
            message_id=new_entry_id or entry_id,
            verified=verified,
            before_state=before, after_state=after,
            error_detail=None if verified else (
                f"post-move verify: moved item's parent={after['folder']!r} "
                f"not {target_folder!r}"
            ),
            backend="win32com",
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    def _delete_win32(
        self, entry_id: str, permanent: bool = False,
    ) -> OperationResult:
        """Delete a message — soft-delete (Deleted Items) by default."""
        start_ms = _time_module.perf_counter()
        operation = "delete"
        if not permanent:
            # Soft-delete is a Move to Deleted Items — same verify path.
            move_res = self._move_win32(entry_id, "Deleted Items")
            return OperationResult(
                status=move_res.status, operation=operation,
                message_id=move_res.message_id, verified=move_res.verified,
                before_state=move_res.before_state,
                after_state={**move_res.after_state, "permanent": False},
                error_detail=move_res.error_detail, backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        # Hard delete.
        try:
            msg = self._get_win32_message_by_entry_id(entry_id)
            before = self._win32_message_snapshot(msg)
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                error_detail=f"pre-fetch failed: {e}", backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        try:
            msg.Delete()
        except Exception as e:
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                before_state=before,
                error_detail=f"COM Delete() error: {e}", backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        # Verify absence — refetch by EntryID should raise.
        try:
            self._get_win32_message_by_entry_id(entry_id)
            # If we got here, the message still exists — silent success.
            return OperationResult(
                status="failed", operation=operation, message_id=entry_id,
                verified=False, before_state=before,
                after_state={"folder": None, "permanent": True},
                error_detail="post-delete verify: EntryID still resolves",
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        except Exception:
            # Can't refetch — that's what we want.
            return OperationResult(
                status="success", operation=operation, message_id=entry_id,
                verified=True, before_state=before,
                after_state={"folder": None, "permanent": True},
                backend="win32com",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )

    # =========================================================================
    # GRAPH API IMPLEMENTATION
    # =========================================================================

    def _search_graph(
        self,
        date_from, date_to, subject_contains, has_attachments,
        folder, is_read, body_contains, importance, max_results,
    ) -> List[Dict]:
        """Search emails via Microsoft Graph API."""
        # Build OData filter
        filters = []
        if date_from:
            filters.append(f"receivedDateTime ge {date_from.strftime('%Y-%m-%dT00:00:00Z')}")
        if date_to:
            end = date_to + timedelta(days=1)
            filters.append(f"receivedDateTime lt {end.strftime('%Y-%m-%dT00:00:00Z')}")
        if has_attachments is not None:
            filters.append(f"hasAttachments eq {'true' if has_attachments else 'false'}")
        if is_read is not None:
            filters.append(f"isRead eq {'true' if is_read else 'false'}")
        if importance:
            filters.append(f"importance eq '{importance.lower()}'")
        if subject_contains:
            # Sanitize for OData: escape single quotes and strip parentheses
            # to prevent OData filter injection
            safe = subject_contains.replace("'", "''")
            safe = re.sub(r"[()]", "", safe)
            filters.append(f"contains(subject, '{safe}')")

        filter_str = " and ".join(filters) if filters else None

        # Determine folder endpoint
        if folder and folder.lower() != "inbox":
            folder_segment = f"mailFolders('{_url_quote(folder, safe='')}')/messages"
        else:
            folder_segment = "messages"

        url = f"{GRAPH_API_BASE}/me/{folder_segment}"
        params = {
            "$select": "id,subject,from,receivedDateTime,hasAttachments,body,isRead,importance,toRecipients",
            "$expand": "attachments($select=id,name,size,contentType)",
            "$orderby": "receivedDateTime desc",
            "$top": min(max_results, 250),
        }
        if filter_str:
            params["$filter"] = filter_str

        all_messages = []
        max_pages = 10
        page = 0
        while url and len(all_messages) < max_results and page < max_pages:
            data = self._graph._api_get(url, params)
            messages = data.get("value", [])
            all_messages.extend(messages)
            url = data.get("@odata.nextLink")
            params = None
            page += 1

        results = []
        for msg in all_messages[:max_results]:
            sender_info = msg.get("from", {}).get("emailAddress", {})
            sender_email_addr = sender_info.get("address", "")

            received_str = msg.get("receivedDateTime", "")
            try:
                received_dt = datetime.fromisoformat(received_str.replace("Z", "+00:00"))
                received_dt = received_dt.replace(tzinfo=None)
            except (ValueError, AttributeError):
                received_dt = datetime.now()

            attachments = []
            for att in msg.get("attachments", []):
                if att.get("@odata.type") == "#microsoft.graph.itemAttachment":
                    continue
                attachments.append({
                    "id": att.get("id", ""),
                    "name": att.get("name", ""),
                    "size": att.get("size", 0),
                    "content_type": att.get("contentType", ""),
                })

            to_addrs = ", ".join(
                r.get("emailAddress", {}).get("address", "")
                for r in msg.get("toRecipients", [])
            )

            results.append({
                "id": msg.get("id", ""),
                "subject": msg.get("subject", "(No Subject)"),
                "sender_name": sender_info.get("name", ""),
                "sender_email": sender_email_addr,
                "received_datetime": received_dt,
                "received_date": received_dt.strftime("%Y-%m-%d"),
                "day_of_week": received_dt.strftime("%A"),
                "is_read": msg.get("isRead", False),
                "has_attachments": msg.get("hasAttachments", False),
                "importance": msg.get("importance", "normal"),
                "body_preview": (msg.get("body", {}).get("content", ""))[:5000],
                "to": to_addrs,
                "attachments": attachments,
            })

        return results

    def _download_graph(self, email: Dict, attachment: Dict, dest: Path):
        """Download attachment via Graph API."""
        msg_id = _url_quote(str(email['id']), safe='')
        att_id = _url_quote(str(attachment['id']), safe='')
        url = f"{GRAPH_API_BASE}/me/messages/{msg_id}/attachments/{att_id}/$value"
        resp = _requests.get(url, headers=self._graph._headers())
        if resp.status_code == 401:
            self._graph._token = None
            resp = _requests.get(url, headers=self._graph._headers())
        resp.raise_for_status()
        dest.write_bytes(resp.content)

    def _send_graph(
        self, to, subject, body, cc, bcc, attachments, html, importance,
    ) -> bool:
        """Send email via Graph API."""
        # Ensure we have Mail.Send scope
        self._graph.upgrade_scopes(GRAPH_SCOPES_SEND)

        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if html else "Text",
                "content": body,
            },
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to
            ],
        }

        if cc:
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc
            ]
        if bcc:
            message["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc
            ]
        if importance:
            message["importance"] = importance.lower()

        # Handle attachments
        if attachments:
            import base64
            message["attachments"] = []
            for att_path in attachments:
                content = att_path.read_bytes()
                message["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att_path.name,
                    "contentBytes": base64.b64encode(content).decode("ascii"),
                })

        url = f"{GRAPH_API_BASE}/me/sendMail"
        self._graph._api_post(url, {"message": message})
        return True

    # =========================================================================
    # GRAPH WRITE OPERATIONS (Bolt 4) — scope-gated
    # =========================================================================
    #
    # Phase 1 does NOT auto-upgrade to Mail.ReadWrite. If the current token
    # lacks that scope we return status="unsupported" so callers see the
    # scope cliff instead of a runtime 403. If the scope IS granted and the
    # API returns 403, we surface a failed status with the 403 detail —
    # scope-missing and permission-denied are NEVER conflated (panel F2).

    _UNSUPPORTED_DETAIL = (
        "OAuth grant lacks Mail.ReadWrite — use COM/AppleScript backend, "
        "or request a scope upgrade and construct OutlookClient with "
        "scopes=['Mail.ReadWrite']"
    )

    def _graph_unsupported(self, operation: str, message_id: Optional[str],
                           elapsed_ms: int) -> OperationResult:
        return OperationResult(
            status="unsupported",
            operation=operation,
            message_id=message_id,
            error_detail=self._UNSUPPORTED_DETAIL,
            backend="graph",
            elapsed_ms=elapsed_ms,
        )

    def _graph_write_request(
        self, method: str, url: str, operation: str, message_id: str,
        start_ms: float, json_body: Optional[Dict[str, Any]] = None,
    ) -> OperationResult:
        """Issue a write request against Graph and translate the response.

        Returns ``OperationResult`` — caller does not need to try/except. On
        4xx we distinguish scope-missing (should never happen past the
        has_write_scope check, but defence in depth) from 403 permission
        denied.
        """
        try:
            kwargs: Dict[str, Any] = {}
            if json_body is not None:
                kwargs["json"] = json_body
            resp = self._graph._api_request(method, url, **kwargs)
        except Exception as e:
            # requests.HTTPError on 4xx; our _api_request calls raise_for_status.
            status_code = getattr(getattr(e, "response", None), "status_code", None)
            body_text = ""
            try:
                body_text = e.response.text  # type: ignore[attr-defined]
            except Exception:
                pass
            if status_code == 403:
                return OperationResult(
                    status="failed", operation=operation, message_id=message_id,
                    error_detail=f"HTTP 403 (permission denied): {body_text or e!s}",
                    backend="graph",
                    elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
                )
            if status_code == 401:
                return OperationResult(
                    status="failed", operation=operation, message_id=message_id,
                    error_detail=f"HTTP 401 (auth failed): {body_text or e!s}",
                    backend="graph",
                    elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
                )
            return OperationResult(
                status="failed", operation=operation, message_id=message_id,
                error_detail=f"Graph request error: {e}",
                backend="graph",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        return OperationResult(
            status="success", operation=operation, message_id=message_id,
            verified=True,  # Graph returns 2xx only when it actually applied
            backend="graph",
            after_state={"http_status": resp.status_code},
            elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
        )

    def _mark_read_graph(
        self, message_id: str, read: bool = True,
    ) -> OperationResult:
        """Mark a Graph message as read/unread.

        Returns status="unsupported" unless the current token has Mail.ReadWrite.
        """
        start_ms = _time_module.perf_counter()
        operation = "mark_read"
        if not self._graph.has_write_scope():
            return self._graph_unsupported(
                operation, message_id,
                int((_time_module.perf_counter() - start_ms) * 1000),
            )
        msg_id_quoted = _url_quote(str(message_id), safe="")
        url = f"{GRAPH_API_BASE}/me/messages/{msg_id_quoted}"
        return self._graph_write_request(
            "patch", url, operation, message_id, start_ms,
            json_body={"isRead": bool(read)},
        )

    def _move_graph(
        self, message_id: str, target_folder: str,
    ) -> OperationResult:
        """Move a Graph message to ``target_folder`` (well-known name or ID).

        Returns status="unsupported" unless the current token has Mail.ReadWrite.
        """
        start_ms = _time_module.perf_counter()
        operation = "move"
        if not self._graph.has_write_scope():
            return self._graph_unsupported(
                operation, message_id,
                int((_time_module.perf_counter() - start_ms) * 1000),
            )
        if not target_folder:
            return OperationResult(
                status="failed", operation=operation, message_id=message_id,
                error_detail="target_folder must be a non-empty folder name",
                backend="graph",
                elapsed_ms=int((_time_module.perf_counter() - start_ms) * 1000),
            )
        msg_id_quoted = _url_quote(str(message_id), safe="")
        url = f"{GRAPH_API_BASE}/me/messages/{msg_id_quoted}/move"
        return self._graph_write_request(
            "post", url, operation, message_id, start_ms,
            json_body={"destinationId": target_folder},
        )

    def _delete_graph(
        self, message_id: str, permanent: bool = False,
    ) -> OperationResult:
        """Delete a Graph message (soft-delete via move to deleteditems by default).

        Returns status="unsupported" unless the current token has Mail.ReadWrite.
        """
        start_ms = _time_module.perf_counter()
        operation = "delete"
        if not self._graph.has_write_scope():
            return self._graph_unsupported(
                operation, message_id,
                int((_time_module.perf_counter() - start_ms) * 1000),
            )
        msg_id_quoted = _url_quote(str(message_id), safe="")
        if not permanent:
            # Soft-delete = move to deleteditems well-known folder.
            url = f"{GRAPH_API_BASE}/me/messages/{msg_id_quoted}/move"
            result = self._graph_write_request(
                "post", url, operation, message_id, start_ms,
                json_body={"destinationId": "deleteditems"},
            )
            if result.status == "success":
                result.after_state = {**result.after_state, "permanent": False}
            return result
        # Hard delete.
        url = f"{GRAPH_API_BASE}/me/messages/{msg_id_quoted}"
        result = self._graph_write_request(
            "delete", url, operation, message_id, start_ms,
        )
        if result.status == "success":
            result.after_state = {**result.after_state, "permanent": True, "folder": None}
        return result

    def _get_events_graph(self, date_from, date_to, max_results) -> List[Dict]:
        """Get calendar events via Microsoft Graph API."""
        start_str = date_from.strftime("%Y-%m-%dT00:00:00Z")
        end_str = (date_to + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00Z")

        params = {
            "startdatetime": start_str,
            "enddatetime": end_str,
            "$select": "id,subject,start,end,location,organizer,isAllDay,showAs,body,attendees",
            "$orderby": "start/dateTime",
            "$top": min(max_results, 250),
        }

        url = f"{GRAPH_API_BASE}/me/calendarview"
        query = "&".join(f"{k}={_url_quote(str(v), safe='/:,')}" for k, v in params.items())
        full_url = f"{url}?{query}"

        print(f"Scanning Outlook calendar: {date_from.strftime('%Y-%m-%d')} to {date_to.strftime('%Y-%m-%d')}...", file=sys.stderr)

        resp = self._graph._api_get(full_url)
        events = resp.get("value", [])

        results = []
        for evt in events[:max_results]:
            try:
                start_dt = datetime.fromisoformat(
                    evt["start"]["dateTime"].replace("Z", "+00:00")
                ).replace(tzinfo=None)
            except (ValueError, KeyError):
                start_dt = datetime.now()

            try:
                end_dt = datetime.fromisoformat(
                    evt["end"]["dateTime"].replace("Z", "+00:00")
                ).replace(tzinfo=None)
            except (ValueError, KeyError):
                end_dt = start_dt

            organizer = evt.get("organizer", {}).get("emailAddress", {})
            show_as_map = {"free": "free", "tentative": "tentative",
                           "busy": "busy", "oof": "out of office",
                           "workingElsewhere": "working elsewhere"}

            attendees = []
            for att in evt.get("attendees", []):
                ea = att.get("emailAddress", {})
                resp_status = att.get("status", {}).get("response", "unknown")
                attendees.append({
                    "name": ea.get("name", ""),
                    "email": ea.get("address", ""),
                    "status": resp_status,
                })

            results.append({
                "id": evt.get("id", ""),
                "subject": evt.get("subject", "(No Subject)"),
                "start_datetime": start_dt,
                "end_datetime": end_dt,
                "start_date": start_dt.strftime("%Y-%m-%d"),
                "end_date": end_dt.strftime("%Y-%m-%d"),
                "location": evt.get("location", {}).get("displayName", ""),
                "organizer_name": organizer.get("name", ""),
                "organizer_email": organizer.get("address", ""),
                "is_all_day": evt.get("isAllDay", False),
                "status": show_as_map.get(evt.get("showAs", ""), "busy"),
                "body_preview": evt.get("body", {}).get("content", "")[:2000],
                "attendees": attendees,
            })

        print(f"  Found {len(results)} event(s)", file=sys.stderr)
        return results

    # =========================================================================
    # POST-FILTERS (applied in Python, after backend fetch)
    # =========================================================================

    @staticmethod
    def _apply_post_filters(
        results: List[Dict],
        subject_contains: Optional[str] = None,
        subject_matches: Optional[str] = None,
        sender_name: Optional[str] = None,
        sender_email: Optional[str] = None,
        sender_domains: Optional[set] = None,
        body_contains: Optional[str] = None,
        to_contains: Optional[str] = None,
    ) -> List[Dict]:
        """Apply Python-side filters that can't be pushed to the backend."""
        filtered = results

        # subject_contains — win32 doesn't support server-side, Graph does but
        # we double-check here for consistency
        if subject_contains:
            term = subject_contains.lower()
            filtered = [e for e in filtered if term in e["subject"].lower()]

        if subject_matches:
            pat = re.compile(subject_matches, re.IGNORECASE)
            filtered = [e for e in filtered if pat.search(e["subject"])]

        if sender_name:
            term = sender_name.lower()
            filtered = [e for e in filtered if term in e["sender_name"].lower()]

        if sender_email:
            term = sender_email.lower()
            filtered = [e for e in filtered if e["sender_email"].lower() == term]

        if sender_domains:
            filtered = [
                e for e in filtered
                if _extract_domain(e["sender_email"]) in sender_domains
            ]

        if body_contains:
            term = body_contains.lower()
            filtered = [e for e in filtered if term in e.get("body_preview", "").lower()]

        if to_contains:
            term = to_contains.lower()
            filtered = [e for e in filtered if term in e.get("to", "").lower()]

        return filtered


# =============================================================================
# HELPERS
# =============================================================================

def _parse_date(val: Union[str, datetime]) -> datetime:
    """Parse a date string or pass through a datetime."""
    if isinstance(val, datetime):
        return val
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {val}. Use YYYY-MM-DD format.")


def _extract_domain(email_addr: str) -> str:
    """Extract lowercase domain from an email address."""
    if "@" in email_addr:
        return email_addr.split("@")[-1].lower()
    return ""


def _parse_as_msg_index(message_id: str) -> int:
    """Parse the ``as_msg_<N>`` form used by the AppleScript backend.

    Accepts both ``"as_msg_42"`` (as returned by ``search()``) and a bare
    numeric string ``"42"``. Raises ``ValueError`` on anything else.
    """
    s = str(message_id).strip()
    if s.startswith("as_msg_"):
        s = s[len("as_msg_"):]
    try:
        return int(s)
    except ValueError:
        raise ValueError(
            f"AppleScript message_id must be of the form 'as_msg_<N>' or a "
            f"numeric string; got {message_id!r}"
        )
