#!/usr/bin/env python3
"""
Outlook Tool — General-purpose Outlook email client for Python.

Cross-platform wrapper around Microsoft Outlook that supports:
  - Searching/filtering emails by date, subject, sender, domain, folder, etc.
  - Downloading attachments from matched emails
  - Sending emails with optional attachments

Platform detection:
  - Windows: win32com (Outlook COM automation) — talks directly to desktop Outlook
  - Mac/Linux: Microsoft Graph API via MSAL device code auth

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

import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# =============================================================================
# PLATFORM DETECTION
# =============================================================================

HAS_WIN32 = False
HAS_GRAPH = False

if sys.platform == "win32":
    try:
        import win32com.client
        HAS_WIN32 = True
    except ImportError:
        pass

if not HAS_WIN32:
    try:
        import msal
        import requests as _requests
        HAS_GRAPH = True
    except ImportError:
        HAS_GRAPH = False


# =============================================================================
# GRAPH API BACKEND
# =============================================================================

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES_READ = ["Mail.Read"]
GRAPH_SCOPES_SEND = ["Mail.Read", "Mail.Send"]
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
            self.token_cache_path.write_text(self._cache.serialize())

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

    def _api_get(self, url: str, params: Optional[dict] = None) -> dict:
        resp = _requests.get(url, headers=self._headers(), params=params)
        if resp.status_code == 401:
            self._token = None
            resp = _requests.get(url, headers=self._headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    def _api_post(self, url: str, json_body: dict) -> _requests.Response:
        resp = _requests.post(url, headers=self._headers(), json=json_body)
        if resp.status_code == 401:
            self._token = None
            resp = _requests.post(url, headers=self._headers(), json=json_body)
        resp.raise_for_status()
        return resp

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


# =============================================================================
# OUTLOOK CLIENT
# =============================================================================

class OutlookClient:
    """
    General-purpose Outlook email client.

    Automatically selects the best available backend:
      - Windows: win32com (Outlook COM)
      - Mac/Linux: Microsoft Graph API

    All methods return plain dicts/lists — no COM objects or opaque handles leak out.
    """

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        token_cache_path: Optional[Path] = None,
    ):
        """
        Initialize the Outlook client.

        Args:
            client_id: Azure AD app client ID (Graph API only, optional).
            tenant_id: Azure AD tenant ID (Graph API only, optional).
            token_cache_path: Path for Graph API token cache file.
        """
        self.backend = "win32com" if HAS_WIN32 else ("graph" if HAS_GRAPH else None)

        if self.backend is None:
            platform_hint = (
                "Install pywin32: pip install pywin32"
                if sys.platform == "win32"
                else "Install msal + requests: pip install msal requests"
            )
            raise RuntimeError(f"No email backend available. {platform_hint}")

        self._graph: Optional[_GraphBackend] = None
        if self.backend == "graph":
            self._graph = _GraphBackend(
                client_id=client_id,
                tenant_id=tenant_id,
                token_cache_path=token_cache_path,
            )

        # For win32com, we store COM references keyed by email ID for attachment download
        self._win32_msg_cache: Dict[str, Any] = {}

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

        if self.backend == "win32com":
            results = self._search_win32(
                date_from=date_from,
                date_to=date_to,
                subject_contains=subject_contains,
                has_attachments=has_attachments,
                folder=folder,
                is_read=is_read,
                importance=importance,
                max_results=max_results,
            )
        else:
            results = self._search_graph(
                date_from=date_from,
                date_to=date_to,
                subject_contains=subject_contains,
                has_attachments=has_attachments,
                folder=folder,
                is_read=is_read,
                body_contains=body_contains,
                importance=importance,
                max_results=max_results,
            )

        # Apply post-filters that both backends handle uniformly in Python
        results = self._apply_post_filters(
            results,
            subject_contains=subject_contains,
            subject_matches=subject_matches,
            sender_name=sender_name,
            sender_email=sender_email,
            sender_domains=all_domains if all_domains else None,
            body_contains=body_contains,
            to_contains=to_contains,
        )

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
            dest = Path(output_dir) / attachment["name"]
        else:
            raise ValueError("Specify either output_dir or output_path")

        dest.parent.mkdir(parents=True, exist_ok=True)

        if self.backend == "win32com":
            self._download_win32(email, attachment, dest)
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
        else:
            return self._send_graph(to, subject, body, cc, bcc, att_paths, html, importance)

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
                    print(f"  Scanned {i + 1} emails...")

            except Exception as e:
                print(f"  Warning: Error processing email {i}: {e}")
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
            safe = subject_contains.replace("'", "''")
            filters.append(f"contains(subject, '{safe}')")

        filter_str = " and ".join(filters) if filters else None

        # Determine folder endpoint
        if folder and folder.lower() != "inbox":
            folder_segment = f"mailFolders('{folder}')/messages"
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
        while url and len(all_messages) < max_results:
            data = self._graph._api_get(url, params)
            messages = data.get("value", [])
            all_messages.extend(messages)
            url = data.get("@odata.nextLink")
            params = None

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
        url = (
            f"{GRAPH_API_BASE}/me/messages/{email['id']}"
            f"/attachments/{attachment['id']}/$value"
        )
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
