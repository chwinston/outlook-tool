# Requirements — Outlook Tool

## Purpose

A cross-platform CLI and Python library for searching, downloading attachments from, and sending emails via Microsoft Outlook. Designed for solo developers and AI assistants who need programmatic email access without cloud API setup.

## Scope

### What this tool does

| ID | Requirement | Priority |
|----|-------------|----------|
| REQ-001 | Search emails by date range, sender, subject, folder, attachments, read status, importance, and body content | **Must** |
| REQ-002 | Search across multiple folders in a single call (Inbox, Archive, Snoozed, Sent Items, etc.) | **Must** |
| REQ-003 | Download email attachments to a specified directory | **Must** |
| REQ-004 | Send emails with optional attachments, CC/BCC, HTML body, and importance level | **Must** |
| REQ-005 | Provide both a CLI (`outlook-tool`) and Python API (`OutlookClient`) | **Must** |
| REQ-006 | Auto-detect platform and select the appropriate backend | **Must** |
| REQ-007 | Output search results as formatted text or JSON | **Should** |
| REQ-008 | Search calendar events by date range, subject, and organizer | **Should** |
| REQ-009 | List calendars available to the signed-in user | **Could** |
| REQ-010 | Create calendar events and send meeting invitations | **Could** |

### What this tool does NOT do

- Modify or delete existing emails
- Access contacts or tasks
- Run as a background service or daemon
- Provide a web interface or GUI
- Store or cache email data between sessions
- Require cloud API keys for Windows or Mac usage

## Platform Support

| Platform | Backend | Requirements |
|----------|---------|-------------|
| Windows | Win32 COM (pywin32) | Outlook desktop running, signed in |
| macOS | AppleScript/JXA | Outlook for Mac in Legacy/Classic mode, running, signed in |
| Linux/Other | Microsoft Graph API | MSAL device code auth, internet access |

## Non-Functional Requirements

- **NFR-001**: Python 3.10+ required
- **NFR-002**: No API keys or Azure setup needed for Windows/Mac backends
- **NFR-003**: All search results returned as plain Python dicts (no opaque handles)
- **NFR-004**: CLI installable via `pip install -e .` for global access
- **NFR-005**: Test suite runs without Outlook installed (mocked backends)

## Security Requirements

- **REQ-SEC-001**: No command injection via crafted email data (AppleScript sanitization)
- **REQ-SEC-002**: Graph API token cache stored with restrictive permissions (0600)
- **REQ-SEC-003**: No email content persisted to disk beyond user-requested attachment downloads

*Priority: Must (required), Should (important), Could (nice-to-have), Won't (out of scope)*
