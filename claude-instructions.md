# Outlook Tool — Instructions for Claude

This is a general-purpose Outlook email tool. You are likely being asked to help a user search emails, download attachments, or send emails via Outlook. This file tells you everything you need to set up and use the tool.

## Architecture

```
outlook_tool.py    # Core library — OutlookClient class (search, download, send)
cli.py             # Command-line interface wrapping OutlookClient
pyproject.toml     # Package config — enables `pip install -e .` for global use
requirements.txt   # Dependencies
tests/             # Unit tests (pytest)
```

The tool auto-detects the platform and selects the right backend:
1. **Windows** → win32com (Outlook COM automation)
2. **Mac** → AppleScript/JXA (talks to Outlook for Mac in Legacy/Classic mode)
3. **Fallback** → Microsoft Graph API (needs MSAL auth)

All backends produce the same output format. No API keys are needed for Windows or Mac — the tool talks directly to the local Outlook app.

## Setup for a New User

### Check prerequisites first

1. **Python 3.10+** must be installed and on PATH
2. **Outlook desktop app** must be running and signed in
3. **Mac only:** Outlook must be in Legacy/Classic mode (not "New Outlook")

### Installation

If the user already has the repo cloned, set up from their clone. Otherwise:

```bash
# Clone
git clone https://github.com/chwinston/outlook-tool.git
cd outlook-tool
```

**Option A — Project-local install** (use from inside the outlook-tool folder only):
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Mac
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**Option B — Global install** (recommended — use from any folder on the system):
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate
pip install -e .

# Mac
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

The `-e .` (editable install) registers `outlook_tool` as an importable module and `outlook-tool` as a CLI command within the virtual environment. The user can then use it from any working directory as long as the venv is activated.

### Verify it works

```bash
outlook-tool search --from 2026-03-20 --to-date 2026-03-26 --max-results 3
```

Or in Python:
```python
from outlook_tool import OutlookClient
client = OutlookClient()
emails = client.search(date_from="2026-03-20", date_to="2026-03-26", max_results=3)
for e in emails:
    print(f"{e['received_date']} — {e['sender_name']}: {e['subject']}")
```

If this returns emails, setup is complete.

## Using the Tool from Other Projects

After `pip install -e .`, the user does NOT need to be in the outlook-tool directory. They just need the virtual environment active.

### From the CLI
```bash
# From any folder — works because outlook-tool is on the venv PATH
outlook-tool search --from 2026-03-01 --subject "report"
```

### From Python scripts in other projects
```python
# This works from any Python file, as long as the venv with outlook-tool is active
from outlook_tool import OutlookClient

client = OutlookClient()
emails = client.search(sender_domain="example.com", has_attachments=True)
```

### If the user wants a system-wide install (no venv activation needed)
Install directly into the system Python — use with caution:
```bash
# Windows
pip install -e C:\Users\USERNAME\Projects\outlook-tool

# Mac
pip install -e ~/Projects/outlook-tool
```

This makes `outlook-tool` and `from outlook_tool import OutlookClient` available everywhere without activating any virtual environment.

## How to Use the Tool

### Search emails

The `search()` method and `outlook-tool search` CLI accept these filters (all optional, AND logic):

| Filter | Python kwarg | CLI flag |
|---|---|---|
| Date range | `date_from=`, `date_to=` | `--from`, `--to-date` |
| Subject keyword | `subject_contains=` | `--subject` |
| Subject regex | `subject_matches=` | `--subject-regex` |
| Sender name | `sender_name=` | `--sender-name` |
| Sender email | `sender_email=` | `--sender-email` |
| Sender domain | `sender_domain=` | `--domain` |
| Multiple domains | `sender_domains=[]` | `--domains` |
| Has attachments | `has_attachments=True` | `--has-attachments` |
| Folder | `folder=` | `--folder` |
| Multiple folders | `folders=[]` | `--folders` |
| Read/unread | `is_read=` | `--unread` / `--read` |
| Body keyword | `body_contains=` | `--body` |
| Recipient | `to_contains=` | `--to` |
| Importance | `importance=` | `--importance` |
| Max results | `max_results=` | `--max-results` |

### Download attachments

```python
emails = client.search(has_attachments=True, sender_domain="example.com")
for email in emails:
    for att in email["attachments"]:
        client.download_attachment(email, att, output_dir="./downloads")
```

CLI: `outlook-tool search --has-attachments --download ./downloads`

### Send emails

```python
client.send(
    to=["user@example.com"],
    subject="Subject line",
    body="Email body text",
    attachments=["file.pdf"],  # optional
    cc=["cc@example.com"],     # optional
    html=False,                # set True for HTML body
    importance="normal",       # high/normal/low
)
```

CLI: `outlook-tool send --to user@example.com --subject "Hi" --body "Hello" --attach file.pdf`

### Email dict structure

Every email from `search()` returns:
```python
{
    "id": "...",
    "subject": "...",
    "sender_name": "Jane Smith",
    "sender_email": "jane@example.com",
    "received_datetime": datetime(...),  # Python datetime object
    "received_date": "2026-03-10",       # String
    "day_of_week": "Tuesday",
    "is_read": True,
    "has_attachments": True,
    "importance": "normal",
    "body_preview": "First 5000 chars...",
    "to": "recipient@example.com",
    "attachments": [
        {"id": "...", "name": "report.pdf", "size": 102400},
    ],
}
```

## Troubleshooting

| Problem | Fix |
|---|---|
| `No email backend available` | Windows: `pip install pywin32`. Mac: verify Outlook is running in Legacy mode. |
| `Cannot connect to Outlook for Mac` | Open Outlook, switch to Legacy/Classic mode (toggle in top-right corner). |
| macOS permissions popup | The user must click "Allow" to let Terminal/Claude control Outlook. If denied, go to System Settings > Privacy & Security > Automation. |
| `ModuleNotFoundError: outlook_tool` | The venv isn't activated, or the tool wasn't installed with `pip install -e .` |
| No emails returned | Widen the date range. Check that Outlook has synced the emails. Try `--max-results 5` with just `--from`. |

## Testing

```bash
# From the outlook-tool directory, with venv active
pytest tests/ -v
```

43+ unit tests covering helpers, post-filters, platform detection, and input validation. Tests do not require Outlook — they mock the backends.

## Key Design Decisions

- **AppleScript preferred over Graph API on Mac** — talks to the local Outlook app, avoids IT/Azure AD auth barriers.
- **No API keys needed on Windows or Mac** — win32com and AppleScript both talk directly to the installed Outlook app.
- **All search results are plain dicts** — no COM objects or opaque handles leak out. Results are safe to serialize, print, or pass around.
- **Post-filters in Python** — some filters (subject regex, sender name, body contains) are applied in Python after the backend fetch, so they work identically across all backends.
- **`backend=` kwarg** — advanced users can force a specific backend: `OutlookClient(backend="graph")`.
- **Multi-folder search** — `folders=["Inbox", "Archive", "Snoozed"]` searches multiple folders and merges results sorted by date. Follows the same pattern as `sender_domain`/`sender_domains`. Attachment downloads are folder-aware (tracked via `_as_folder_name`).
- **Calendar support** — `get_events()` method and `outlook-tool events` CLI command for searching calendar events by date range across all three backends.
