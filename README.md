# Outlook Tool

General-purpose Python wrapper around Microsoft Outlook for searching emails, downloading attachments, and sending emails — from scripts or the command line.

## Platform Support

| Feature | Windows | Mac / Linux |
|---|---|---|
| Search emails | win32com (Outlook COM) | Microsoft Graph API |
| Download attachments | win32com SaveAsFile | Graph API REST |
| Send emails | win32com CreateItem | Graph API sendMail |

**Windows** talks directly to the Outlook desktop app via COM automation. No API keys needed.

**Mac/Linux** uses the Microsoft Graph API with MSAL device code authentication. First run opens a browser for sign-in; token is cached for ~90 days.

## Setup

### Windows

```bash
git clone https://github.com/chwinston/outlook-tool.git
cd outlook-tool
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

That's it. `pywin32` installs automatically. Make sure Outlook is running and signed in.

### Mac / Linux

```bash
git clone https://github.com/chwinston/outlook-tool.git
cd outlook-tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On first use, you'll be prompted to authenticate via device code (opens browser). The token is cached at `~/.outlook_tool_token_cache.bin`.

**Default auth** uses Microsoft Office's public client ID — no Azure app registration needed. If your organization requires a registered app:

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `Outlook Tool`
3. Account types: **Single tenant**
4. Redirect URI: leave blank
5. After registration:
   - **Authentication** → **Advanced settings** → enable **Allow public client flows** → Save
   - **API Permissions** → **Microsoft Graph** → **Delegated** → add `Mail.Read` and `Mail.Send` → **Grant admin consent**
6. Copy the **Application (client) ID** and **Directory (tenant) ID**
7. Set environment variables:
   ```bash
   export KPI_GRAPH_CLIENT_ID=<your-client-id>
   export KPI_GRAPH_TENANT_ID=<your-tenant-id>
   ```

## Usage

### Python API

```python
from outlook_tool import OutlookClient

client = OutlookClient()

# Search by date range
emails = client.search(date_from="2026-03-01", date_to="2026-03-15")

# Search with multiple filters (AND logic)
emails = client.search(
    date_from="2026-03-01",
    sender_domain="example.com",
    subject_contains="quarterly report",
    has_attachments=True,
    is_read=False,
)

# Search by sender
emails = client.search(sender_email="boss@example.com")
emails = client.search(sender_name="Jane")
emails = client.search(sender_domains=["example.com", "partner.org"])

# Search by subject regex
emails = client.search(subject_matches=r"Q[1-4] \d{4} Report")

# Search specific folder
emails = client.search(folder="Sent Items", date_from="2026-03-01")

# Search by importance
emails = client.search(importance="high", is_read=False)

# Search by body content or recipients
emails = client.search(body_contains="action required")
emails = client.search(to_contains="team@example.com")
```

### Download Attachments

```python
emails = client.search(has_attachments=True, sender_domain="example.com")

for email in emails:
    for att in email["attachments"]:
        # Save to directory (uses original filename)
        path = client.download_attachment(email, att, output_dir="./downloads")
        print(f"Saved: {path}")

        # Or save to exact path
        path = client.download_attachment(email, att, output_path="./report.xlsx")
```

### Send Emails

```python
# Simple text email
client.send(
    to="colleague@example.com",
    subject="Meeting notes",
    body="Here are the notes from today's meeting.",
)

# HTML email with attachments
client.send(
    to=["team@example.com", "manager@example.com"],
    subject="Q1 Report",
    body="<h1>Q1 Report</h1><p>Please review the attached report.</p>",
    html=True,
    attachments=["report.pdf", "data.xlsx"],
    cc="stakeholder@example.com",
    importance="high",
)
```

### CLI

```bash
# Search by date range
python cli.py search --from 2026-03-01 --to-date 2026-03-15

# Search with filters
python cli.py search --domain example.com --subject "report" --has-attachments

# Search and download attachments
python cli.py search --sender-email boss@example.com --download ./attachments

# Search unread emails
python cli.py search --unread --importance high

# Output as JSON
python cli.py search --from 2026-03-01 --json

# Send an email
python cli.py send --to user@example.com --subject "Hello" --body "Hi there"

# Send with attachments
python cli.py send --to user@example.com --subject "Report" --body "Attached." --attach report.pdf data.xlsx
```

## Search Filters Reference

All filters are optional and combined with AND logic.

| Filter | Python API | CLI | Description |
|---|---|---|---|
| Date from | `date_from="2026-03-01"` | `--from 2026-03-01` | Start date (inclusive) |
| Date to | `date_to="2026-03-15"` | `--to-date 2026-03-15` | End date (inclusive) |
| Subject contains | `subject_contains="report"` | `--subject "report"` | Case-insensitive substring |
| Subject regex | `subject_matches=r"Q\d"` | `--subject-regex "Q\d"` | Regex pattern |
| Sender name | `sender_name="Jane"` | `--sender-name "Jane"` | Display name substring |
| Sender email | `sender_email="j@ex.com"` | `--sender-email j@ex.com` | Exact email match |
| Sender domain | `sender_domain="ex.com"` | `--domain ex.com` | Single domain |
| Sender domains | `sender_domains=["a.com"]` | `--domains a.com b.com` | Multiple domains |
| Has attachments | `has_attachments=True` | `--has-attachments` | Only with attachments |
| Folder | `folder="Sent Items"` | `--folder "Sent Items"` | Folder name |
| Read status | `is_read=False` | `--unread` / `--read` | Read/unread filter |
| Body contains | `body_contains="urgent"` | `--body "urgent"` | Body text substring |
| To contains | `to_contains="team@"` | `--to "team@"` | Recipient substring |
| Importance | `importance="high"` | `--importance high` | high/normal/low |
| Max results | `max_results=100` | `--max-results 100` | Limit results |

## Email Dict Structure

Every email returned by `search()` has this shape:

```python
{
    "id": "...",                     # Unique identifier
    "subject": "Q1 Report",
    "sender_name": "Jane Smith",
    "sender_email": "jane@example.com",
    "received_datetime": datetime(...),
    "received_date": "2026-03-10",
    "day_of_week": "Tuesday",
    "is_read": True,
    "has_attachments": True,
    "importance": "normal",
    "body_preview": "First 5000 chars...",
    "to": "team@example.com",
    "attachments": [
        {"id": "...", "name": "report.pdf", "size": 102400},
    ],
}
```

## Architecture

```
outlook_tool.py    # Core library — OutlookClient class
cli.py             # Command-line interface
requirements.txt   # Dependencies (cross-platform)
tests/             # Unit tests
```

The `OutlookClient` auto-detects the platform at import time and selects the best backend. All methods return plain Python dicts — no COM objects or opaque handles leak out.
