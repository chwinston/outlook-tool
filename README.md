# Outlook Tool

Search your Outlook emails, download attachments, and send emails — all from Python scripts or the command line. Works on both Windows and Mac.

---

## How It Works

This tool talks **directly to your Outlook desktop app** — no cloud API keys, no IT permissions, no Azure setup needed.

| Platform | How it connects | Requirements |
|---|---|---|
| **Windows** | Outlook COM (win32com) | Outlook desktop app, running and signed in |
| **Mac** | AppleScript/JXA | Outlook for Mac in **Legacy/Classic mode**, running and signed in |

> **Mac users:** You must be running Outlook in **Legacy (Classic) mode**, not the "New Outlook." To check: open Outlook, look for the toggle in the top-right corner that says "New Outlook" — make sure it's **off**. If you don't see the toggle, you're already in Legacy mode.

---

## Getting Started (Windows)

### Prerequisites

- **Python 3.10 or newer** — [Download here](https://www.python.org/downloads/) if you don't have it
  - During installation, **check the box** that says "Add Python to PATH"
- **Outlook desktop app** — must be open and signed in to your email account
- **Git** — [Download here](https://git-scm.com/downloads) if you don't have it

### Step-by-Step Installation

1. **Open a terminal** (Command Prompt or PowerShell):
   - Press `Win + R`, type `cmd`, and press Enter

2. **Navigate to where you want the tool** (pick a folder you'll remember):
   ```
   cd C:\Users\YourName\Projects
   ```

3. **Download the tool from GitHub:**
   ```
   git clone https://github.com/chwinston/outlook-tool.git
   ```

4. **Go into the folder:**
   ```
   cd outlook-tool
   ```

5. **Create a Python virtual environment** (keeps dependencies isolated):
   ```
   python -m venv .venv
   ```

6. **Activate the virtual environment:**
   ```
   .venv\Scripts\activate
   ```
   You should see `(.venv)` appear at the start of your terminal prompt.

7. **Install the tool:**
   ```
   pip install -e .
   ```
   This installs the tool so you can use it from any folder on your computer (not just the outlook-tool folder).

8. **Verify it works:**
   ```
   outlook-tool search --from 2026-03-01 --to-date 2026-03-26 --max-results 5
   ```
   You should see your 5 most recent emails from that date range printed out.

That's it — you're ready to go.

---

## Getting Started (Mac)

### Prerequisites

- **Python 3.10 or newer** (macOS usually has Python 3 pre-installed — check with `python3 --version`)
- **Outlook for Mac** in **Legacy/Classic mode**, running and signed in
- **Git** (pre-installed on macOS)

### Step-by-Step Installation

1. **Open Terminal** (press `Cmd + Space`, type "Terminal", press Enter)

2. **Navigate to where you want the tool:**
   ```
   cd ~/Projects
   ```

3. **Download the tool:**
   ```
   git clone https://github.com/chwinston/outlook-tool.git
   ```

4. **Go into the folder:**
   ```
   cd outlook-tool
   ```

5. **Create a virtual environment:**
   ```
   python3 -m venv .venv
   ```

6. **Activate it:**
   ```
   source .venv/bin/activate
   ```

7. **Install the tool:**
   ```
   pip install -e .
   ```
   This installs the tool so you can use it from any folder on your computer (not just the outlook-tool folder).

8. **Verify it works:**
   ```
   outlook-tool search --from 2026-03-01 --to-date 2026-03-26 --max-results 5
   ```

9. **If macOS asks for permission** to control Outlook, click **Allow**. This is a one-time prompt.

---

## Daily Usage

Every time you open a new terminal to use the tool, you need to activate the virtual environment first. You do **not** need to be in the outlook-tool folder — activate from wherever you are:

**Windows:**
```
C:\Users\YourName\Projects\outlook-tool\.venv\Scripts\activate
```

**Mac:**
```
source ~/Projects/outlook-tool/.venv/bin/activate
```

You'll know it's active when you see `(.venv)` at the start of your terminal line. Once active, `outlook-tool` and `from outlook_tool import ...` work from any folder.

---

## Command Line Examples

### Search Emails

```bash
# Find emails from the last week
outlook-tool search --from 2026-03-19 --to-date 2026-03-26

# Find emails from a specific person
outlook-tool search --sender-name "Jane Smith"
outlook-tool search --sender-email jane.smith@example.com

# Find emails from a specific company
outlook-tool search --domain example.com

# Find emails with a keyword in the subject
outlook-tool search --subject "quarterly report"

# Find only unread emails
outlook-tool search --unread

# Find emails with attachments
outlook-tool search --has-attachments

# Search a specific folder (default is Inbox)
outlook-tool search --from 2026-03-01 --folder "Sent Items"
outlook-tool search --from 2026-03-01 --folder Archive

# Search multiple folders at once
outlook-tool search --from 2026-03-15 --to-date 2026-03-21 --folders Inbox Archive Snoozed

# Combine multiple filters (finds emails matching ALL conditions)
outlook-tool search --from 2026-03-01 --domain example.com --subject "report" --has-attachments

# Limit how many results you get (default is 50)
outlook-tool search --from 2026-01-01 --max-results 10

# Get results as raw JSON (useful for piping to other tools)
outlook-tool search --from 2026-03-01 --json
```

### Download Attachments

```bash
# Search and download all attachments to a folder
outlook-tool search --sender-email boss@example.com --has-attachments --download ./my-downloads
```

This creates a `my-downloads` folder and saves all attachments from matching emails into it.

### Calendar Events

```bash
# Show today's events
outlook-tool events --today

# Show this week's events
outlook-tool events --week

# Events in a specific date range
outlook-tool events --from 2026-03-24 --to-date 2026-03-28

# Filter events by subject
outlook-tool events --from 2026-03-01 --subject "standup"

# Get events as JSON
outlook-tool events --from 2026-03-27 --json
```

### Send Emails

```bash
# Send a simple email
outlook-tool send --to colleague@example.com --subject "Hello" --body "Just checking in."

# Send to multiple people
outlook-tool send --to user1@example.com user2@example.com --subject "Team update" --body "See below."

# Send with CC
outlook-tool send --to user@example.com --cc manager@example.com --subject "FYI" --body "Sharing this."

# Send with file attachments
outlook-tool send --to user@example.com --subject "Report" --body "Attached." --attach report.pdf data.xlsx

# Send a high-importance email
outlook-tool send --to user@example.com --subject "Urgent" --body "Please review ASAP" --importance high
```

---

## Python Script Examples

If you want to use this in your own Python scripts:

```python
from outlook_tool import OutlookClient

client = OutlookClient()

# Search emails
emails = client.search(
    date_from="2026-03-01",
    date_to="2026-03-15",
    sender_domain="example.com",
    subject_contains="report",
    has_attachments=True,
)

# Print what you found
for email in emails:
    print(f"{email['received_date']} — {email['sender_name']}: {email['subject']}")

# Download attachments
for email in emails:
    for att in email["attachments"]:
        path = client.download_attachment(email, att, output_dir="./downloads")
        print(f"Saved: {path}")

# Get calendar events
events = client.get_events(
    date_from="2026-03-27",
    date_to="2026-03-28",
)
for evt in events:
    print(f"{evt['start_date']} {evt['start_datetime'].strftime('%H:%M')} — {evt['subject']}")

# Search across multiple folders
emails = client.search(
    date_from="2026-03-15",
    date_to="2026-03-21",
    folders=["Inbox", "Archive", "Snoozed"],
)

# Send an email
client.send(
    to="colleague@example.com",
    subject="Meeting notes",
    body="Here are the notes from today.",
)

# Send with attachments
client.send(
    to=["team@example.com", "manager@example.com"],
    subject="Q1 Report",
    body="Please review the attached report.",
    attachments=["report.pdf"],
    cc="stakeholder@example.com",
    importance="high",
)
```

---

## Using From Other Projects / Folders

You do **not** need to be inside the `outlook-tool` folder to use this tool. After running `pip install -e .` during setup, the tool is available from anywhere — as long as you've activated the virtual environment.

### From the command line (any folder)

```bash
# First, activate the venv (one-time per terminal session)
# Windows:
C:\Users\YourName\Projects\outlook-tool\.venv\Scripts\activate

# Mac:
source ~/Projects/outlook-tool/.venv/bin/activate

# Now use it from wherever you are
cd C:\Users\YourName\Projects\some-other-project
outlook-tool search --from 2026-03-01 --subject "report"
```

### From Python scripts (any folder)

```python
# This works in any .py file, as long as the outlook-tool venv is active
from outlook_tool import OutlookClient

client = OutlookClient()
emails = client.search(sender_domain="example.com")
```

### Making it truly global (no venv activation needed)

If you want `outlook-tool` available without activating any virtual environment:

```bash
# Windows (run from any terminal):
pip install -e C:\Users\YourName\Projects\outlook-tool

# Mac:
pip install -e ~/Projects/outlook-tool
```

This installs into your system Python so it's always available. The trade-off is that it's harder to manage updates — the virtual environment approach is recommended for most users.

### Summary

| Install method | Need to be in outlook-tool folder? | Need to activate venv? |
|---|---|---|
| `pip install -e .` (recommended) | No | Yes, once per terminal session |
| `pip install -e /path/to/outlook-tool` | No | No |
| `pip install -r requirements.txt` only | Yes (must use `python cli.py`) | Yes |

---

## All Search Filters

Every filter is optional. When you use multiple filters, only emails matching **all** of them are returned.

| What you want to filter by | CLI flag | Example |
|---|---|---|
| Start date | `--from` | `--from 2026-03-01` |
| End date | `--to-date` | `--to-date 2026-03-15` |
| Subject keyword | `--subject` | `--subject "quarterly report"` |
| Subject pattern (regex) | `--subject-regex` | `--subject-regex "Q[1-4] 2026"` |
| Sender's display name | `--sender-name` | `--sender-name "Jane"` |
| Sender's exact email | `--sender-email` | `--sender-email jane@example.com` |
| Sender's company domain | `--domain` | `--domain example.com` |
| Multiple domains | `--domains` | `--domains example.com partner.org` |
| Has attachments | `--has-attachments` | `--has-attachments` |
| Specific folder | `--folder` | `--folder "Sent Items"` |
| Multiple folders | `--folders` | `--folders Inbox Archive Snoozed` |
| Only unread | `--unread` | `--unread` |
| Only read | `--read` | `--read` |
| Body text keyword | `--body` | `--body "action required"` |
| Recipient contains | `--to` | `--to "team@example.com"` |
| Importance level | `--importance` | `--importance high` |
| Limit results | `--max-results` | `--max-results 20` |
| Download attachments | `--download` | `--download ./my-folder` |
| Output as JSON | `--json` | `--json` |

---

## Troubleshooting

### "No email backend available"
- **Windows:** Make sure you ran `pip install -r requirements.txt` with the virtual environment activated
- **Mac:** Make sure `osascript` is available (it should be by default on macOS)

### "Cannot connect to Outlook for Mac"
- Make sure Outlook is **open and running**
- Make sure you're in **Legacy/Classic mode** (not "New Outlook")
- If macOS asked for automation permissions and you clicked Deny, go to **System Settings > Privacy & Security > Automation** and allow Terminal to control Microsoft Outlook

### "python is not recognized" (Windows)
- You may need to use `python3` instead of `python`
- Or re-install Python and make sure to check "Add Python to PATH"

### Emails not showing up in results
- Double-check your date range — dates are inclusive, format is `YYYY-MM-DD`
- Make sure Outlook has the emails downloaded/synced (not still loading)
- Try a broader search first (just `--from` with no other filters) to confirm things work
- By default, only the **Inbox** folder is searched. If your emails are in Archive, Snoozed, or other folders, use `--folders Inbox Archive Snoozed` to search multiple folders at once

### Virtual environment issues
- If `pip install` gives errors, make sure you see `(.venv)` in your terminal prompt
- If not, re-run the activate command for your platform

---

## Updating the Tool

To get the latest version:

```bash
cd outlook-tool
git pull
```

If dependencies changed, re-install:
```bash
# Activate your virtual environment first, then:
pip install -r requirements.txt
```
