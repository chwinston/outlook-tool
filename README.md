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

7. **Install dependencies:**
   ```
   pip install -r requirements.txt
   ```

8. **Verify it works:**
   ```
   python cli.py search --from 2026-03-01 --to-date 2026-03-26 --max-results 5
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

7. **Install dependencies:**
   ```
   pip install -r requirements.txt
   ```

8. **Verify it works:**
   ```
   python3 cli.py search --from 2026-03-01 --to-date 2026-03-26 --max-results 5
   ```

9. **If macOS asks for permission** to control Outlook, click **Allow**. This is a one-time prompt.

---

## Daily Usage

Every time you open a new terminal to use the tool, you need to activate the virtual environment first:

**Windows:**
```
cd C:\Users\YourName\Projects\outlook-tool
.venv\Scripts\activate
```

**Mac:**
```
cd ~/Projects/outlook-tool
source .venv/bin/activate
```

You'll know it's active when you see `(.venv)` at the start of your terminal line.

---

## Command Line Examples

### Search Emails

```bash
# Find emails from the last week
python cli.py search --from 2026-03-19 --to-date 2026-03-26

# Find emails from a specific person
python cli.py search --sender-name "Jane Smith"
python cli.py search --sender-email jane.smith@example.com

# Find emails from a specific company
python cli.py search --domain example.com

# Find emails with a keyword in the subject
python cli.py search --subject "quarterly report"

# Find only unread emails
python cli.py search --unread

# Find emails with attachments
python cli.py search --has-attachments

# Combine multiple filters (finds emails matching ALL conditions)
python cli.py search --from 2026-03-01 --domain example.com --subject "report" --has-attachments

# Limit how many results you get (default is 50)
python cli.py search --from 2026-01-01 --max-results 10

# Get results as raw JSON (useful for piping to other tools)
python cli.py search --from 2026-03-01 --json
```

### Download Attachments

```bash
# Search and download all attachments to a folder
python cli.py search --sender-email boss@example.com --has-attachments --download ./my-downloads
```

This creates a `my-downloads` folder and saves all attachments from matching emails into it.

### Send Emails

```bash
# Send a simple email
python cli.py send --to colleague@example.com --subject "Hello" --body "Just checking in."

# Send to multiple people
python cli.py send --to user1@example.com user2@example.com --subject "Team update" --body "See below."

# Send with CC
python cli.py send --to user@example.com --cc manager@example.com --subject "FYI" --body "Sharing this."

# Send with file attachments
python cli.py send --to user@example.com --subject "Report" --body "Attached." --attach report.pdf data.xlsx

# Send a high-importance email
python cli.py send --to user@example.com --subject "Urgent" --body "Please review ASAP" --importance high
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
