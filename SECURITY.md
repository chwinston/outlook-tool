# Security — Outlook Tool

## Trust Model


This tool runs **locally on your machine** and communicates directly with your Outlook desktop app. No data is sent to external servers (except when using the optional Graph API fallback, which authenticates via Microsoft's device code flow).

### What the tool can access

- Email subjects, senders, recipients, body text, and metadata
- Email attachments (can save to disk)
- Send emails as the signed-in Outlook user
- Calendar events (subject, times, location, organizer, attendees, body)

### What the tool does NOT do

- Store or cache email content beyond the current session
- Send data to any third-party service
- Retain credentials (Graph API tokens are cached locally at `~/.outlook_tool_token_cache.bin` with `0600` permissions)
- Modify or delete existing emails

## Credential Handling

| Backend | Credential Type | Storage | Scope |
|---------|----------------|---------|-------|
| Win32 COM (Windows) | None — uses running Outlook session | N/A | Whatever the signed-in Outlook user can access |
| AppleScript/JXA (Mac) | None — uses running Outlook session | N/A | Whatever the signed-in Outlook user can access |
| Graph API (fallback) | OAuth2 device code token | `~/.outlook_tool_token_cache.bin` (mode 0600) | `Mail.Read`, `Mail.Send` |


## Attachment Risks

Downloaded attachments are saved to the directory specified by the user. The tool does not scan attachments for malware. Users should exercise the same caution as opening attachments in any email client.

## Input Sanitization

- AppleScript string interpolation is sanitized via `_escape_applescript()` to prevent command injection through crafted email addresses, subjects, or file paths
- JXA folder names are interpolated via `json.dumps()` for safe JavaScript string encoding
- Graph API folder names are URL-encoded to prevent URL injection

## Reporting a Vulnerability

If you discover a security vulnerability, please report it by:

1. Opening a private security advisory at https://github.com/chwinston/outlook-tool/security/advisories
2. Or emailing: cfossenier@membersolutions.com

Please do **not** open a public issue for security vulnerabilities.

## Known Limitations

- The Graph API token cache file stores OAuth tokens. If your machine is compromised, these tokens could be used to read email until they expire.
- AppleScript/JXA backend requires macOS Automation permissions. If granted broadly, other scripts could also control Outlook.
