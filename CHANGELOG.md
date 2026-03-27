# Changelog

All notable changes to this project are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

*No unreleased changes.*

## [1.3.0] - 2026-03-27

### Security
- Fix path traversal in attachment downloads (sanitize filenames)
- Atomic token cache writes (no permission race condition)
- OData injection prevention in Graph API subject filter
- URL-encode Graph API message/attachment IDs in download URLs
- Null byte stripping in AppleScript sanitizer
- Integer validation for AppleScript message/attachment indices

### Changed
- `--json` now outputs clean JSON only (progress output moved to stderr)
- Multi-folder search logic deduplicated into shared helper
- `subject_contains` no longer double-applied on Graph backend
- `max_results` applied after post-filters, not before
- Graph API calls now have 30s timeout, 429 backoff, and 3-attempt retry
- AppleScript `send()` warns when `importance` is ignored
- Event IDs use stable `hashlib.md5` instead of non-deterministic `hash()`
- Win32 COM cache cleared per search to prevent memory leak

### Changed (Dependencies)
- `msal` and `requests` moved to optional `[graph]` extras in `pyproject.toml`

## [1.2.0] - 2026-03-27

### Added
- Calendar event search: `outlook-tool events --from DATE --to DATE`
- Convenience flags: `--today`, `--week` (mutually exclusive)
- Subject filter for events: `--subject "standup"`
- Python API: `client.get_events(date_from, date_to, subject_contains)`
- All 3 backends supported (AppleScript/JXA, Win32 COM, Graph API)
- `SECURITY.md` with trust model, credential handling, vulnerability reporting
- `docs/REQUIREMENTS.md` with MoSCoW-prioritized requirements (REQ-001–REQ-010)
- First captain's log and sprint archive

### Fixed
- CLI `--today`/`--week` made mutually exclusive via argparse
- Win32 calendar `status_map` variable shadowing resolved
- SECURITY.md updated to include calendar event access

## [1.1.0] - 2026-03-27

### Added
- Multi-folder email search: `--folders Inbox Archive Snoozed`
- Single folder search on Mac: `--folder Archive` (was silently ignored)
- Folder-aware attachment downloads (fixes wrong-file bug)
- Per-folder progress output during scans

### Security
- AppleScript/JXA injection prevention (`_escape_applescript` with tests)
- Token cache file permissions set to `0600`
- Graph API pagination bounded to prevent infinite loops
- `_requests` import scoping fixed (prevented crash on Mac)
- JXA scan loop `break` assumption fixed
- Graph API folder names URL-encoded

### Added (Infrastructure)
- GitHub Actions CI pipeline (ruff lint + pytest on Python 3.10/3.12)

## [1.0.0] - 2026-03-27

### Added
- Initial release: cross-platform Outlook email client
- Email search with 15+ filters (date, sender, subject, folder, attachments, etc.)
- Attachment download to specified directory
- Email sending with CC/BCC, attachments, HTML body, importance
- Three backends: Win32 COM (Windows), AppleScript/JXA (Mac), Graph API (fallback)
- CLI tool (`outlook-tool`) and Python API (`OutlookClient`)
- JSON output mode for scripting
- 34 unit tests
