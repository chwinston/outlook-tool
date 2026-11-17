# Changelog

All notable changes to this project are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

### Added — Bolt 4 Phase 1 (write verbs)

- `OutlookClient.move(message_id, target_folder, source_folder=None)` — move a
  message between folders; verified on round-trip. Detect-and-creates the
  target folder on AppleScript and COM.
- `OutlookClient.delete(message_id, permanent=False, yes_really_permanent=False, folder_name=None)`
  — soft-delete (into Deleted Items) by default. `permanent=True` is
  REFUSED at the method boundary unless `yes_really_permanent=True` is set
  in the same call. The double-flag is deliberate (panel F6, unanimous).
- `OutlookClient.mark_read(message_id, read=True, folder_name=None)` — flip
  the read/unread flag with round-trip verify.
- Unified `OperationResult` dataclass — `{status, operation, message_id,
  verified, before_state, after_state, error_detail, backend, elapsed_ms}`.
  `status ∈ {"success", "failed", "unsupported", "dry_run"}`. `verified=True`
  only when a post-op re-fetch confirms the state change.
- CLI verbs: `outlook-tool move`, `outlook-tool delete`, `outlook-tool mark-read`.
  Every verb is **dry-run by default**; `--apply` is required to touch the
  mailbox. Hard-delete additionally requires `--permanent --yes-really-permanent`.
  Supports `--json` output for scripting.
- Backend coverage: AppleScript (Mac Legacy Outlook) + win32com (Windows) +
  Graph. Graph returns `status="unsupported"` without the `Mail.ReadWrite`
  scope — no auto-upgrade, so the scope cliff is visible to callers
  instead of hiding behind a mid-triage auth prompt. Scope-missing and
  403 permission-denied are kept distinct in the result (panel F2).
- Integration test harness: a disposable `_bolt4_test` folder created on
  the real mailbox via a session-scoped pytest fixture, torn down on exit.
  Integration tests are marked `@pytest.mark.integration` and skipped from
  the default `pytest` run — CI without Outlook stays green; opt-in locally
  with `pytest -m integration`.

### Safety notes

- **Dry-run default, everywhere.** CLI write verbs will not touch the
  mailbox without `--apply`. The dry-run path does not even instantiate an
  `OutlookClient`.
- **Hard-delete is double-flag-gated.** Both the `--permanent --yes-really-permanent`
  pair at the CLI and the `permanent=True, yes_really_permanent=True`
  kwargs at the Python API. Either path alone raises `ValueError` (API) or
  exits 2 (CLI). The pair is per-call on purpose — no `--yes` global
  carries across invocations.
- **Round-trip verify on every write.** `OperationResult.verified=True`
  requires the backend to re-fetch the target message after the op and
  confirm the state change. This defends against the AppleScript
  silent-success failure mode (op returns 0 bytes to stderr but Outlook
  refused the change).
- **COM Move() handle invalidation.** The Windows backend captures the
  return value of `MailItem.Move()` — the original EntryID is invalidated
  after the move, so the result's `message_id` is rewritten to the new
  EntryID so callers can chain operations.

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
