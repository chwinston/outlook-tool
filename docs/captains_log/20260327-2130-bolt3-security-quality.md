# Captain's Log — Bolt 3: Security Hardening & Code Quality Sweep

**Date:** 2026-03-27
**Branch:** bolt/20260327-security-quality-sweep
**Previous log:** docs/captains_log/20260327-2015-bolt2-dlc-calendar.md

## What was done

Fixed all 15 findings from the combined staff engineer panel review and security audit.

### P0 (3 items — all fixed)
- Path traversal in attachment downloads — filenames now sanitized via Path.name
- --json CLI output now clean (human output suppressed, progress to stderr)
- Token cache writes are now atomic (tempfile + rename, permissions set before content written)

### P1 (6 items — all fixed)
- OData injection: parentheses stripped from subject filter values
- Graph API IDs: URL-encoded in download URL construction
- Multi-folder search: deduplicated into single _search_single_folder helper
- subject_contains: no longer double-applied when Graph already filtered server-side
- max_results: now applied AFTER post-filters, not before
- Graph API: 30s timeout, 429 backoff with Retry-After, 3-attempt retry loop

### P2 (6 items — all fixed)
- AppleScript send warns when importance is ignored
- Event IDs use hashlib.md5 (deterministic) instead of hash() (randomized)
- Win32 COM cache cleared per search to prevent memory leak
- msal/requests moved to optional [graph] extras in pyproject.toml
- msg_index/att_index validated as int before AppleScript interpolation
- Null bytes stripped in _escape_applescript

## Decisions made

1. All progress/diagnostic output moved to stderr across all backends — enables clean `| jq` piping
2. Multi-folder search deduplicated: single `_search_single_folder` closure captures all search params and dispatches to the active backend
3. Graph API subject_contains skipped in post-filter since it's already applied server-side via OData

## Next steps

- Consider extracting JXA scripts to separate .js files (Tim's suggestion, deferred)
- Monitor for any issues from the pyproject.toml dependency change
