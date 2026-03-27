# Current Sprint

## Bolt 1 — Security Hardening & Critical Bug Fixes

**Status:** COMPLETE
**Goal:** Fix all Critical and High findings from the Staff Engineer Panel review (F1-F5), plus Medium findings (F6-F8), dead dependency (F9), CI pipeline, and new tests.
**Opened:** 2026-03-27
**Completed:** 2026-03-27

---

### Items

| # | Item | Size | Status | Finding |
|---|------|------|--------|---------|
| 1 | Fix AppleScript command injection in `send_email()` | M | done | F1 (Critical) |
| 2 | Fix AppleScript injection in `save_attachment()` | S | done | F2 (Critical) |
| 3 | Expand AppleScript escaping (newlines, tabs, special chars) | S | done | F3 (High) |
| 4 | Fix `_requests` import scoping — module crashes on Mac | S | done | F4 (High) |
| 5 | Fix token cache file permissions (chmod 0600) | S | done | F5 (High) |
| 6 | Fix JXA scan loop `break` assumption | S | done | F6 (Medium) |
| 7 | URL-encode Graph API folder name | S | done | F7 (Medium) |
| 8 | Add pagination circuit breaker to Graph API | S | done | F8 (Medium) |
| 9 | Remove dead `openpyxl` from requirements.txt | S | done | F9 (Low) |
| 10 | Add GitHub Actions CI (ruff + pytest) | M | done | Panel recommendation |
| 11 | Add tests for `_escape_applescript` helper | S | done | Test coverage for F1-F3 |

### Metrics

- Commits this bolt: pending (not yet committed)
- Tests: 43 (was 34, +9 new escape tests, all passing)
- Files changed: 3 (outlook_tool.py, requirements.txt, tests/test_outlook_tool.py)
- Files added: 3 (CI workflow, PM docs)
