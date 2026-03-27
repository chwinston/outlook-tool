# Sprint Log

Archive of completed bolts.

---

## Bolt 1 — Security Hardening & Critical Bug Fixes
**Date:** 2026-03-27 | **Status:** COMPLETE
**Items:** 11 (all done) | **Tests:** 43 (was 34, +9 new)
**Summary:** Fixed all Critical/High findings from staff panel review (AppleScript injection, import scoping, token permissions), added CI pipeline, expanded test coverage.
**Tag:** v1.0.1 → v1.1.0

---

## Bolt 1.5 — Multi-Folder Email Search
**Date:** 2026-03-27 | **Status:** COMPLETE
**Items:** 7 | **Tests:** 43 (unchanged)
**Summary:** Added --folders multi-select for AppleScript/JXA backend. Fixed save_attachment folder-awareness bug. All 3 backends support multi-folder search.
**Tag:** v1.1.0

---

## Bolt 2 — DLC Hardening + Calendar Support
**Date:** 2026-03-27 | **Status:** COMPLETE
**Items:** 12 (Phase 1: 4 DLC fixes, Phase 2: 8 calendar items)
**Tests:** 50 (was 43, +7 calendar tests)
**Summary:** Closed DLC audit gaps (SECURITY.md, REQUIREMENTS.md, context file). Added calendar event search across all 3 backends with CLI command and Python API.
**Captain's log:** docs/captains_log/20260327-2015-bolt2-dlc-calendar.md
**Tag:** v1.2.0
