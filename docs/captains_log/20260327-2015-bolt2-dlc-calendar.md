# Captain's Log — Bolt 2: DLC Hardening + Calendar Support

**Date:** 2026-03-27
**Branch:** bolt/20260327-dlc-hardening-calendar-support
**Previous log:** None (first captain's log)

## What was done

### Phase 1: DLC Audit Fixes
- Filled all TODO markers in SECURITY.md (contact email, removed placeholder comments)
- Created docs/REQUIREMENTS.md with MoSCoW prioritization (REQ-001 through REQ-010)
- Updated claude-instructions.md with multi-folder search docs and calendar feature
- Staff panel confirmed only 2 of 13 AI-DLC foundation docs are relevant (SECURITY.md + REQUIREMENTS.md)

### Phase 2: Calendar Support
- Added `get_events()` to OutlookClient with all 3 backends
- AppleScript/JXA: queries default calendar via Exchange account, extracts events with attendees
- Win32 COM: uses olFolderCalendar with date restriction and recurrence expansion
- Graph API: uses /me/calendarview endpoint with OData params
- CLI: `outlook-tool events --from DATE --to DATE [--today] [--week] [--subject] [--json]`
- 7 new tests (50 total), all passing
- README updated with calendar examples

## Decisions made

1. **Staff panel (calendar architecture):** `get_events()` (not `search_events()`) — you're fetching by range, not keyword searching. MVP is read-only. Subject filter is post-filtered in Python like email search.
2. **CLI command:** `outlook-tool events` (not `calendar`) — flat subcommand pattern.
3. **Convenience flags:** `--today` and `--week` set date range automatically.
4. **DLC scope:** Only SECURITY.md and REQUIREMENTS.md created — 11 other foundation docs are irrelevant for a small CLI tool (panel vote: 4-0 on each skip).

## Issues encountered

- AppleScript `whose` clause works for counting messages but fails when iterating results — this quirk was discovered during the multi-folder work in Bolt 1.5 and avoided here by using JXA iteration.
- Calendar events via JXA don't have a reliable unique ID — used hash of subject + start time as fallback.

## Lessons learned

- For small projects, AI-DLC should be right-sized. A staff panel is a great way to filter out ceremony that doesn't serve the project.
- The 3-backend architecture (Win32/AppleScript/Graph) extends cleanly to new features — calendar followed the exact same pattern as email.

## Next steps

- Review findings from code review (pending)
- Consider: create meeting invites (`send_meeting()`) — REQ-010, priority: Could
- Consider: contacts search — not yet in requirements
- Tag v1.2.0 after merge
