# Multi-Folder Email Search — Staff Engineer Panel Analysis

**Date:** 2026-03-27
**Panel:** Tim (SpaceX), Rob (Roblox), Fran (Meta), Al (AWS), Will Larson (Moderator)
**Trigger:** AppleScript/JXA backend only searches Inbox folder, missing emails in Archive, Snoozed, and other folders

---

## Problem Statement

The AppleScript/JXA backend hardcodes `outlook.inbox` in the JXA scan script (line 169). On Mac, emails are distributed across many folders (Archive: 1,088, Snoozed: 45, Sent Items: 5,491). A date-range search that should return 51+ emails returns only 4.

Additionally, `save_attachment()` (line 323) hardcodes `message {index} of inbox`, creating a latent data integrity bug if folder search is added without fixing this.

## Decisions

```
DECISION: Add --folders multi-select to AppleScript backend | VOTE: 4-0 | CONFIDENCE: 5.0 | DISSENT: NONE
DECISION: Fix save_attachment folder-awareness | VOTE: 4-0 | CONFIDENCE: 4.8 | DISSENT: NONE
DECISION: Defer list-folders command | VOTE: 2-2 | CONFIDENCE: 3.3 | DISSENT: Al: useful for discovery
```

## Implementation Plan

1. Add `folder_name` parameter to `_AppleScriptBackend.scan_emails()` — resolve via JXA Exchange account
2. Modify JXA script to use resolved folder reference instead of `outlook.inbox`
3. Fix `save_attachment()` to accept and use folder name
4. Update `_search_applescript()` to accept `folders` list, loop and merge results
5. Add `folders` parameter to `OutlookClient.search()`, merge with `folder` like domains pattern
6. Add `--folders` CLI flag with `nargs="+"`
7. Update README with new features (preserve all existing content)

## Files Referenced

| File | Role |
|------|------|
| outlook_tool.py:169 | JXA script hardcodes `outlook.inbox` |
| outlook_tool.py:323 | `save_attachment` hardcodes inbox |
| outlook_tool.py:803 | `_search_applescript` doesn't forward folder |
| outlook_tool.py:647-651 | Domain merge pattern to follow |
| cli.py:140 | Existing `--folder` flag |
