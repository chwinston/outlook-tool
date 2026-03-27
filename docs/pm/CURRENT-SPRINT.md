# Current Sprint

## Bolt 2 — DLC Hardening + Calendar Support

**Status:** COMPLETE
**Goal:** Close DLC audit gaps (D2, D5, D8 improvements), then add calendar event search to all 3 backends with CLI and Python API support.
**Opened:** 2026-03-27
**Branch:** bolt/20260327-dlc-hardening-calendar-support

---

### Phase 1: DLC Audit Fixes (D2, D5, D8)

| # | Item | Size | Status | Dimension |
|---|------|------|--------|-----------|
| 1 | Fill SECURITY.md TODOs (contact email, Graph API scope verification) | S | todo | D5 Security |
| 2 | Update claude-instructions.md with multi-folder search + calendar features | S | todo | D8 Evolution |
| 3 | Prioritize requirements in docs/REQUIREMENTS.md | S | todo | D2 Requirements |
| 4 | Run /security-audit for formal dated findings | M | todo | D5 Security |

### Phase 2: Calendar Support (via /staff-panel)

| # | Item | Size | Status | Notes |
|---|------|------|--------|-------|
| 5 | Staff panel: calendar architecture decisions | — | todo | Design before code |
| 6 | Add get_events() to AppleScript/JXA backend | M | todo | Primary Mac backend |
| 7 | Add get_events() to Win32 COM backend | M | todo | Windows backend |
| 8 | Add get_events() to Graph API backend | M | todo | Fallback backend |
| 9 | Add get_events() to OutlookClient (unified API) | M | todo | Delegates to backends |
| 10 | Add `events` CLI command | S | todo | outlook-tool events --from ... |
| 11 | Update README with calendar examples | S | todo | Preserve existing content |
| 12 | Add calendar tests | M | todo | Mock-based, no Outlook needed |

### Success Criteria

- [ ] SECURITY.md has no TODO markers
- [ ] claude-instructions.md reflects current feature set
- [ ] Requirements prioritized with MoSCoW
- [ ] Security audit completed with dated findings
- [ ] `outlook-tool events --from 2026-03-27 --to-date 2026-03-28` returns calendar events
- [ ] `client.get_events(date_from="2026-03-27")` works in Python
- [ ] All 3 backends support calendar search
- [ ] Tests pass (existing + new calendar tests)
