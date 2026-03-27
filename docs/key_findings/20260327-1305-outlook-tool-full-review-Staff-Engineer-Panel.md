# Outlook Tool — Staff Engineer Panel Analysis + AI-DLC Compliance Audit

**Date:** 2026-03-27
**Panel:** Tim (SpaceX), Rob (Roblox), Fran (Meta), Al (AWS), Will Larson (Moderator)
**Trigger:** Full code review requested by project owner

---

## Problem Statement

**What's happening:** The outlook-tool repo (~1,250 lines across 2 source files) is a cross-platform Outlook email client supporting search, download, and send via three backends (win32com, AppleScript/JXA, Graph API). It's at v1.0.0 with 5 commits and no formal development process.

**Why it matters:** The tool handles email operations — a sensitive surface. It constructs AppleScript/JXA dynamically from user input, sends emails, and manages OAuth tokens. Code quality issues here become security issues.

**Constraints:** Solo developer, small utility repo, no CI/CD, no formal requirements or security review.

---

## Panel Analysis

---

### Tim — Staff Engineer, SpaceX

**Risk assessment:**

The #1 risk in this codebase is **command injection via AppleScript string interpolation**. Every AppleScript command is constructed by string formatting user input directly into script strings. An email with a subject containing `"` or `\` could break the script, and a malicious email address in the `send()` path could inject arbitrary AppleScript commands.

Specific dangerous patterns:

| File | Lines | Risk | P x C |
|------|-------|------|-------|
| `outlook_tool.py` | 301-308 | `save_path` interpolated into AppleScript without sanitization | 0.3 x 9 = **2.7** |
| `outlook_tool.py` | 322-365 | `to`, `cc`, `bcc` addresses interpolated into AppleScript — attacker-controlled if forwarding to user-supplied addresses | 0.5 x 9 = **4.5** |
| `outlook_tool.py` | 340-341 | Subject/body escaping only handles `\` and `"` — doesn't handle newlines, tabs, or AppleScript special chars | 0.4 x 7 = **2.8** |
| `outlook_tool.py` | 475 | `_requests.Response` type annotation crashes the module on Mac because `_requests` is only imported conditionally | 1.0 x 5 = **5.0** |

The `_requests` bug (line 475) is a **confirmed live bug** — the module crashes on import on macOS when Outlook is not in Legacy mode. We already had to fix it during this session.

**Options evaluated:**
- Option A: Sanitize inputs at the `send()` boundary — **FAVORED**. Fix the 4 dangerous interpolation sites. ~2 hours.
- Option B: Rewrite AppleScript backend to use JXA exclusively (JavaScript is easier to safely parameterize) — good but overkill for now.
- Option C: Do nothing — **REJECTED**. Command injection in a tool that sends emails is unacceptable.

**Key quote:**
"You have a code injection vector in a tool that sends emails. That's not technical debt — that's a loaded gun."

**Recommendation:** Fix the 4 injection sites. Add input sanitization helper for AppleScript strings. ~2 hours.

**On the Graph API backend:**
"The `_requests` import scoping is a design smell. Either import unconditionally at the top (behind a try/except), or don't use the name at class-definition scope. The current approach fails on any platform where Graph isn't the selected backend."

**Unique contribution:** Noticed that `save_path` in `save_attachment` (line 305) is interpolated into AppleScript with no escaping at all — a path containing `"` would break the script, and a crafted path could inject commands.

---

### Rob — Staff Engineer, Roblox

**Risk assessment:**

The code is well-structured for its size. The three-backend pattern with a unified dict output is clean. But there are latent bugs hiding in the "it works on my machine" assumptions:

1. **JXA scan loop assumes chronological ordering** (line 170: `if (recvMs < startMs) break;`). If Outlook's inbox isn't sorted newest-first, the loop exits early and silently drops results. This is a **live bug** for anyone with a non-standard sort order.

2. **Win32 COM object caching** (`_win32_msg_cache`, line 959) stores COM references keyed by `id(msg)`. Python's `id()` returns the memory address, which can be reused after garbage collection. If a user does `results = client.search(...)`, then later `client.search(...)` again, old COM references may point to garbage. The error message says "Re-run search()" but doesn't explain why.

3. **`openpyxl` in requirements.txt but not in pyproject.toml** — `requirements.txt` lists `openpyxl>=3.1.0` for "Excel file handling" but it's never imported anywhere in the codebase. Dead dependency.

4. **Body preview truncation inconsistency** — JXA truncates to 5000 chars in JavaScript (line 222), Win32 truncates to 5000 in Python (line 953), Graph returns the full body and truncates in Python (line 1110). The JXA path truncates UTF-16 while the others truncate UTF-8. Multi-byte characters could produce different results.

**Options evaluated:**
- Fix the JXA sort assumption — add explicit sort or remove the early-exit optimization
- Fix the COM cache key — use a stable identifier, not `id(msg)`
- Remove openpyxl from requirements.txt

**Key quote:**
"The JXA `break` on line 170 is a performance optimization that becomes a correctness bug if anyone sorts their inbox differently."

**Recommendation:** Fix the 3 latent bugs. Remove dead dependency. ~1.5 hours.

**On the architecture:**
"Three backends behind one interface is the right call. The post-filter pattern is smart — push what you can to the backend, normalize everything in Python. Don't change the architecture."

**Unique contribution:** Found the `id(msg)` COM cache instability and the openpyxl ghost dependency.

---

### Fran — Staff Engineer, Facebook/Meta

**Risk assessment:**

Bucketing everything into two categories:

**Fix yesterday (dangerous):**
1. AppleScript injection in `send_email()` — email addresses, subject, body all interpolated
2. `_requests` import scoping — crashes on import on non-Graph platforms
3. No CI/CD — zero automated quality gates
4. Token cache file (`.outlook_tool_token_cache.bin`) written to `$HOME` with default permissions — on multi-user systems, other users can read the OAuth token

**Don't care (aesthetic):**
1. Code could be split into multiple files (backend per file) — but at 1,250 lines, single-file is fine
2. No type stubs — acceptable for a utility tool
3. No logging module — print() is fine for CLI tools
4. Some docstrings are long — doesn't matter

**Key quote:**
"Pre-commit is a developer convenience; CI is the contract. You have neither."

**Recommendation:**
1. Add a GitHub Actions workflow: lint (ruff), test (pytest), security (bandit). ~30 minutes.
2. Fix the 4 injection sites. ~1 hour.
3. Set token cache permissions to 0o600. ~5 minutes.

**On testing:**
"34 tests, all passing, good coverage of the pure-Python logic. Zero integration tests, but that's appropriate — you can't CI-test against a running Outlook instance. The mock-based approach is correct here."

**Unique contribution:** Flagged the token cache file permissions issue (line 426: `write_text` uses default umask, no `chmod 0600`).

---

### Al — Staff Engineer, AWS (Microsoft Graph / Identity team perspective)

**Risk assessment:**

From a Graph API and identity perspective:

1. **Hardcoded public client ID** (line 378: `d3590ed6-52b3-4102-aeff-aad2292ab01c`). This is Microsoft Office's well-known client ID. It works for device code flow, but Microsoft can deprecate or restrict it at any time. This is a **brittleness risk**, not a security risk.

2. **Token cache has no encryption** (lines 420-426). The MSAL SerializableTokenCache stores tokens in plaintext JSON. On shared machines, this is a credential exposure vector. MSAL provides `PersistedTokenCache` with OS keychain integration — use it.

3. **Scope escalation is silent** (line 483-492: `upgrade_scopes`). When sending requires `Mail.Send`, the code silently re-authenticates with broader scopes. The user gets a device code prompt but may not understand they're granting additional permissions. Should log a warning.

4. **Graph API pagination** (lines 1064-1069) correctly follows `@odata.nextLink`, but there's no circuit breaker. A malformed API response could cause an infinite loop. Add a max-page limit.

5. **Graph folder lookup** (line 1049) uses `mailFolders('{folder}')` — the folder name is interpolated into the URL path without encoding. Folder names with special characters would break the API call or potentially cause SSRF-like issues against the Graph endpoint.

**Options evaluated:**
- Option A: Switch to OS keychain for token storage — architecturally correct, ~2 hours
- Option B: Just chmod the file — quick fix, ~5 minutes
- Option C: Add encryption wrapper around the cache — over-engineered for this use case

**Key quote:**
"You're storing OAuth tokens in a plaintext file in `$HOME`. That's one `cat` command away from account compromise on any shared machine."

**Recommendation:** Option B (chmod) as immediate fix, Option A as follow-up. Add pagination circuit breaker (max 10 pages). URL-encode folder names. ~3 hours total.

**On the architecture:**
"The three-backend pattern with auto-detection is the right approach. The fallback chain (win32 → AppleScript → Graph) correctly prioritizes local access over network. The Graph backend implementation is solid — proper device code flow, token caching, silent refresh."

**Unique contribution:** Identified the pagination infinite loop risk, Graph folder name injection, and the unencrypted token cache as the highest-impact issues from a platform perspective.

---

## Consensus Matrix

| Question | Tim (SpaceX) | Rob (Roblox) | Fran (Meta) | Al (AWS) |
|----------|-------------|-------------|-------------|----------|
| Fix AppleScript injection? | YES (5) | YES (5) | YES (5) | YES (4) |
| Fix `_requests` import bug? | YES (5) | YES (5) | YES (5) | YES (5) |
| Add CI pipeline? | YES (3) | YES (4) | YES (5) | YES (3) |
| Fix token cache permissions? | YES (4) | YES (3) | YES (5) | YES (5) |
| Fix JXA sort assumption? | YES (4) | YES (5) | NO (2) | YES (3) |
| Fix COM cache key? | NO (2) | YES (5) | NO (2) | NO (2) |
| Remove openpyxl? | YES (4) | YES (5) | YES (4) | YES (3) |
| Add pagination limit? | YES (3) | NO (2) | NO (2) | YES (5) |
| URL-encode Graph folder? | YES (3) | YES (3) | YES (3) | YES (5) |
| Move to OS keychain? | NO (2) | NO (2) | NO (2) | YES (4) |
| Total effort | 3h | 4h | 2h | 5h |

**Unanimous agreements:**
1. Fix AppleScript command injection (all 4, avg confidence 4.8)
2. Fix `_requests` import scoping (all 4, confidence 5.0)
3. Remove dead openpyxl dependency (all 4, avg confidence 4.0)

**Majority agreements (3-of-4):**
4. Add CI pipeline (4-0 but confidence varies)
5. Fix token cache permissions (4-0, avg confidence 4.3)
6. Fix JXA sort assumption (3-1, Fran dissents)
7. URL-encode Graph folder names (4-0, avg confidence 3.5)

**Key disagreements:**
- COM cache key: Only Rob wants to fix it (others say the risk is theoretical)
- Pagination limit: Only Al and Tim want it (others say Graph API is well-behaved)
- OS keychain: Only Al wants it (others say chmod is sufficient)

### Vote Tally

| Decision | For | Against | Confidence (weighted avg) | Result |
|----------|-----|---------|--------------------------|--------|
| Fix AppleScript injection | 4 | 0 | 4.8 | **APPROVED** |
| Fix `_requests` import | 4 | 0 | 5.0 | **APPROVED** |
| Add CI pipeline | 4 | 0 | 3.8 | **APPROVED** |
| Fix token cache perms | 4 | 0 | 4.3 | **APPROVED** |
| Fix JXA sort assumption | 3 | 1 | 3.5 | **APPROVED** |
| Remove openpyxl | 4 | 0 | 4.0 | **APPROVED** |
| Fix COM cache key | 1 | 3 | 2.8 | REJECTED |
| Add pagination limit | 2 | 2 | 3.0 | SPLIT |
| URL-encode Graph folder | 4 | 0 | 3.5 | **APPROVED** |
| Move to OS keychain | 1 | 3 | 2.5 | REJECTED |

### Dissent Record

| Panelist | Decision | Position | Key Concern | Risk if Ignored |
|----------|----------|----------|-------------|-----------------|
| Fran | JXA sort assumption | AGAINST | "Outlook always sorts newest-first in the inbox object" | Theoretical — only affects non-standard setups |
| Rob | COM cache key | FOR | "id() reuse is a real Python pitfall" | Stale COM reference → crash on attachment download |
| Al | OS keychain | FOR | "Plaintext tokens on shared machines" | Token theft on multi-user systems |
| Al | Pagination limit | FOR | "Infinite loops against external APIs" | Hung process consuming memory indefinitely |

---

## Clarifying Questions & Answers

| Question | Answer | Impact on Decision |
|----------|--------|--------------------|
| Is there a CI pipeline? | No. Zero CI/CD. | Confirms Fran's push for GitHub Actions |
| How often is the code modified? | 5 commits total, all within initial development | Low churn — fixes can be batched |
| Is the token cache actually used? | Only when Graph backend is selected (Mac/Linux fallback) | Narrows the blast radius of the token issue |
| Does JXA guarantee sort order? | Outlook for Mac returns messages in reverse chronological by default, but this is not documented as guaranteed | Supports Rob's concern, but Tim notes it's never been observed to fail |
| Are there any `.env` files committed? | No — `.env.example` exists with empty values, `.gitignore` excludes `.env` | Good hygiene |

---

## Will Larson's Decision

**Scope:** Tim's risk-first approach, with Fran's pragmatic bucketing and Al's platform hardening.

| Step | What | Why | Effort |
|------|------|-----|--------|
| 1 | **Fix `_requests` import scoping** — move the type annotation to a string or restructure the import | Live bug, crashes on Mac. Already partially fixed in this session but the root cause (conditional import used at class-definition scope) needs a proper fix | 15m |
| 2 | **Fix AppleScript command injection** — create a `_escape_applescript(s)` helper that escapes `\`, `"`, newlines, and other special chars. Apply to all 4 interpolation sites in `send_email()` and `save_attachment()` | Highest-severity finding. P x C = 4.5 on the send path | 1h |
| 3 | **Fix token cache permissions** — add `os.chmod(path, 0o600)` after writing the token cache | Quick win, reduces credential exposure | 5m |
| 4 | **Remove openpyxl from requirements.txt** — dead dependency | Reduces attack surface, removes confusion | 2m |
| 5 | **URL-encode folder name in Graph API path** | Prevents breakage and potential SSRF with special-char folder names | 10m |
| 6 | **Add pagination circuit breaker** (Will breaks the tie) — max 10 pages in Graph pagination loop | Defense in depth against infinite loops. Al's concern is valid | 10m |
| 7 | **Fix JXA sort assumption** — remove the `break` on early date, replace with `continue` so the loop scans all messages up to the 5000 limit | Correctness over performance. The 5000 hard limit already bounds the scan | 10m |
| 8 | **Add GitHub Actions CI** — ruff lint + pytest on push | Zero quality gates is unacceptable even for a small repo | 20m |

**Total: ~2.5 hours. Fix the injection, fix the bugs, add a CI gate.**

```
DECISION: Fix AppleScript command injection | VOTE: 4-0 | CONFIDENCE: 4.8 | DISSENT: NONE
DECISION: Fix _requests import scoping | VOTE: 4-0 | CONFIDENCE: 5.0 | DISSENT: NONE
DECISION: Add GitHub Actions CI | VOTE: 4-0 | CONFIDENCE: 3.8 | DISSENT: NONE
DECISION: Fix token cache permissions | VOTE: 4-0 | CONFIDENCE: 4.3 | DISSENT: NONE
DECISION: Fix JXA sort assumption | VOTE: 3-1 | CONFIDENCE: 3.5 | DISSENT: Fran: Outlook always sorts newest-first
DECISION: Remove openpyxl dependency | VOTE: 4-0 | CONFIDENCE: 4.0 | DISSENT: NONE
DECISION: Add pagination circuit breaker | VOTE: 2-2 | CONFIDENCE: 3.0 | DISSENT: Rob/Fran: Graph API is well-behaved (Will breaks tie: FOR)
DECISION: URL-encode Graph folder names | VOTE: 4-0 | CONFIDENCE: 3.5 | DISSENT: NONE
DECISION: Fix COM cache key | VOTE: 1-3 | CONFIDENCE: 2.8 | DISSENT: NONE (REJECTED — Rob's concern noted for future)
DECISION: Move to OS keychain | VOTE: 1-3 | CONFIDENCE: 2.5 | DISSENT: NONE (REJECTED — chmod sufficient for now)
```

### What's explicitly deferred

| Item | Rationale | Revisit When |
|------|-----------|--------------|
| COM cache key fix | Risk is theoretical; `id()` reuse requires GC between searches | If users report "COM reference expired" errors |
| OS keychain for tokens | chmod 0600 is sufficient for single-user machines | If tool is deployed on shared infrastructure |
| Split into multi-file package | 1,250 lines is manageable as a single module | If codebase exceeds ~2,000 lines |
| Integration tests | Can't CI-test against Outlook; mock tests are appropriate | If a Graph API test sandbox becomes available |

### Key takeaways

> "You have a code injection vector in a tool that sends emails. That's not technical debt — that's a loaded gun." — Tim

> "The JXA `break` is a performance optimization that becomes a correctness bug if anyone sorts their inbox differently." — Rob

> "Pre-commit is a developer convenience; CI is the contract. You have neither." — Fran

> "You're storing OAuth tokens in a plaintext file in $HOME. That's one `cat` command away from account compromise." — Al

---

## AI-DLC Compliance Report

### Foundation Status

| # | Document | Status | Path |
|---|----------|--------|------|
| 1 | Requirements | MISSING | — |
| 2 | Traceability Matrix | MISSING | — |
| 3 | User Stories | MISSING | — |
| 4 | AI Context File | EXISTS | `claude-instructions.md` |
| 5 | Security Controls | MISSING | — |
| 6 | PM Framework | MISSING | — |
| 7 | Solo+AI Workflow Guide | MISSING | — |
| 8 | CI/CD Deployment Proposal | MISSING | — |
| 9 | Multi-Developer Guide | MISSING | — |
| 10 | Infrastructure Playbook | MISSING | — |
| 11 | Cost Management Guide | MISSING | — |
| 12 | Security Review Protocol | MISSING | — |
| 13 | Ops Readiness Checklist | MISSING | — |
| 14 | AI-DLC Case Study | MISSING | — |

**1/14 foundational documents present.**

### Process Adherence (9 Dimensions)

| # | Dimension | Score | Grade | Details |
|---|-----------|-------|-------|---------|
| 1 | Foundation & Context | 4/10 | D | Context file exists and covers architecture, conventions, setup; missing governance model, no pre-commit hooks, no CI pipeline, 13/14 foundation docs missing |
| 2 | Requirements & Architecture | 2/10 | F | No formal requirements doc, no ADRs, no threat model; architecture described informally in claude-instructions.md only |
| 3 | Specification & Elaboration | 1/10 | F | No user stories, no traceability matrix, no Five Questions usage, no specs beyond README |
| 4 | Construction Process | 3/10 | D | 5 structured commits with clear messages; 34 tests exist and pass; no bolt discipline, no captain's logs, no sprint tracking, no AI-generated code tagging |
| 5 | Security Posture | 1/10 | F | No security review ever conducted; command injection vulnerabilities found by panel; plaintext token storage; no finding tracking |
| 6 | Operational Readiness | 1/10 | F | No CI/CD, no monitoring, no deployment automation, no runbooks, no health checks |
| 7 | Cost Management | 0/10 | F | No cost awareness, no budget, N/A for a local tool but Graph API usage is unbounded |
| 8 | Evolution & Learning | 2/10 | F | Context file created once, updated once (renamed); no retrospectives, no drift detection, no learning artifacts |
| 9 | Human-AI Collaboration | 3/10 | D | claude-instructions.md shows some human curation of AI context; no evidence of human decision gates, no review records, no approval workflow |

**Overall Score: 1.9 / 10**
**Maturity Rating: Foundational (F)**

### Action Items

**Critical (score 0-2 — fix immediately):**
- **D2 Requirements:** Create `docs/REQUIREMENTS.md` with REQ IDs for core features (search, send, download) and security requirements (input sanitization, token handling) -> Run `/dlc-audit init`
- **D3 Specification:** Create `docs/USER-STORIES.md` and `docs/TRACEABILITY-MATRIX.md` -> Run `/dlc-audit init`
- **D5 Security:** Fix AppleScript injection, fix token cache permissions, conduct first security review -> Run `/security-audit`
- **D6 Ops Readiness:** Add GitHub Actions CI (lint + test) -> Manual setup
- **D7 Cost Management:** Document that Graph API has no pagination limit (now being fixed) -> Run `/budget init`

**High (score 3-4 — address this sprint):**
- **D1 Foundation:** Add pre-commit hooks, create remaining foundation docs -> Run `/dlc-audit init`
- **D4 Construction:** Start captain's log practice, tag AI contributions -> Run `/captainslog new`
- **D8 Evolution:** Set up quarterly context file review -> Run `/motherhen`
- **D9 Human-AI Collaboration:** Document decision rationale in captain's logs -> Run `/captainslog new`

### Improvement Plan

| Dimension | Current | Target | Actions | Skill | Timeline |
|-----------|---------|--------|---------|-------|----------|
| D5 Security | 1/10 | 5/10 | Fix injection bugs, create SECURITY.md, run first review | `/security-audit` | 1 week |
| D6 Ops Readiness | 1/10 | 5/10 | Add GitHub Actions CI, add basic health check | Manual | 1 week |
| D1 Foundation | 4/10 | 6/10 | Bootstrap remaining docs, add pre-commit | `/dlc-audit init` | 2 weeks |
| D2 Requirements | 2/10 | 5/10 | Write REQUIREMENTS.md with REQ IDs | `/arch-audit` | 2 weeks |
| D4 Construction | 3/10 | 5/10 | Start captain's logs, tag AI code | `/captainslog new` | Ongoing |

---

## Findings to Fix

| # | Severity | File | Lines | Description | Fix |
|---|----------|------|-------|-------------|-----|
| F1 | **CRITICAL** | `outlook_tool.py` | 322-365 | AppleScript command injection via email addresses in `send_email()` | Create `_escape_applescript()` helper; apply to all string interpolations |
| F2 | **CRITICAL** | `outlook_tool.py` | 301-308 | AppleScript injection via `save_path` in `save_attachment()` | Escape or validate the path before interpolation |
| F3 | **HIGH** | `outlook_tool.py` | 340-341 | Incomplete escaping — only `\` and `"` handled, missing newlines/tabs/special chars | Expand `_escape_applescript()` to cover all special characters |
| F4 | **HIGH** | `outlook_tool.py` | 70-76, 381-475 | `_requests` import conditional but used at class-definition scope | Move import to top-level with try/except, or use string annotation everywhere |
| F5 | **HIGH** | `outlook_tool.py` | 420-426 | Token cache written with default umask (world-readable on some systems) | Add `os.chmod(path, 0o600)` after `write_text()` |
| F6 | **MEDIUM** | `outlook_tool.py` | 170 | JXA scan loop `break` assumes reverse-chronological order | Replace `break` with `continue` |
| F7 | **MEDIUM** | `outlook_tool.py` | 1049 | Graph API folder name not URL-encoded | Use `urllib.parse.quote()` on folder name |
| F8 | **MEDIUM** | `outlook_tool.py` | 1064-1069 | No pagination circuit breaker in Graph API fetch loop | Add `max_pages = 10` counter |
| F9 | **LOW** | `requirements.txt` | 2 | `openpyxl` listed but never imported — dead dependency | Remove the line |
| F10 | **LOW** | `outlook_tool.py` | 940 | COM cache uses `id(msg)` which can be reused after GC | Deferred — use stable key if users report issues |

---

## Files Referenced

| File | Role | Lines |
|------|------|-------|
| `outlook_tool.py` | Core library — all 3 backends, client class, helpers | 1,249 |
| `cli.py` | CLI interface wrapping OutlookClient | 169 |
| `tests/test_outlook_tool.py` | Unit tests (34 tests, all passing) | 307 |
| `pyproject.toml` | Package metadata and build config | 21 |
| `requirements.txt` | Dependencies (includes dead openpyxl) | 6 |
| `claude-instructions.md` | AI context file | 215 |
| `README.md` | User-facing documentation | 363 |
| `.env.example` | Environment variable template | 11 |
