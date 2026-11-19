# Marketplace Scan — Bolt 6 Phase A3

**Date:** 2026-11-17
**Scope:** What exists in the wild for Outlook/email triage via Claude Code (plugins, skills, MCP servers) + patterns to steal from commercial AI mail tools. Prevents re-solving solved problems.
**Author:** A3 researcher (parallel to A1 pattern-rules, A2 LLM-classifier)

---

## One-page summary

- **The official Anthropic plugin marketplace has zero Outlook/email plugins** as of 2026-11-17. The `claude-plugins-official` repo ships 43 first-party plugins (code-review, LSPs, security-guidance, skill-creator, etc.) and 14 external integrations (asana, github, linear, imessage, telegram, discord, etc.) — none touch Outlook, Exchange, Microsoft Graph, Gmail, or mail triage. This is open territory.
- **The community MCP ecosystem has ~10 Outlook-over-Graph servers, but they're mostly abandoned toys** (1-star, one-commit repos). The two credible ones are `XenoXilus/outlook-mcp` (32★, updated 2026-11-03) and `dockndevai/mcp-outlook` (0★ but safety-first design, updated 2026-11-11). None solves the actual hard problem — classifying noise at scale with sender-reputation signals. They're plumbing, not product.
- **Community "inbox triage" skills exist but are weaker than Chris's Bolt 4.** `ericporres/email-triage-plugin` (28★) is Gmail-only and single-commit abandoned. `AgriciDaniel/claude-email` (139★ — the highest-star email-focused skill) is a marketing-strategy grab-bag, not a triage engine. Chris's classifier→executor split with dry-run gate, bucket-level reasoning, and bulk-move primitive is architecturally ahead of everything found.
- **The commercial tools (SaneBox, Shortwave, Superhuman) have 3 patterns worth borrowing**: (a) SaneBox's header-only privacy stance and training-period expectation, (b) Shortwave's **automated bundling** of newsletters into collapsible stacks with a single-key dismiss, (c) alfred_ai's **alias-aware classification** ("mail sent to `shopping@` always = low priority") which gives 80/20 cheap wins before you reach for the LLM.
- **Recommendation:** Do NOT build an Outlook MCP server from scratch — Bolt 6 should treat that as a fallback if the direct CLI hits limits, not a goal. Do steal the Shortwave bundling pattern (bundle-by-sender with batch dismiss) and SaneBox's VIP learning (sender interaction history → VIP tier). Do publish Chris's work as a Claude Code plugin once Bolt 6 is mature — nothing else in this category exists on the official marketplace.

---

## 1. Claude Code plugins (official + community)

### Anthropic official marketplace — `anthropics/claude-plugins-official`

Pulled the full plugin list via `gh api` on 2026-11-17. **Zero** email/Outlook/mail/inbox/triage plugins exist in either `/plugins` or `/external_plugins`:

- **First-party (43 plugins):** agent-sdk-dev, champion-kit, clangd-lsp, claude-code-setup, claude-for-manufacturing, claude-md-management, claude-security, code-modernization, code-review, code-simplifier, commit-commands, csharp-lsp, cwc-makers, example-plugin, explanatory-output-style, feature-dev, frontend-design, gopls-lsp, hookify, jdtls-lsp, kotlin-lsp, learning-output-style, lua-lsp, math-olympiad, math-proof, mcp-server-dev, mcp-tunnels, php-lsp, playground, plugin-dev, pr-review-toolkit, project-artifact, pyright-lsp, ralph-loop, receipts, ruby-lsp, rust-analyzer-lsp, security-guidance, session-report, skill-creator, swift-lsp, theme-examples, typescript-lsp.
- **External/partner (14):** asana, context7, discord, fakechat, firebase, github, gitlab, **imessage**, laravel-boost, linear, playwright, serena, telegram, terraform.

**Relevance:** Chris's outlook-tool occupies an **unclaimed category** on the official marketplace. `imessage` is the closest analog (personal-communication plugin) and sets the design precedent. Worth publishing once Bolt 6 stabilizes.

### Community plugins found

- **`gabrieldenny-del/outlook-extension`** — https://github.com/gabrieldenny-del/outlook-extension
  - **What:** Claude Code plugin + Desktop Extension using **local COM** (no Azure/OAuth) — Windows only. Slash commands `/triage-inbox`, `/daily-digest`, `/draft-reply`, `/clean-inbox`.
  - **Relevance:** **Already-reimplemented** — same architecture as Chris's `--backend outlook` on Mac, but Windows-only via COM. Confirms the no-Graph local-backend approach is valid. Weak signal (3★, last push 2026-04-17, 7 months stale).
  - **Signal:** 3★, pushed 2026-04-17 — **likely dead**.
  - **Notes:** The slash-command vocabulary (`/triage-inbox`, `/clean-inbox`, `/daily-digest`) is worth copying for Chris's plugin surface. Bundled `outlook-assistant` skill enforces "draft before sending, soft-delete by default, confirm batch" — same safety posture as Chris's dry-run gate.

- **`ericporres/email-triage-plugin`** — https://github.com/ericporres/email-triage-plugin
  - **What:** Three-tier classifier (Reply Needed / Review / Noise) with **alias-aware routing** — maps `shopping@yourdomain.com` → low priority automatically.
  - **Relevance:** **Steal the alias-aware pattern.** Chris's tool doesn't currently exploit `+alias` or per-alias sender maps. Cheap win: every Gmail/O365 user with a `+shopping` or `shopping@` alias gets instant 80/20 classification without touching a classifier.
  - **Signal:** 28★, pushed 2026-02-19 — **dead** (single commit, 9 months stale).
  - **Notes:** Gmail-only, no Outlook support. Classifier buckets are coarser than Chris's Bolt 4 (safe-delete / resolved-ticket / low-signal-client / needs-action / keep).

- **`AgriciDaniel/claude-email`** — https://github.com/AgriciDaniel/claude-email
  - **What:** Marketing-strategy grab-bag — importance scoring, composition frameworks, deliverability audits (SPF/DKIM/DMARC), automation sequences, blacklist monitoring. Requires Gmail or O365 MCP server underneath.
  - **Relevance:** **Ignore for triage.** This is a marketer's skill, not an inbox-management skill. 139★ reflects marketing-strategy demand, not triage quality. No VIP/sender-reputation/bulk-action features in docs.
  - **Signal:** 139★, pushed 2026-05-23 — **stale** (6 months).
  - **Notes:** Interesting that 139★ is the ceiling for email-related Claude skills. Low ceiling = category is underserved.

- **Marketplace-listed skills** (mcpmarket.com, getclaudeskills.com, claudskills.com) — all three list "Outlook Inbox Triage" variants from `ilderaj`, `alirezarezvani`, and others. Buckets used: "Urgent / Needs Reply Soon / Waiting / FYI". All **read-only by default**, modifications only on explicit request. **Chris's Bolt 4 bucket vocabulary (safe-delete / resolved-ticket / low-signal-client / needs-action / keep) is more decision-useful** — it maps directly to actions instead of leaving the user to interpret "Waiting" vs "FYI".

---

## 2. Claude skills ecosystem (non-plugin)

### Superpowers (Jesse Vincent / `obra`) — https://github.com/obra/superpowers

- **What:** Skills library Chris already uses (brainstorming, systematic-debugging, writing-plans, TDD, worktrees, parallel-agents). No email or Outlook skill in the catalog.
- **Relevance:** **Already-reimplemented** for workflow. The `inbox-triage` skill Chris built in Bolt 4 is NOT in superpowers — it's Chris's own addition, and nothing comparable exists upstream. Worth considering whether `/inbox-triage` belongs in a public "obra/superpowers-communications" companion repo.
- **Signal:** Very active, 257k stars reported across coverage.

### `ComposioHQ/awesome-claude-skills` — https://github.com/ComposioHQ/awesome-claude-skills

Curated list. Email-adjacent skills found: **Gmail Automation, Outlook Automation, Outlook Calendar Automation, SendGrid Automation, Postmark Automation, Internal Comms, Meeting Insights Analyzer, File Organizer**. All are **thin API wrappers** ("send email, list folders, manage contacts") via Composio's MCP Gateway — not triage engines. **Ignore for Bolt 6**, but the Composio MCP Gateway is worth knowing about as a plumbing option if Chris ever needs Gmail support.

### `alirezarezvani/claude-skills` (410 skills) — https://github.com/alirezarezvani/claude-skills

Scanned the "productivity" directory. Contains an `inbox-triage` skill (listed on getclaudeskills.com) — single-file SKILL.md, no code, no Outlook integration. **Ignore** — vaporware-tier.

### `ai-orchestrators/email-triage` SKILL.md

Interactive-setup skill that generates a per-mailbox triage skill for Gmail/Outlook. Approach: ask user about their mailbox, generate a bespoke rules file. **Steal the "interactive setup wizard" pattern** — Chris could add a `/inbox-triage --setup` that walks a new user through their top-10 senders and builds a custom `inbox_rules.yml`.

---

## 3. MCP servers — Outlook / Graph / Exchange / Gmail

Live GitHub data pulled 2026-11-17 via `gh api`:

| Server | Stars | Last push | Status |
|---|---|---|---|
| `XenoXilus/outlook-mcp` | **32** | 2026-11-03 | **Live** — most-starred, 2 weeks fresh |
| `dockndevai/mcp-outlook` | 0 | 2026-11-11 | **Live** — safety-first design, brand new |
| `kelvinuduma/microsoft-graph-mcp` | 0 | 2026-11-16 | **Live** — 20 tools across M365, 1 day old |
| `gabrieldenny-del/outlook-extension` | 3 | 2026-04-17 | Stale (7mo) |
| `hmmroger/simply-outlook-mcp` | 1 | 2026-04-16 | Stale (7mo) |
| `nsakki55/outlook-mcp` | 1 | 2026-04-12 | **Dead** (7mo, single-day commits) |
| `ampcome-mcps/outlook-mcp` | 1 | 2025-07-03 | **Dead** (16mo) |
| `jibberish69/outlook-email-mcp` | 0 | 2025-03-13 | **Dead** (20mo) |
| `Cam10001110101/mcp-server-outlook-email` | — | — | Not inspected (low signal) |
| `juvantlabs/m365-mail-mcp-server` | — | — | Draft/send gated phase — not inspected |
| `jspv/msgraph-email-calendar-mcp` | — | — | Not inspected |

### Top 3 inspected in depth

- **`XenoXilus/outlook-mcp`** (32★, pushed 2026-11-03)
  - 16 tools: auth, email ops (list/get/search/send/draft/delete/move/mark/flag), folder mgmt (list/create), inbox rules (create/list/reorder).
  - Auth: Auth Code + PKCE OAuth2, no client secret. AES-256-CBC token encryption.
  - **Relevance:** Most credible Graph-based alternative to Chris's `--backend graph` mode. **Ignore for adoption** — Chris already has Graph implemented; swapping to a third-party MCP adds dependency surface without new capability. **But** the `inbox rules create/list/reorder` tool is worth studying — Chris's tool doesn't yet manipulate server-side Graph rules, only does client-side moves. Server-side rules would persist across devices and survive the tool not running.
  - Node.js 22+ required. MIT license.

- **`dockndevai/mcp-outlook`** (0★, pushed 2026-11-11)
  - **Steal this design.** Explicit safety modes: `read-only` → `admin`. Env flags `OUTLOOK_ALLOW_SEND`, `OUTLOOK_ALLOW_DELETE`. Folder allowlist/protected-list. Dry-run mode. Audit log of guarded ops. Attachment bytes never returned (metadata only). Version pinning to prevent concurrent mod conflicts.
  - **Relevance:** This is **the safety posture Chris already converged on** (dry-run gate, classifier/executor split). Validates Bolt 4 architecture. Worth copying the `OUTLOOK_ALLOW_SEND`-style env-flag pattern for Bolt 6's forthcoming write verbs — simpler than per-command confirmation prompts.

- **`kelvinuduma/microsoft-graph-mcp`** (0★, pushed 2026-11-16)
  - 20-tool Graph server spanning Outlook + Calendar + OneDrive + Teams + SharePoint. Device-code auth, encrypted token cache.
  - **Relevance:** Notable for breadth but **absence of folder-move, categorization, mass-delete** means it's still plumbing, not triage. The "drafts not sends, safe folders not the whole disk" default is good hygiene to copy.

### MCP official registry

- `modelcontextprotocol/servers`: **no email/Outlook servers** in reference set — only Everything, Fetch, Filesystem, Git, Memory, Sequential Thinking, Time.
- `registry.modelcontextprotocol.io`: returned 503 on fetch; listed as the canonical index but the SSR didn't complete. Worth a manual browse later.

### Gaps

- **No Exchange on-prem MCP server exists.** Everything is Graph (cloud M365) or local COM (Windows Outlook).
- **No sender-reputation or VIP-learning MCP exists.** Every server is plumbing; none remembers "Chris always replies to Bryant → VIP tier."
- **No bulk-triage MCP exists.** Every server exposes one-email-at-a-time primitives. Chris's `inbox move --query "from:x"` bulk primitive from Bolt 5 is novel in this ecosystem.

---

## 4. Non-Claude commercial tools — patterns to steal

| Tool | $/mo | Approach | Relevance to outlook-tool |
|---|---|---|---|
| **SaneBox** | $7 | Header-only ML, no body reads. Trains 1-2 weeks on your behavior. Routes to `@SaneLater`, `@SaneNews`, `@SaneBlackHole` folders server-side. Daily digest. | **Steal:** VIP learning via interaction history. "If Chris opens and replies to Bryant → bump Bryant to VIP tier." Header-only stance is a privacy selling point. |
| **Superhuman** | $30 | Keyboard-driven. Auto Labels (Important/Notifications/Marketing/Other). Split inboxes. Auto Drafts. | **Steal:** Four-bucket default (Important/Notifications/Marketing/Other) is simpler than Chris's 5-bucket Bolt 4 vocabulary. Auto Drafts for high-signal threads is a future Bolt. |
| **Shortwave** | $20 | AI-native Gmail. **Automated bundling** of newsletters into collapsible stacks with keyboard dismiss (`E`). Thread-level RAG. Multi-step AI actions: "Find all unread from investors this month and summarize." | **Steal #1:** Auto-bundling by sender-class with single-command dismiss. Map to Chris's tool as `inbox bundle --class newsletters --dismiss` — groups by List-Unsubscribe presence, then bulk-archives. **Steal #2:** Multi-step natural-language actions like "summarize all unread from my team this week" (cross between `inbox search` + `inbox summary`). |
| **Boomerang / Spark** | $0-15 | Newsletter auto-categorization by domain. Mass unsubscribe via header-sourced `List-Unsubscribe`. Scheduled send, read receipts, follow-up reminders. | **Steal:** Honor `List-Unsubscribe` header as a first-class triage signal — presence of it is a strong "this is bulk, bucket as newsletter" rule. |
| **alfred_ai** | $20 | Content-relationship analysis — "reply you've been chasing for 2 weeks" ranks above a newsletter titled `URGENT`. Alias-aware classification. | **Steal:** Reply-history as a VIP signal. If Chris has replied to a sender in the last 30 days, that thread outranks bulk mail regardless of urgency words. |

### Patterns Chris's tool is already better at

1. **Classifier/executor split with dry-run gate.** No commercial tool exposes this — they all treat triage as a black box. Chris's tool produces inspectable JSON report + reviewable plan before any write. This is a differentiator worth leaning into.
2. **Bulk-move primitive from CLI.** Bolt 5 `inbox move --query` operates on arbitrary result sets. SaneBox et al. are confined to their server-side folders. Shortwave's bundling is UI-bound. Chris's version composes with pipes.
3. **Mac-native backend without Graph.** None of the commercial players run without cloud OAuth. Chris's AppleScript backend is unique — matters for users on self-hosted Exchange or paranoid compliance orgs.

### Patterns Chris's tool is missing

1. **Sender-reputation learning.** Right now rules are hand-written in `inbox_rules.yml`. SaneBox's 1-2 week training loop is a proven UX — Chris could log "Chris opened + did not delete" events to a ledger and use a sliding window to compute a per-sender VIP score.
2. **List-Unsubscribe header as a triage signal.** RFC 2369 header is present on virtually all legitimate bulk mail. Honoring it as a `low-signal-bulk` rule gives near-perfect newsletter detection with zero LLM calls.
3. **Bundle-then-dismiss UX.** Shortwave groups newsletters into one collapsible item with one keystroke. CLI equivalent: `inbox bundle --from-domain github.com | inbox delete --dry-run` or a dedicated `inbox sweep newsletters`.

---

## Recommendations for Bolt 6

1. **Do not build a new MCP server.** Chris's direct CLI is faster to iterate than a Graph-MCP indirection. If Graph performance becomes a bottleneck, `XenoXilus/outlook-mcp` (32★, fresh) is the fallback — fork it, don't start fresh.
2. **Borrow three patterns from the commercial tools:**
   - List-Unsubscribe header → `low-signal-bulk` bucket (zero LLM, near-perfect newsletter detection).
   - Reply-history ledger → VIP tier (SaneBox-style learning in <100 LOC).
   - Bundle-by-sender → `inbox sweep <domain>` convenience verb.
3. **Borrow the alias-aware pattern from `ericporres/email-triage-plugin`.** Map `+alias` / `alias@domain` to default buckets in `inbox_rules.yml`. 80/20 cheap win.
4. **Study `dockndevai/mcp-outlook`'s env-flag safety model** before building Bolt 6 write verbs. `INBOX_ALLOW_SEND=1`, `INBOX_ALLOW_DELETE=1` is simpler than per-command dry-run arguments and composes well with CI/automation.
5. **Plan to publish Chris's work as an official-marketplace plugin.** Zero competition on the Anthropic marketplace, nearest neighbor is `imessage`. Namespace suggestion: `outlook-triage` or `inbox-triage`. This is a flag-planting opportunity before someone else does it less carefully.

---

**Sources:**
- https://github.com/anthropics/claude-plugins-official (plugin list pulled via `gh api` 2026-11-17)
- https://github.com/XenoXilus/outlook-mcp
- https://github.com/dockndevai/mcp-outlook
- https://github.com/kelvinuduma/microsoft-graph-mcp
- https://github.com/gabrieldenny-del/outlook-extension
- https://github.com/ericporres/email-triage-plugin
- https://github.com/AgriciDaniel/claude-email
- https://github.com/ComposioHQ/awesome-claude-skills
- https://github.com/ai-orchestrators/email-triage
- https://github.com/obra/superpowers
- https://www.aicentralresources.com/blog/superhuman-vs-shortwave-vs-sanebox-ai-email-2026
- https://get-alfred.ai/blog/best-ai-email-triage-tools
- https://sparkmailapp.com/glossary/unsubscribe
- https://mcpmarket.com/tools/skills/outlook-email-inbox-triage
