"""Inbox triage rule base — Bolt 4 Phase 2 (B4.2.3).

Each rule is a callable that takes a message dict (shape from
:meth:`outlook_tool.OutlookClient.search`) and returns ``(bucket, rule_name)``
if it matches, or ``None`` if it doesn't.

Buckets
-------
``safe-delete``       — automated noise with no ongoing signal. Soft-delete.
``resolved-ticket``   — platform ticket activity that is archivable, not
                        immediately destroyable. Move to Archive-style folder.
``low-signal-client`` — routine client FYI mail with no ask. Archive.
``needs-action``      — contains a question, direct ask, or action verb aimed
                        at the operator. Must NOT be auto-deleted.
``keep``              — default. Everything that isn't clearly one of the
                        above.

Design notes
------------
* Rules are **ordered** — the classifier applies them top-to-bottom, and the
  first match wins. This matters: ``needs-action`` precedes ``safe-delete``
  so that a "Build failed" message (action required) doesn't get swept up
  by a broader "Azure DevOps = safe-delete" rule.
* Rules are deliberately simple regex/substring matches. No ML, no
  embeddings. When precision drops below 90% on a bucket, add a more
  specific rule rather than tune a model (panel F10).
* Each rule gets a stable ``name`` attribute — written into the classifier
  output as ``matched_rule`` for provenance (panel F9 — Rob). Operators
  tune the rule, they do not retrain anything.
* All rules take the same input shape:
    {
        "subject": str,
        "sender_email": str,
        "sender_name": str,
        "received_datetime": datetime,  # tz-naive
        "is_read": bool,
        ...
    }
* Rules are pure — no I/O, no state.

The rule set below was derived from a 500-message sample of Chris's real
inbox pulled on 2026-11-16 (see ``output/inbox-sample-20261116-1635.json``).
Every rule names the cluster it came from in its docstring. When the rule
needs tuning, start from that sample.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Tuple


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

Message = Dict[str, Any]
# Per-run context used by rules that need cross-message signals (e.g.
# "did Chris ever reply to this sender?"). Keys currently in use:
#   - ``sent_to_counts``: ``Dict[str, int]`` — lowercased recipient email →
#     number of messages Chris sent TO that address in the lookback window.
#   - ``sender_counts``: ``Dict[str, int]`` — lowercased sender email →
#     number of inbox messages seen from that sender in the sample.
# Rules that don't need context accept it anyway and ignore it, so the
# dispatcher stays uniform.
Context = Optional[Dict[str, Any]]
RuleResult = Optional[Tuple[str, str]]
Rule = Callable[[Message, Context], RuleResult]

# Valid buckets. Classifier and executor both read from this.
BUCKETS = (
    "safe-delete",
    "resolved-ticket",
    "low-signal-client",
    "needs-action",
    "keep",
)


def _subject(msg: Message) -> str:
    return (msg.get("subject") or "").strip()


def _sender(msg: Message) -> str:
    return (msg.get("sender_email") or "").lower().strip()


def _sender_domain(msg: Message) -> str:
    s = _sender(msg)
    return s.split("@", 1)[1] if "@" in s else ""


def _received_dt(msg: Message) -> Optional[datetime]:
    v = msg.get("received_datetime")
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        try:
            return datetime.fromisoformat(v.replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            return None
    return None


def _age_days(msg: Message, *, now: Optional[datetime] = None) -> Optional[float]:
    """Message age in days. ``None`` if the timestamp can't be parsed."""
    dt = _received_dt(msg)
    if dt is None:
        return None
    reference = now or datetime.now()
    return (reference - dt).total_seconds() / 86400.0


def _rule(
    name: str,
    bucket: str,
    predicate: Callable[..., bool],
) -> Rule:
    """Wrap a predicate into a named rule callable.

    Predicates may be 1-arg ``(msg) -> bool`` or 2-arg
    ``(msg, context) -> bool``. The wrapper detects arity once at
    registration and dispatches accordingly, so both styles coexist.
    """
    assert bucket in BUCKETS, f"invalid bucket: {bucket}"

    # Introspect predicate arity so legacy 1-arg predicates still work.
    try:
        import inspect

        sig = inspect.signature(predicate)
        wants_context = len(sig.parameters) >= 2
    except (TypeError, ValueError):
        wants_context = False

    def _impl(msg: Message, context: Context = None) -> RuleResult:
        matched = predicate(msg, context) if wants_context else predicate(msg)
        return (bucket, name) if matched else None

    _impl.__name__ = name
    _impl.rule_name = name  # type: ignore[attr-defined]
    _impl.bucket = bucket  # type: ignore[attr-defined]
    return _impl


# ---------------------------------------------------------------------------
# Patterns used by multiple rules
# ---------------------------------------------------------------------------

_ACTION_SUBJECT_RE = re.compile(
    r"\b("
    r"please (?:review|advise|approve|confirm|respond|reply|action|fix|check)"
    r"|action required|action needed|approval needed|approve|urgent"
    r"|needs your (?:approval|attention|input|review)"
    r"|can you |could you |would you |will you "
    r"|questions?\?|time to review"
    r")\b",
    re.I,
)

_QUESTION_END_RE = re.compile(r"\?\s*$")

_CI_FAIL_RE = re.compile(
    r"\b(build failed|run failed|pr run failed|pipeline failed|ci failed"
    r"|workflow failed|tests failed|deploy(?:ment)? failed)\b",
    re.I,
)

_CI_SUCCESS_RE = re.compile(
    r"\[(?:build succeeded|build success|pipeline passed|ci passed|run succeeded)\]",
    re.I,
)

_UNSUBSCRIBE_HINT_RE = re.compile(
    r"\b(unsubscribe|newsletter|webinar|sign up now|% off|free trial"
    r"|exclusive offer|limited time|sale ends|promo code)\b",
    re.I,
)

# Bolt 6 Phase B.2 (from A3 marketplace scan): list-unsubscribe-header is
# the single highest-precision bulk-sender signal in commercial AI mail
# tools. OutlookClient doesn't expose raw headers today — so as a cheap
# fallback we look for the markers that bulk senders almost always
# render into the message body alongside the header. Any one of these
# phrases + the word "unsubscribe" in the body preview is near-perfect
# bulk precision.
_LIST_UNSUBSCRIBE_BODY_RE = re.compile(
    r"("
    r"unsubscribe (?:here|link|from this|from all|at any time)"
    r"|click (?:here )?to unsubscribe"
    r"|to unsubscribe,"
    r"|manage (?:your )?(?:email )?(?:subscription|preferences|notifications)"
    r"|update (?:your )?(?:email )?(?:subscription|preferences|notifications)"
    r"|email preferences"
    r"|you(?:'| a)?re receiving this (?:email |message )?because"
    r"|you are receiving this (?:email |message )?because"
    r"|you have received this (?:email|message) because"
    r"|view (?:this )?(?:email |message )?in (?:your )?browser"
    r"|was sent to you by"
    r")",
    re.I,
)

_RESOLVED_HINT_RE = re.compile(
    r"\b(resolved|closed|completed|complete|finished|has been closed)\b",
    re.I,
)

_TICKET_ID_RE = re.compile(r"(?:ticket|case|incident)\s*#?\s*\d{3,}", re.I)

# Automated-sender local-parts. Narrower than "contains @" — these are the
# local-part markers most often used by machines, not humans.
_AUTOMATED_LOCAL_RE = re.compile(
    r"^(?:"
    r"no-?reply|do-?not-?reply|notifications?|alerts?|mailer-daemon"
    r"|automated|robot|team|hello|info|support|digest|news"
    r"|forwarding-noreply|googlealerts-noreply|cloudplatform-noreply"
    r"|googleaistudio-noreply|donotreply"
    r")@",
    re.I,
)

# Marketing / cold-outbound / promo domains observed in the sample.
_MARKETING_DOMAINS = {
    "serchen.com",            # cold outbound — 8 msgs
    "engage.canva.com",       # product marketing — 5 msgs
    "brightedge.com",         # SEO cold — 5 msgs
    "birdeye.com",            # SaaS vendor — 5 msgs
    "instantly.ai",           # outbound tool — 1 msg
    "campaign.eventbrite.com",
    "marketing.trustpilotinfo.com",
    "notify.aircall.io",      # aircall verification codes — short-lived
    "microsoftadvertising.com",
    "publer.com",             # social scheduler
    "e.godaddy.com",          # godaddy promos
    "mail.googlealerts",      # alerts digest
    "e.bitly.com",
    "c.bitly.com",
    "e.instantly.ai",
    "email.openai.com",       # OpenAI product marketing (not product-alert)
    "screenshotapi.net",
    "send.intercom.com",
    "send.zapier.com",        # Zapier newsletter
    "mystudio.academy",       # vendor platform — bulk vs needs-action handled below
    "pb02.wixemails.com",     # Wix hosted bulk
    "aws-marketing-email-replies.amazon.com",  # AWS marketing drip
    "customer.aircall.io",    # Aircall newsletter
}

# Senders whose local-part marks them as marketing/outbound, independent of
# domain. Used as a cheap belt-and-braces for the long tail.
_MARKETING_LOCAL_RE = re.compile(
    r"^(?:"
    r"marketing|discover|news|hello|offers|team|promo|updates?"
    r"|aws-marketing-email-replies|digest|press|community|growth"
    r")@",
    re.I,
)

# Platform-notification domains — automated, archivable.
_PLATFORM_NOTIFICATION_DOMAINS = {
    "sns.amazonaws.com",
    "sharepointonline.com",
    "teams.mail.microsoft",
    "atlassian.net",
}

_INTERNAL_DOMAIN = "membersolutions.com"

# Operator — used by Bolt 6 behavioral rules that need to know "did Chris
# ever reply." Kept here (not in a config) because the rule set is
# single-tenant today. When multi-tenant lands, lift to a settings object.
_OPERATOR_EMAIL = "cfossenier@membersolutions.com"

# Cold-outreach / outbound-SaaS senders — patterns, not an enumeration.
# These regexes match against the full sender email (local@domain,
# lowercased). Hitting one of these + no reply-from-operator + age > 30d
# is strong enough signal to safe-delete.
#
# Sources: observed in Chris's inbox 2026-11-16 + 2026-11-17 Bolt 5 sweep
# + well-known outbound-tooling domains. Regex form lets us catch
# per-tenant subdomains ("foo.outreach.io", "mail.salesloft.com") without
# enumerating every one.
_COLD_OUTREACH_DOMAIN_PATTERNS = [
    re.compile(p, re.I)
    for p in (
        r".*\.outreach\.io$",
        r".*outreach\.io$",
        r".*salesloft\.com$",
        r".*\.salesloft\.com$",
        r".*apollo\.io$",
        r".*\.apollo\.io$",
        r".*lemlist\.com$",
        r".*\.lemlist\.com$",
        r".*instantly\.ai$",
        r".*\.instantly\.ai$",
        r".*send\.intercom\.com$",
        r".*pb02\.wixemails\.com$",
        r".*wixemails\.com$",
        r".*mailchimpapp\.net$",
        r".*mail\.mailchimp\.com$",
        r".*sendgrid\.net$",
        r".*mandrillapp\.com$",
        r".*hubspotemail\.net$",  # outbound HubSpot blasts (not notifications.hubspot.com)
        r".*gartner\.com$",       # cold analyst outreach
        r".*reply\.io$",
        r".*woodpecker\.co$",
        r".*persistiq\.com$",
    )
]


def _matches_cold_outreach_domain(email: str) -> bool:
    """True when the sender domain matches any cold-outreach pattern."""
    if "@" not in email:
        return False
    domain = email.split("@", 1)[1].lower().strip()
    return any(p.match(domain) for p in _COLD_OUTREACH_DOMAIN_PATTERNS)


def _sent_scan_available(context: Context) -> bool:
    """True when a Sent Items scan returned USABLE data for this run.

    ``sent_to_counts`` is only present when :func:`build_sent_context`
    (or an equivalent) was called. Rules that depend on "did the operator
    reply" must short-circuit when:
      - the scan wasn't done (key missing), OR
      - the scan returned empty (JXA parse bug returned 0 recipients)

    An empty dict means either "operator has sent zero emails ever" (not
    plausible) or the recipient extraction is broken. Either way the rule
    cannot distinguish "ignored" from "never asked," so we fail safe.

    Learned 2026-11-19 Bolt 6 live validation: ``ignored-sender-cold``
    fired 120× against real customer correspondence because
    ``toRecipients()`` returned empty strings and the gate let it through.
    """
    if not context or "sent_to_counts" not in context:
        return False
    counts = context.get("sent_to_counts") or {}
    return len(counts) > 0


def _operator_has_replied(sender_email: str, context: Context) -> bool:
    """True if the operator has sent anything to this address in-window.

    Reads ``context['sent_to_counts']`` populated by
    :func:`build_sent_context`. Callers MUST guard with
    :func:`_sent_scan_available` first — this helper assumes the scan
    ran and only answers the per-sender question.
    """
    if not context:
        return False
    counts = context.get("sent_to_counts") or {}
    return counts.get(sender_email.lower().strip(), 0) > 0


# ---------------------------------------------------------------------------
# Rules — ORDERED. First match wins.
# ---------------------------------------------------------------------------
# Rule ordering principles:
#   1. needs-action beats everything — a human owes a reply if we see
#      an action signal, even if the sender is automated (CI failure,
#      "please review" request, etc.).
#   2. keep beats safe-delete for internal people and direct human replies.
#   3. resolved-ticket / low-signal-client sit between needs-action and
#      safe-delete so archivable noise gets archived, not deleted.
#   4. safe-delete catches the widest net of automated marketing + alarms
#      older than 7 days.
#
# Each rule's docstring names the sample cluster it came from.


# ===== safety exemptions (highest priority) =================================
#
# Added 2026-11-16 after Chris's walk-through surfaced four categories that
# the generic marketing/noreply rules kept sweeping into safe-delete. Each
# exemption routes to ``needs-action`` so the operator sees it explicitly.

def _r_gcp_project_at_risk(msg: Message) -> bool:
    """Google Cloud 'Project ... is at risk of suspension' — real action.

    Cluster: ``CloudPlatform-noreply@google.com``. If ignored, the project
    gets suspended. Walk-through 2026-11-16 flagged two of these
    (msi-web + SalesClutch) as must-see.
    """
    sender = _sender(msg)
    subj = _subject(msg)
    return (
        "cloudplatform-noreply@google.com" in sender
        and "at risk of suspension" in subj.lower()
    )


def _r_google_account_recoverability(msg: Message) -> bool:
    """Google Account recoverability / account-state notices — must-see.

    Cluster: ``no-reply@accounts.google.com``. Walk-through 2026-11-16:
    'Your Google Account is no longer recoverable' is load-bearing if it
    pertains to a real account — operator must review.
    """
    sender = _sender(msg)
    return "no-reply@accounts.google.com" in sender


def _r_msi_billing_alert(msg: Message) -> bool:
    """Internal MSI billing/credit-card alerts — operator must see.

    Cluster: ``Alerts@Membersolutions.com``. Example: 'Credit Card Decline'.
    Flagged in walk-through 2026-11-16 as needing operator eyes.
    """
    sender = _sender(msg)
    return sender == "alerts@membersolutions.com"


def _r_sharepoint_permission_change(msg: Message) -> bool:
    """SharePoint anonymous-access link changes — audit trail, operator-review.

    Cluster: ``no-reply@sharepointonline.com`` with 'changed an anonymous
    access link' in subject. Walk-through 2026-11-16: these are real
    permission audit-trail items (e.g. Jungle MMA YTD summary link).
    """
    sender = _sender(msg)
    subj = _subject(msg).lower()
    return (
        sender.endswith("@sharepointonline.com")
        and "anonymous access link" in subj
    )


# ===== needs-action =========================================================

def _r_build_or_ci_failed(msg: Message) -> bool:
    """Azure DevOps / GitHub CI failures — 10 in sample.

    Cluster: ``azuredevops@microsoft.com`` + ``notifications@github.com``.
    A failed build is always on the critical path.
    """
    return bool(_CI_FAIL_RE.search(_subject(msg)))


def _r_hubspot_task_assignment(msg: Message) -> bool:
    """HubSpot 'You've been assigned' notifications — 7 in sample.

    Cluster: ``noreply@notifications.hubspot.com``. These are literal
    work-tracking notifications; never auto-delete.
    """
    subj = _subject(msg).lower()
    sender = _sender(msg)
    return (
        sender.endswith("@notifications.hubspot.com")
        and (
            "you've been assigned" in subj
            or "you have been assigned" in subj
            or "new task assigned" in subj
        )
    )


def _looks_like_marketing_sender(msg: Message) -> bool:
    """Shared precision guard — True if the sender is a known marketing source.

    Combines domain enumeration (``_MARKETING_DOMAINS``) with a local-part
    regex (``_MARKETING_LOCAL_RE``) so we catch the long tail cheaply.
    """
    if _sender_domain(msg) in _MARKETING_DOMAINS:
        return True
    if _MARKETING_LOCAL_RE.match(_sender(msg)):
        return True
    return False


def _r_action_verb_in_subject(msg: Message) -> bool:
    """Subject contains an action verb aimed at Chris.

    E.g. "please review", "action required", "can you …". We skip known
    marketing senders — cold-outbound mail often mimics this voice
    ("Can payroll handle this one?") and would otherwise poison the
    needs-action bucket.
    """
    if _looks_like_marketing_sender(msg):
        return False
    return bool(_ACTION_SUBJECT_RE.search(_subject(msg)))


def _r_question_in_subject(msg: Message) -> bool:
    """Subject ends in '?' — a question directed at the recipient.

    Observed: 13 in sample. We exclude marketing senders; those are the
    dominant false-positive source (discover@serchen.com, marketing@…).
    """
    if _looks_like_marketing_sender(msg):
        return False
    return bool(_QUESTION_END_RE.search(_subject(msg)))


def _r_direct_from_internal_person(msg: Message) -> bool:
    """A real internal person emailing us, non-automated.

    Internal domain = ``@membersolutions.com``, local part is NOT a
    no-reply/alias. Covers coworkers writing directly (14 psimko + 15 aleary
    + 14 starreeves … etc.).

    Technically this maps to ``keep`` more than ``needs-action`` — but we
    surface internal people in the needs-action tier because mis-archiving
    a coworker reply is worse than mis-archiving a vendor newsletter.
    """
    sender = _sender(msg)
    if not sender.endswith(f"@{_INTERNAL_DOMAIN}"):
        return False
    if _AUTOMATED_LOCAL_RE.match(sender):
        return False
    return True


# ===== resolved-ticket ======================================================

def _r_hubspot_ticket_comment(msg: Message) -> bool:
    """HubSpot 'New comments on the ticket "<...>"' — several in sample.

    Cluster: ``noreply@notifications.hubspot.com`` with "new comments on
    the ticket" in subject. Archivable after you've read it; not deletable.
    """
    subj = _subject(msg).lower()
    sender = _sender(msg)
    return (
        sender.endswith("@notifications.hubspot.com")
        and "new comments on the ticket" in subj
    )


def _r_resolved_or_closed_notification(msg: Message) -> bool:
    """Subject says the ticket/case/incident is resolved, closed, completed.

    Cluster: generic resolved/closed noise. Combined with a ticket id or
    an automated sender for precision.
    """
    subj = _subject(msg)
    if not _RESOLVED_HINT_RE.search(subj):
        return False
    # Precision guard: must also look like a ticket notification, otherwise
    # "Project complete — final docs attached" would get mis-bucketed.
    return bool(_TICKET_ID_RE.search(subj) or _AUTOMATED_LOCAL_RE.match(_sender(msg)))


def _r_zendesk_support_update(msg: Message) -> bool:
    """Vendor Zendesk notifications (Jonas, Pyxis, Zenoti support).

    Cluster: ``noreply@*.zendesk.com``, ``support@zenoti.com``. Usually
    ticket activity worth archiving but not deleting.
    """
    sender = _sender(msg)
    return (
        ".zendesk.com" in sender
        or sender == "support@zenoti.com"
        or sender == "noreply@zenoti.com"
    )


# ===== low-signal-client ====================================================

def _r_mystudio_failed_payment_rerun(msg: Message) -> bool:
    """MyStudio 'Failed Payment Re-run Details' — 9 in sample.

    Cluster: ``success@mystudio.academy``. Routine client FYI mail — Chris
    doesn't act on these individually; archive.
    """
    sender = _sender(msg)
    subj = _subject(msg).lower()
    return (
        sender.endswith("@mystudio.academy")
        and "failed payment re-run" in subj
    )


def _r_hubspot_tasks_today_digest(msg: Message) -> bool:
    """HubSpot "Your tasks for today" daily digest — 6 in sample.

    Cluster: ``noreply@notifications.hubspot.com``. Daily digest, read
    same day, then archivable. We keep it out of ``safe-delete`` because
    the aggregate can be useful when revisiting a week.
    """
    subj = _subject(msg).lower()
    sender = _sender(msg)
    return (
        sender.endswith("@notifications.hubspot.com")
        and "your tasks for today" in subj
    )


def _r_workday_task(msg: Message) -> bool:
    """Workday HR workflow mail — 5 in sample.

    Cluster: ``talentmanagementsolution@myworkday.com``. Usually routine
    (acknowledge a policy, review a form). Not delete-worthy; archive after
    action.
    """
    return _sender(msg).endswith("@myworkday.com")


# ===== Bolt 6 behavioral rules (2026-11-17) =================================
#
# These rules look at CROSS-MESSAGE signals — "did Chris reply?", "how
# old is this thread?", "has the sender gone dormant?". They go above
# the classic safe-delete rules so pattern-based detection wins over
# domain enumeration on the long tail.
#
# All five require a fresh ``context`` dict (see :func:`build_sent_context`).
# Rules short-circuit to False when no Sent-Items scan was built — this
# preserves the behavior of every existing ``classify(msg)`` caller
# that doesn't pass a context.


def _r_ignored_sender_cold(msg: Message, context: Context) -> bool:
    """Operator has NEVER replied to this sender and oldest msg > 60 days.

    Signal: sender has ≥ 3 messages in the inbox sample, zero sent-items
    hits for the operator, and this message is > 60 days old. If Chris
    never wrote back in two months, the sender is cold.

    Requires:
      - ``context['sent_to_counts']`` from :func:`build_sent_context`
      - ``context['sender_counts']`` (optional) — map of sender_email to
        inbox-sample count. If absent, we skip the volume floor and
        rely on the "never replied" + age signal alone.
    """
    if not _sent_scan_available(context):
        return False
    sender = _sender(msg)
    if not sender:
        return False
    # Skip internal — internal-person owns that bucket.
    if sender.endswith(f"@{_INTERNAL_DOMAIN}"):
        return False
    # Skip the operator themselves (defensive; sometimes shows up as a bcc echo).
    if sender == _OPERATOR_EMAIL.lower():
        return False
    if _operator_has_replied(sender, context):
        return False
    age = _age_days(msg)
    if age is None or age <= 60:
        return False
    # Optional volume floor — only fires when the context exposes
    # per-sender inbox counts. Keeps us out of false-positive territory
    # on rare one-off senders.
    sender_counts = (context or {}).get("sender_counts") or {}
    if sender_counts:
        if sender_counts.get(sender, 0) < 3:
            return False
    return True


def _r_dead_thread_reference(msg: Message, context: Context) -> bool:
    """Reply thread (``Re:`` chain) > 90 days old, operator spoke last.

    Signal: subject starts with ``Re:`` (or contains a ``Re:`` chain)
    AND received > 90 days ago AND the operator has sent at least one
    message to this address within the window. That combination means
    "Chris already replied; nothing new came back." Archivable.

    This is an archive-bucket rule (routes to ``resolved-ticket``), not
    a delete — these are prior conversations that may be worth
    referencing later.
    """
    if not _sent_scan_available(context):
        return False
    subj = _subject(msg)
    if not subj:
        return False
    low = subj.lower()
    # "Re:" chain detection — covers "Re:", "RE:", "Re: Re:", "FWD: Re:".
    if not (
        low.startswith("re:")
        or low.startswith("re ")
        or " re:" in low[:25]
    ):
        return False
    age = _age_days(msg)
    if age is None or age <= 90:
        return False
    sender = _sender(msg)
    if not sender:
        return False
    # Internal people are owned by internal-person; don't archive those.
    if sender.endswith(f"@{_INTERNAL_DOMAIN}"):
        return False
    # Operator must have engaged at least once — otherwise it's a cold
    # thread the operator ignored, which belongs to ignored-sender-cold.
    if not _operator_has_replied(sender, context):
        return False
    return True


def _r_cold_outreach_dormant(msg: Message, context: Context) -> bool:
    """Cold-outreach domain + no operator reply + > 30 days.

    Signal: sender domain matches one of the known outbound-tooling
    patterns in :data:`_COLD_OUTREACH_DOMAIN_PATTERNS`, the operator
    never replied, and the message is > 30 days old. This is the
    "SalesLoft drip that never got a response" cluster.
    """
    if not _sent_scan_available(context):
        return False
    sender = _sender(msg)
    if not sender or not _matches_cold_outreach_domain(sender):
        return False
    if _operator_has_replied(sender, context):
        return False
    age = _age_days(msg)
    return age is not None and age > 30


def _r_bulk_marketing_tail(msg: Message, context: Context) -> bool:
    """Known marketing sender + > 30 days + unread.

    Signal: sender matches :func:`_looks_like_marketing_sender`
    (domain or local-part), the message is > 30 days old, AND it's
    unread. Marketing mail Chris never even opened is the clearest
    delete-candidate in the inbox — Chris's explicit rule:
    "unread + junk = delete".

    Note: distinct from the existing ``marketing-sender`` rule, which
    fires regardless of age/read-state. This rule sits ABOVE the
    broader one so unread-tail marketing gets its more specific label.
    """
    if not _looks_like_marketing_sender(msg):
        return False
    if msg.get("is_read") is not False:  # treat missing as "assume read"
        return False
    age = _age_days(msg)
    return age is not None and age > 30


def _r_unread_automation_30d(msg: Message, context: Context) -> bool:
    """Automated-sender pattern + unread + > 30 days.

    Signal: ``is_read == False`` AND local-part matches an automated
    pattern (``noreply`` / ``notifications`` / ``alerts`` / ...) AND
    > 30 days old. Same "unread + junk = delete" principle as
    ``bulk-marketing-tail``, but for automation that isn't marketing.

    Exemption rules still win because they sit higher in RULES — a
    CloudPlatform at-risk message or SharePoint permission change
    won't fall here even if unread and old.
    """
    if msg.get("is_read") is not False:
        return False
    sender = _sender(msg)
    if not sender or not _AUTOMATED_LOCAL_RE.match(sender):
        return False
    age = _age_days(msg)
    return age is not None and age > 30


# ===== safe-delete ==========================================================

def _r_cloudwatch_alarm_old(msg: Message) -> bool:
    """CloudWatch ALARM emails older than 3 days — 24 in sample.

    Cluster: ``no-reply@sns.amazonaws.com`` with "ALARM:" subject prefix.
    Fresh alarms stay (operator may still be investigating); stale ones go.
    """
    sender = _sender(msg)
    subj = _subject(msg)
    if sender != "no-reply@sns.amazonaws.com":
        return False
    if not subj.startswith("ALARM:"):
        return False
    age = _age_days(msg)
    return age is not None and age > 3


def _r_ci_success_old(msg: Message) -> bool:
    """CI/build-succeeded notifications older than 3 days.

    Cluster: ``[Build succeeded]`` from Azure DevOps and parallel patterns.
    Successes are confirmations — valuable the day of, noise after.
    """
    if not _CI_SUCCESS_RE.search(_subject(msg)):
        return False
    age = _age_days(msg)
    return age is not None and age > 3


def _r_dependabot_old(msg: Message) -> bool:
    """Dependabot / dependency-bump PR notifications older than 7 days.

    Cluster: ``notifications@github.com`` with "chore(deps)" or "Bump".
    These age out fast — the PR either merged or was superseded.
    """
    subj = _subject(msg)
    is_dep = bool(
        re.search(r"chore\(deps\)|dependabot|bump .+ from .+ to ", subj, re.I)
    )
    if not is_dep:
        return False
    age = _age_days(msg)
    return age is not None and age > 7


def _r_google_alerts_digest(msg: Message) -> bool:
    """Google Alerts noreply — 7 in sample, pure digest.

    Cluster: ``googlealerts-noreply@google.com``. No live follow-up ever
    comes of these. Delete.
    """
    return _sender(msg) == "googlealerts-noreply@google.com"


def _r_list_unsubscribe_bulk(msg: Message) -> bool:
    """Body preview contains List-Unsubscribe-style markers — bulk sender.

    Bolt 6 Phase B.2 (win from A3 marketplace scan). Commercial AI mail
    tools lean on the ``List-Unsubscribe`` RFC header as the single
    highest-precision bulk-sender signal — near-perfect precision at
    zero LLM cost. OutlookClient does not surface headers today, so we
    use a body-side fallback: any one of the common "unsubscribe here",
    "manage subscriptions", or "you're receiving this because" markers
    is strong enough to safe-delete. Narrower than
    ``unsubscribe-marketing`` (which only fires on the subject line) so
    it catches newsletter bodies with plain subjects.
    """
    body = (msg.get("body_preview") or "")
    if not body:
        return False
    return bool(_LIST_UNSUBSCRIBE_BODY_RE.search(body))


def _r_marketing_sender(msg: Message) -> bool:
    """Known marketing / cold-outbound senders — domain OR local-part match.

    Cluster: Serchen, BrightEdge, Birdeye, trustpilot, instantly.ai,
    microsoftadvertising, publer, godaddy promos, plus the long-tail of
    ``marketing@…``, ``discover@…``, ``news@…``, ``aws-marketing-…`` locals.
    """
    return _looks_like_marketing_sender(msg)


def _r_platform_notification_old(msg: Message) -> bool:
    """SharePoint / Teams / Atlassian / SNS platform-notification noise > 7 days.

    Cluster: ``no-reply@sharepointonline.com``, ``no-reply@teams.mail.microsoft``.
    These accumulate when you've been @-mentioned across threads.
    """
    if _sender_domain(msg) not in _PLATFORM_NOTIFICATION_DOMAINS:
        return False
    age = _age_days(msg)
    return age is not None and age > 7


def _r_unsubscribe_marketing(msg: Message) -> bool:
    """Subject contains marketing/newsletter signals like 'unsubscribe' or '% off'.

    Cluster: catches long-tail marketing we didn't domain-enumerate.
    """
    return bool(_UNSUBSCRIBE_HINT_RE.search(_subject(msg)))


def _r_generic_noreply_old(msg: Message) -> bool:
    """Catch-all: automated sender local-part + older than 14 days.

    Long tail. We intentionally keep the age threshold high so this rule
    only fires when it is almost certainly stale.
    """
    sender = _sender(msg)
    if not _AUTOMATED_LOCAL_RE.match(sender):
        return False
    age = _age_days(msg)
    return age is not None and age > 14


# ---------------------------------------------------------------------------
# Rule registry — order matters
# ---------------------------------------------------------------------------

RULES: List[Rule] = [
    # safety exemptions (added 2026-11-16 from walk-through) — route to
    # needs-action so the operator sees these, never auto-swept.
    _rule("gcp-project-at-risk", "needs-action", _r_gcp_project_at_risk),
    _rule("google-account-recoverability", "needs-action", _r_google_account_recoverability),
    _rule("msi-billing-alert", "needs-action", _r_msi_billing_alert),
    _rule("sharepoint-permission-change", "needs-action", _r_sharepoint_permission_change),

    # needs-action (highest priority — never auto-delete action-triggered mail)
    _rule("ci-build-failed", "needs-action", _r_build_or_ci_failed),
    _rule("hubspot-task-assignment", "needs-action", _r_hubspot_task_assignment),
    _rule("action-verb-in-subject", "needs-action", _r_action_verb_in_subject),
    _rule("question-in-subject", "needs-action", _r_question_in_subject),
    _rule("internal-person", "needs-action", _r_direct_from_internal_person),

    # Bolt 6 behavioral archive — dead "Re:" threads where Chris already
    # replied and nothing new came back. Sits above the ticket rules so
    # a 90+ day old reply chain gets the behavioral label, not a generic
    # ticket-notification one.
    _rule("dead-thread-reference", "resolved-ticket", _r_dead_thread_reference),

    # resolved-ticket (archive, don't delete)
    _rule("hubspot-ticket-comment", "resolved-ticket", _r_hubspot_ticket_comment),
    _rule("resolved-or-closed-notification", "resolved-ticket", _r_resolved_or_closed_notification),
    _rule("zendesk-support-update", "resolved-ticket", _r_zendesk_support_update),

    # low-signal-client (routine FYI)
    _rule("mystudio-failed-payment-rerun", "low-signal-client", _r_mystudio_failed_payment_rerun),
    _rule("hubspot-tasks-today-digest", "low-signal-client", _r_hubspot_tasks_today_digest),
    _rule("workday-notification", "low-signal-client", _r_workday_task),

    # Bolt 6 behavioral safe-delete — cross-message signals that beat
    # the classic domain enumeration. They sit above the pattern rules
    # so provenance stays precise (``cold-outreach-dormant`` is a more
    # useful matched_rule than ``marketing-sender`` when both apply).
    #
    # Order within this block, most specific → least specific:
    #   cold-outreach-dormant  (known outbound-SaaS domain)
    #   bulk-marketing-tail    (known marketing domain + unread + age)
    #   unread-automation-30d  (automated local-part + unread + age)
    #   ignored-sender-cold    (behavioral — any non-internal sender)
    _rule("cold-outreach-dormant", "safe-delete", _r_cold_outreach_dormant),
    _rule("bulk-marketing-tail", "safe-delete", _r_bulk_marketing_tail),
    _rule("unread-automation-30d", "safe-delete", _r_unread_automation_30d),
    _rule("ignored-sender-cold", "safe-delete", _r_ignored_sender_cold),

    # safe-delete (automated noise, age-gated)
    _rule("cloudwatch-alarm-old", "safe-delete", _r_cloudwatch_alarm_old),
    _rule("ci-success-old", "safe-delete", _r_ci_success_old),
    _rule("dependabot-old", "safe-delete", _r_dependabot_old),
    _rule("google-alerts-digest", "safe-delete", _r_google_alerts_digest),
    # List-Unsubscribe body markers (Bolt 6 Phase B.2 from A3 scan).
    # Sits above the marketing-sender enumeration so the long-tail
    # newsletters that aren't in _MARKETING_DOMAINS still get caught.
    _rule("list-unsubscribe-bulk", "safe-delete", _r_list_unsubscribe_bulk),
    _rule("marketing-sender", "safe-delete", _r_marketing_sender),
    _rule("platform-notification-old", "safe-delete", _r_platform_notification_old),
    _rule("unsubscribe-marketing", "safe-delete", _r_unsubscribe_marketing),
    _rule("generic-noreply-old", "safe-delete", _r_generic_noreply_old),
]


def classify(msg: Message, context: Context = None) -> Tuple[str, str]:
    """Apply RULES in order, return ``(bucket, matched_rule)``.

    Args:
        msg: Inbox message dict (see :class:`OutlookClient.search` output).
        context: Optional per-run context. See the :data:`Context` type
            alias at the top of this module for the supported keys. Rules
            that don't need it ignore it; rules that do (e.g.
            ``ignored-sender-cold``) read from it to detect "did Chris
            ever reply to this sender" without touching Outlook again.

    Falls through to ``("keep", "default")`` if no rule matches.
    """
    for rule in RULES:
        result = rule(msg, context)
        if result is not None:
            return result
    return ("keep", "default")


def build_sent_context(
    client: Any,
    *,
    days_lookback: int = 365,
    max_results: int = 5000,
    now: Optional[datetime] = None,
) -> Dict[str, Any]:
    """Scan Sent Items once and build a ``{recipient_email: count}`` map.

    Used by rules that need to detect "did the operator ever reply to this
    sender" without re-querying Outlook per message. Call once at the top
    of a triage run and pass the result into :func:`classify` via the
    ``context`` argument.

    Args:
        client: An :class:`OutlookClient` (duck-typed; must expose
            ``search(folder=..., date_from=..., date_to=..., max_results=...)``
            returning dicts with a ``to`` field).
        days_lookback: How far back to scan Sent Items. Defaults to one
            year — enough to catch dormant threads without blowing up
            runtime on huge sent folders.
        max_results: Upper bound on messages scanned. Guards against a
            pathological sent folder.
        now: Test hook for the current time.

    Returns:
        A context dict shaped like::

            {
                "sent_to_counts": {"alice@example.com": 3, ...},
                "scanned_messages": 142,
                "days_lookback": 365,
            }

        Safe to pass directly to :func:`classify`.
    """
    reference = now or datetime.now()
    since = reference - timedelta(days=days_lookback)

    try:
        messages = client.search(
            folder="Sent Items",
            date_from=since.strftime("%Y-%m-%d"),
            date_to=reference.strftime("%Y-%m-%d"),
            max_results=max_results,
        )
    except Exception:  # pragma: no cover — defensive; empty context is safe fallback
        messages = []

    counts: Dict[str, int] = {}
    for m in messages:
        to_field = (m.get("to") or "").strip()
        if not to_field:
            continue
        # ``to`` is a comma-separated string in both backends; split and
        # normalize each address. We keep only the raw email token — no
        # name parsing, no angle-bracket tolerance (search() has already
        # projected these to plain addresses in the AppleScript backend,
        # and Graph returns address-only forms).
        for raw in to_field.split(","):
            addr = raw.strip().lower()
            if not addr:
                continue
            # Trim any "Name <addr>" form just in case.
            if "<" in addr and ">" in addr:
                addr = addr.split("<", 1)[1].rstrip(">").strip()
            if "@" not in addr:
                continue
            counts[addr] = counts.get(addr, 0) + 1

    return {
        "sent_to_counts": counts,
        "scanned_messages": len(messages),
        "days_lookback": days_lookback,
    }


__all__ = [
    "BUCKETS",
    "RULES",
    "Rule",
    "Context",
    "classify",
    "build_sent_context",
]
