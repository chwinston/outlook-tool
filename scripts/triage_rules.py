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
RuleResult = Optional[Tuple[str, str]]
Rule = Callable[[Message], RuleResult]

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


def _rule(name: str, bucket: str, predicate: Callable[[Message], bool]) -> Rule:
    """Wrap a predicate into a named rule callable."""
    assert bucket in BUCKETS, f"invalid bucket: {bucket}"

    def _impl(msg: Message) -> RuleResult:
        return (bucket, name) if predicate(msg) else None

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
    # needs-action (highest priority — never auto-delete action-triggered mail)
    _rule("ci-build-failed", "needs-action", _r_build_or_ci_failed),
    _rule("hubspot-task-assignment", "needs-action", _r_hubspot_task_assignment),
    _rule("action-verb-in-subject", "needs-action", _r_action_verb_in_subject),
    _rule("question-in-subject", "needs-action", _r_question_in_subject),
    _rule("internal-person", "needs-action", _r_direct_from_internal_person),

    # resolved-ticket (archive, don't delete)
    _rule("hubspot-ticket-comment", "resolved-ticket", _r_hubspot_ticket_comment),
    _rule("resolved-or-closed-notification", "resolved-ticket", _r_resolved_or_closed_notification),
    _rule("zendesk-support-update", "resolved-ticket", _r_zendesk_support_update),

    # low-signal-client (routine FYI)
    _rule("mystudio-failed-payment-rerun", "low-signal-client", _r_mystudio_failed_payment_rerun),
    _rule("hubspot-tasks-today-digest", "low-signal-client", _r_hubspot_tasks_today_digest),
    _rule("workday-notification", "low-signal-client", _r_workday_task),

    # safe-delete (automated noise, age-gated)
    _rule("cloudwatch-alarm-old", "safe-delete", _r_cloudwatch_alarm_old),
    _rule("ci-success-old", "safe-delete", _r_ci_success_old),
    _rule("dependabot-old", "safe-delete", _r_dependabot_old),
    _rule("google-alerts-digest", "safe-delete", _r_google_alerts_digest),
    _rule("marketing-sender", "safe-delete", _r_marketing_sender),
    _rule("platform-notification-old", "safe-delete", _r_platform_notification_old),
    _rule("unsubscribe-marketing", "safe-delete", _r_unsubscribe_marketing),
    _rule("generic-noreply-old", "safe-delete", _r_generic_noreply_old),
]


def classify(msg: Message) -> Tuple[str, str]:
    """Apply RULES in order, return ``(bucket, matched_rule)``.

    Falls through to ``("keep", "default")`` if no rule matches.
    """
    for rule in RULES:
        result = rule(msg)
        if result is not None:
            return result
    return ("keep", "default")


__all__ = [
    "BUCKETS",
    "RULES",
    "Rule",
    "classify",
]
