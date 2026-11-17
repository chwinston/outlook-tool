"""Tests for the Phase 2 triage rule set and executor — Bolt 4 (B4.2.*).

Covers
------
* Each rule in ``scripts/triage_rules.py`` with a fixture message.
* ``classify()`` fall-through to the ``keep`` default.
* ``scripts/apply_triage.load_report`` freshness + missing-file paths.
* ``scripts/apply_triage.render_preview`` happy path + empty bucket.
* ``apply_triage.main`` dry-run (no --apply) is a no-op and never touches
  the client; --apply branches on ``verified`` not ``status``.
"""

from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "scripts"))

from outlook_tool import OperationResult  # noqa: E402

import triage_rules  # noqa: E402
import apply_triage  # noqa: E402


# ---------------------------------------------------------------------------
# Fixture helper
# ---------------------------------------------------------------------------

def _msg(
    *,
    sender_email: str = "someone@example.com",
    sender_name: str = "Someone",
    subject: str = "hello",
    received_datetime: datetime | None = None,
    message_id: str = "as_msg_1",
    is_read: bool = False,
    folder: str = "Inbox",
    to: str = "",
    body_preview: str = "",
) -> dict:
    return {
        "id": message_id,
        "sender_email": sender_email,
        "sender_name": sender_name,
        "subject": subject,
        "received_datetime": received_datetime or datetime.now(),
        "is_read": is_read,
        "folder": folder,
        "to": to,
        "body_preview": body_preview,
    }


# ===========================================================================
# triage_rules — one assertion per rule
# ===========================================================================

class TestTriageRules:
    # ===== safety exemptions (added 2026-11-16 from walk-through) =============

    def test_gcp_project_at_risk_is_needs_action(self):
        m = _msg(
            sender_email="CloudPlatform-noreply@google.com",
            subject="Your Project: msi-web is at risk of suspension",
        )
        assert triage_rules.classify(m) == ("needs-action", "gcp-project-at-risk")

    def test_google_account_recoverability_is_needs_action(self):
        m = _msg(
            sender_email="no-reply@accounts.google.com",
            subject="Your Google Account is no longer recoverable",
        )
        assert triage_rules.classify(m) == ("needs-action", "google-account-recoverability")

    def test_msi_billing_alert_is_needs_action(self):
        m = _msg(
            sender_email="Alerts@Membersolutions.com",
            subject="Member Solutions Alert - Credit Card Decline",
        )
        assert triage_rules.classify(m) == ("needs-action", "msi-billing-alert")

    def test_sharepoint_permission_change_is_needs_action(self):
        m = _msg(
            sender_email="no-reply@sharepointonline.com",
            subject="Todd Blyth has changed an anonymous access link to \"JUNGLE-2026-YTD-SUMMARY\"",
            received_datetime=datetime.now() - timedelta(days=28),
        )
        assert triage_rules.classify(m) == ("needs-action", "sharepoint-permission-change")

    # ===========================================================================

    def test_ci_build_failed_catches_azure_devops(self):
        m = _msg(
            sender_email="azuredevops@microsoft.com",
            subject="[Build failed] HubSpotCallSync - main - DevOps - 4c99",
        )
        assert triage_rules.classify(m) == ("needs-action", "ci-build-failed")

    def test_ci_build_failed_catches_github(self):
        m = _msg(
            sender_email="notifications@github.com",
            subject="[msifoss/Anny] Run failed: CI - main (5a67022)",
        )
        assert triage_rules.classify(m) == ("needs-action", "ci-build-failed")

    def test_hubspot_task_assignment(self):
        m = _msg(
            sender_email="noreply@notifications.hubspot.com",
            subject="You've been assigned the task \"Raymond\"",
        )
        assert triage_rules.classify(m) == ("needs-action", "hubspot-task-assignment")

    def test_action_verb_in_subject(self):
        m = _msg(
            sender_email="CloudPlatform-noreply@google.com",
            subject="Action required: your billing account is past due",
        )
        assert triage_rules.classify(m) == ("needs-action", "action-verb-in-subject")

    def test_action_verb_from_marketing_sender_not_action(self):
        """Serchen is in _MARKETING_DOMAINS — must NOT land in needs-action."""
        m = _msg(
            sender_email="discover@serchen.com",
            subject="Can payroll handle this one?",
        )
        bucket, _ = triage_rules.classify(m)
        assert bucket == "safe-delete"

    def test_question_in_subject(self):
        m = _msg(
            sender_email="rosemary@battinroof.com",
            subject="Question?",
        )
        assert triage_rules.classify(m) == ("needs-action", "question-in-subject")

    def test_internal_person(self):
        m = _msg(sender_email="aleary@membersolutions.com", subject="Jungle MMA Follow Up Call")
        assert triage_rules.classify(m) == ("needs-action", "internal-person")

    def test_internal_noreply_is_NOT_internal_person(self):
        """An alert from an internal automated address must not fall into
        the internal-person rule — it has its own msi-billing-alert path
        (added 2026-11-16 from walk-through)."""
        m = _msg(sender_email="alerts@membersolutions.com", subject="Hello")
        bucket, matched = triage_rules.classify(m)
        # The msi-billing-alert rule routes Alerts@Membersolutions.com to
        # needs-action with its own rule — distinct from internal-person.
        assert matched != "internal-person"
        assert matched == "msi-billing-alert"

    def test_hubspot_ticket_comment(self):
        m = _msg(
            sender_email="noreply@notifications.hubspot.com",
            subject='New comments on the ticket "Michelle Jones (#4688261-1)"',
        )
        assert triage_rules.classify(m) == ("resolved-ticket", "hubspot-ticket-comment")

    def test_resolved_or_closed_notification(self):
        m = _msg(
            sender_email="no-reply@vendor.com",
            subject="Ticket #87213 resolved — thank you",
        )
        assert triage_rules.classify(m) == ("resolved-ticket", "resolved-or-closed-notification")

    def test_zendesk_support_update(self):
        m = _msg(
            sender_email="noreply@jonassoftwarehelp.zendesk.com",
            subject="Re: Support request #4432",
        )
        assert triage_rules.classify(m) == ("resolved-ticket", "zendesk-support-update")

    def test_mystudio_failed_payment_rerun(self):
        m = _msg(
            sender_email="success@mystudio.academy",
            subject="Failed Payment Re-run Details (The Jungle MMA).",
        )
        assert triage_rules.classify(m) == ("low-signal-client", "mystudio-failed-payment-rerun")

    def test_hubspot_tasks_today_digest(self):
        m = _msg(
            sender_email="noreply@notifications.hubspot.com",
            subject="Your tasks for today",
        )
        assert triage_rules.classify(m) == ("low-signal-client", "hubspot-tasks-today-digest")

    def test_workday_notification(self):
        m = _msg(
            sender_email="talentmanagementsolution@myworkday.com",
            subject="You have 1 task awaiting your attention",
        )
        bucket, _ = triage_rules.classify(m)
        # action-verb may win on "awaiting" wording. Assert a safe bucket.
        assert bucket in ("needs-action", "low-signal-client")

    def test_cloudwatch_alarm_old(self):
        old = datetime.now() - timedelta(days=5)
        m = _msg(
            sender_email="no-reply@sns.amazonaws.com",
            subject="ALARM: \"callhero-foo\" in US East",
            received_datetime=old,
        )
        assert triage_rules.classify(m) == ("safe-delete", "cloudwatch-alarm-old")

    def test_cloudwatch_alarm_fresh_is_kept(self):
        """Fresh alarm — within 3 days — must NOT be auto-deleted."""
        fresh = datetime.now() - timedelta(hours=6)
        m = _msg(
            sender_email="no-reply@sns.amazonaws.com",
            subject="ALARM: \"callhero-foo\" in US East",
            received_datetime=fresh,
        )
        bucket, _ = triage_rules.classify(m)
        assert bucket != "safe-delete"

    def test_ci_success_old(self):
        old = datetime.now() - timedelta(days=5)
        m = _msg(
            sender_email="azuredevops@microsoft.com",
            subject="[Build succeeded] HubSpotCallSync",
            received_datetime=old,
        )
        assert triage_rules.classify(m) == ("safe-delete", "ci-success-old")

    def test_dependabot_old(self):
        old = datetime.now() - timedelta(days=10)
        m = _msg(
            sender_email="notifications@github.com",
            subject="Re: [msifoss/Anny] chore(deps)(deps): Bump google-auth from 2.48.0 to 2.62.0",
            received_datetime=old,
        )
        assert triage_rules.classify(m) == ("safe-delete", "dependabot-old")

    def test_google_alerts_digest(self):
        m = _msg(
            sender_email="googlealerts-noreply@google.com",
            subject="Google Alert - Member Solutions",
        )
        assert triage_rules.classify(m) == ("safe-delete", "google-alerts-digest")

    def test_marketing_sender_domain(self):
        m = _msg(sender_email="hello@birdeye.com", subject="[2026 Report] The State of AI Search")
        assert triage_rules.classify(m) == ("safe-delete", "marketing-sender")

    def test_marketing_sender_local_part(self):
        m = _msg(sender_email="marketing@novel-domain.io", subject="Our latest newsletter")
        assert triage_rules.classify(m) == ("safe-delete", "marketing-sender")

    def test_platform_notification_old(self):
        old = datetime.now() - timedelta(days=14)
        m = _msg(
            sender_email="no-reply@sharepointonline.com",
            subject="You were @mentioned on a page",
            received_datetime=old,
        )
        assert triage_rules.classify(m) == ("safe-delete", "platform-notification-old")

    def test_unsubscribe_marketing(self):
        m = _msg(
            sender_email="promo@weird.example",
            subject="Spring sale — 30% off, unsubscribe here",
        )
        bucket, rule = triage_rules.classify(m)
        # could match marketing-sender first (promo@ local part). Either
        # safe-delete rule is acceptable.
        assert bucket == "safe-delete"
        assert rule in ("marketing-sender", "unsubscribe-marketing")

    # ---- Bolt 6 B.2 — list-unsubscribe-bulk (A3 marketplace-scan win) ---

    def test_list_unsubscribe_bulk_from_body(self):
        """A novel-domain sender with a classic unsubscribe-body marker
        must land in safe-delete — this is the long-tail newsletter
        catchment the rule was designed for."""
        m = _msg(
            sender_email="hello@novel-newsletter-domain-xyz.io",
            subject="November Digest",
            body_preview=(
                "This is the November issue. Read more at the link. "
                "You're receiving this because you signed up at example.com. "
                "Unsubscribe here."
            ),
            received_datetime=datetime.now() - timedelta(days=2),
        )
        bucket, rule = triage_rules.classify(m)
        assert bucket == "safe-delete"
        assert rule == "list-unsubscribe-bulk"

    def test_list_unsubscribe_bulk_manage_preferences_marker(self):
        """The 'manage preferences' variant must also match."""
        m = _msg(
            sender_email="updates@brand.example",
            subject="Product update",
            body_preview=(
                "Here's what's new this month. Manage your email preferences "
                "at the link below."
            ),
            received_datetime=datetime.now() - timedelta(days=5),
        )
        bucket, rule = triage_rules.classify(m)
        assert bucket == "safe-delete"
        # marketing-sender might still win on senders with a marketing local
        # part; here the sender is novel so list-unsubscribe-bulk should win.
        assert rule == "list-unsubscribe-bulk"

    def test_list_unsubscribe_bulk_requires_marker(self):
        """A plain personal email with no unsubscribe marker must NOT
        match this rule (precision guard)."""
        m = _msg(
            sender_email="friend@external-person.example",
            subject="Dinner next week?",
            body_preview="Hey, are you free for dinner next Thursday?",
            received_datetime=datetime.now() - timedelta(days=1),
        )
        bucket, rule = triage_rules.classify(m)
        # The "?" in subject lands it in needs-action via question-in-subject;
        # the critical assertion is: list-unsubscribe-bulk must NOT fire.
        assert rule != "list-unsubscribe-bulk"

    def test_generic_noreply_old(self):
        old = datetime.now() - timedelta(days=20)
        m = _msg(
            sender_email="noreply@unknown-saas.io",
            subject="Monthly product update",
            received_datetime=old,
        )
        assert triage_rules.classify(m) == ("safe-delete", "generic-noreply-old")

    def test_default_keep(self):
        m = _msg(
            sender_email="constantin.manthey@gmx.net",
            subject="Updated Contract.",
            received_datetime=datetime.now() - timedelta(days=2),
        )
        assert triage_rules.classify(m) == ("keep", "default")


# ===========================================================================
# triage_rules.build_sent_context — Bolt 6 A1 sent-items helper
# ===========================================================================

class TestBuildSentContext:
    """``build_sent_context`` scans Sent Items and returns a count-by-
    recipient map used by the ``ignored-sender-cold`` rule."""

    def _client_with_sent(self, sent_messages):
        client = MagicMock()
        client.search.return_value = sent_messages
        return client

    def test_builds_recipient_counts(self):
        client = self._client_with_sent([
            {"to": "alice@example.com"},
            {"to": "alice@example.com, bob@example.com"},
            {"to": "bob@example.com"},
        ])
        ctx = triage_rules.build_sent_context(client, days_lookback=30)
        assert ctx["sent_to_counts"]["alice@example.com"] == 2
        assert ctx["sent_to_counts"]["bob@example.com"] == 2
        assert ctx["scanned_messages"] == 3
        assert ctx["days_lookback"] == 30

    def test_normalizes_case_and_whitespace(self):
        client = self._client_with_sent([
            {"to": "Alice@Example.COM"},
            {"to": "  alice@example.com  "},
        ])
        ctx = triage_rules.build_sent_context(client)
        assert ctx["sent_to_counts"]["alice@example.com"] == 2

    def test_strips_angle_bracket_form(self):
        client = self._client_with_sent([
            {"to": "Alice <alice@example.com>"},
        ])
        ctx = triage_rules.build_sent_context(client)
        assert ctx["sent_to_counts"].get("alice@example.com") == 1

    def test_empty_to_field_ignored(self):
        client = self._client_with_sent([
            {"to": ""},
            {"to": None},
            {},
        ])
        ctx = triage_rules.build_sent_context(client)
        assert ctx["sent_to_counts"] == {}
        assert ctx["scanned_messages"] == 3

    def test_search_called_with_sent_items(self):
        client = self._client_with_sent([])
        triage_rules.build_sent_context(
            client, days_lookback=365, max_results=5000,
        )
        client.search.assert_called_once()
        kwargs = client.search.call_args.kwargs
        assert kwargs["folder"] == "Sent Items"
        assert kwargs["max_results"] == 5000
        assert "date_from" in kwargs
        assert "date_to" in kwargs


# ===========================================================================
# Bolt 6 A1 — new behavioral pattern rules
# ===========================================================================

class TestNewPatternRules:
    """Behavioral rules added in Bolt 6 Phase A1.

    These rules require a per-run ``context`` dict (see
    :func:`triage_rules.build_sent_context`). Each rule has at least a
    positive test, a negative test, and an edge-case test that would
    trip up a less careful implementation.
    """

    # ---- ignored-sender-cold -------------------------------------------

    def test_ignored_sender_cold_positive(self):
        """Cold sender, > 60 days old, no reply from Chris → safe-delete."""
        m = _msg(
            sender_email="rep@unknown-vendor.io",
            subject="Checking in on our last conversation",
            received_datetime=datetime.now() - timedelta(days=75),
        )
        ctx = {
            "sent_to_counts": {},  # Chris has never written to this address
            "sender_counts": {"rep@unknown-vendor.io": 4},
        }
        assert triage_rules.classify(m, ctx) == (
            "safe-delete", "ignored-sender-cold",
        )

    def test_ignored_sender_cold_blocks_when_chris_replied(self):
        """If Chris EVER replied, the rule must NOT fire."""
        m = _msg(
            sender_email="rep@unknown-vendor.io",
            subject="Checking in on our last conversation",
            received_datetime=datetime.now() - timedelta(days=75),
        )
        ctx = {
            "sent_to_counts": {"rep@unknown-vendor.io": 1},
            "sender_counts": {"rep@unknown-vendor.io": 4},
        }
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "ignored-sender-cold"

    def test_ignored_sender_cold_requires_age(self):
        """Age < 60 days → rule must NOT fire even if never replied."""
        m = _msg(
            sender_email="rep@unknown-vendor.io",
            subject="Quick follow-up",
            received_datetime=datetime.now() - timedelta(days=30),
        )
        ctx = {
            "sent_to_counts": {},
            "sender_counts": {"rep@unknown-vendor.io": 4},
        }
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "ignored-sender-cold"

    def test_ignored_sender_cold_skips_internal(self):
        """Internal sender must stay with internal-person, not get
        auto-swept by the cold rule."""
        m = _msg(
            sender_email="coworker@membersolutions.com",
            subject="Old update",
            received_datetime=datetime.now() - timedelta(days=90),
        )
        ctx = {
            "sent_to_counts": {},
            "sender_counts": {"coworker@membersolutions.com": 10},
        }
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule == "internal-person"
        assert bucket == "needs-action"

    def test_ignored_sender_cold_no_context_noop(self):
        """Without a Sent-Items scan, the rule must NOT fire — preserves
        backward compat for existing ``classify(msg)`` callers."""
        m = _msg(
            sender_email="rep@unknown-vendor.io",
            subject="Checking in",
            received_datetime=datetime.now() - timedelta(days=75),
        )
        bucket, rule = triage_rules.classify(m)  # no context
        assert rule != "ignored-sender-cold"

    def test_ignored_sender_cold_volume_floor_enforced(self):
        """Volume floor: fewer than 3 inbox messages from this sender ⇒
        rule must NOT fire (avoid false positives on one-off senders)."""
        m = _msg(
            sender_email="rep@unknown-vendor.io",
            subject="Checking in",
            received_datetime=datetime.now() - timedelta(days=75),
        )
        ctx = {
            "sent_to_counts": {},
            "sender_counts": {"rep@unknown-vendor.io": 1},
        }
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "ignored-sender-cold"

    # ---- dead-thread-reference -----------------------------------------

    def test_dead_thread_reference_positive(self):
        """Re: thread > 90 days old + Chris replied once → archive."""
        m = _msg(
            sender_email="vendor@external.com",
            subject="Re: Contract Renewal Discussion",
            received_datetime=datetime.now() - timedelta(days=120),
        )
        ctx = {"sent_to_counts": {"vendor@external.com": 2}}
        assert triage_rules.classify(m, ctx) == (
            "resolved-ticket", "dead-thread-reference",
        )

    def test_dead_thread_reference_requires_re_prefix(self):
        """Fresh subject (no Re: chain) → rule must NOT fire."""
        m = _msg(
            sender_email="vendor@external.com",
            subject="New Contract Proposal",
            received_datetime=datetime.now() - timedelta(days=120),
        )
        ctx = {"sent_to_counts": {"vendor@external.com": 2}}
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "dead-thread-reference"

    def test_dead_thread_reference_requires_chris_engagement(self):
        """If Chris never wrote back, this is ignored-sender-cold
        territory, not dead-thread."""
        m = _msg(
            sender_email="vendor@external.com",
            subject="Re: Our product",
            received_datetime=datetime.now() - timedelta(days=120),
        )
        ctx = {"sent_to_counts": {}}  # no reply
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "dead-thread-reference"

    # ---- cold-outreach-dormant -----------------------------------------

    def test_cold_outreach_dormant_positive(self):
        """Known outbound-SaaS domain + no reply + > 30 days → safe-delete."""
        m = _msg(
            sender_email="hello@foo.outreach.io",
            subject="Can I get 15 minutes?",
            received_datetime=datetime.now() - timedelta(days=45),
        )
        ctx = {"sent_to_counts": {}}
        assert triage_rules.classify(m, ctx) == (
            "safe-delete", "cold-outreach-dormant",
        )

    def test_cold_outreach_dormant_salesloft(self):
        """SalesLoft subdomain matches the pattern."""
        m = _msg(
            sender_email="noreply@mail.salesloft.com",
            subject="Following up",
            received_datetime=datetime.now() - timedelta(days=60),
        )
        ctx = {"sent_to_counts": {}}
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule == "cold-outreach-dormant"
        assert bucket == "safe-delete"

    def test_cold_outreach_dormant_blocked_by_reply(self):
        """Operator replied to the outreach sender → rule must NOT fire."""
        m = _msg(
            sender_email="hello@foo.outreach.io",
            subject="Can I get 15 minutes?",
            received_datetime=datetime.now() - timedelta(days=45),
        )
        ctx = {"sent_to_counts": {"hello@foo.outreach.io": 1}}
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "cold-outreach-dormant"

    def test_cold_outreach_dormant_age_floor(self):
        """Under 30 days → not dormant yet."""
        m = _msg(
            sender_email="hello@foo.outreach.io",
            subject="Quick note",
            received_datetime=datetime.now() - timedelta(days=15),
        )
        ctx = {"sent_to_counts": {}}
        bucket, rule = triage_rules.classify(m, ctx)
        assert rule != "cold-outreach-dormant"

    # ---- bulk-marketing-tail -------------------------------------------

    def test_bulk_marketing_tail_positive(self):
        """Marketing-domain sender + unread + > 30 days → safe-delete."""
        m = _msg(
            sender_email="hello@birdeye.com",
            subject="Monthly product update",
            received_datetime=datetime.now() - timedelta(days=45),
            is_read=False,
        )
        assert triage_rules.classify(m) == (
            "safe-delete", "bulk-marketing-tail",
        )

    def test_bulk_marketing_tail_read_falls_through(self):
        """Read marketing mail → generic marketing-sender rule wins,
        bulk-marketing-tail must NOT claim it."""
        m = _msg(
            sender_email="hello@birdeye.com",
            subject="Monthly product update",
            received_datetime=datetime.now() - timedelta(days=45),
            is_read=True,
        )
        bucket, rule = triage_rules.classify(m)
        assert bucket == "safe-delete"
        # Read mail must NOT be labelled bulk-marketing-tail.
        assert rule != "bulk-marketing-tail"

    def test_bulk_marketing_tail_age_floor(self):
        """Fresh unread marketing → other rules may handle it, but NOT
        this one (needs the 30-day cliff)."""
        m = _msg(
            sender_email="hello@birdeye.com",
            subject="Monthly product update",
            received_datetime=datetime.now() - timedelta(days=5),
            is_read=False,
        )
        bucket, rule = triage_rules.classify(m)
        assert rule != "bulk-marketing-tail"

    # ---- unread-automation-30d -----------------------------------------

    def test_unread_automation_30d_positive(self):
        """Noreply automation + unread + > 30 days → safe-delete."""
        m = _msg(
            sender_email="noreply@unknown-saas.io",
            subject="Your weekly report",
            received_datetime=datetime.now() - timedelta(days=45),
            is_read=False,
        )
        assert triage_rules.classify(m) == (
            "safe-delete", "unread-automation-30d",
        )

    def test_unread_automation_30d_requires_unread(self):
        """Read automation falls through to generic-noreply-old, NOT
        unread-automation-30d."""
        m = _msg(
            sender_email="noreply@unknown-saas.io",
            subject="Your weekly report",
            received_datetime=datetime.now() - timedelta(days=45),
            is_read=True,
        )
        bucket, rule = triage_rules.classify(m)
        assert bucket == "safe-delete"
        assert rule != "unread-automation-30d"

    def test_unread_automation_30d_exemption_still_wins(self):
        """Exemption rules sit higher: GCP at-risk must NOT get
        swept into unread-automation-30d even when unread + old."""
        m = _msg(
            sender_email="CloudPlatform-noreply@google.com",
            subject="Your Project: msi-web is at risk of suspension",
            received_datetime=datetime.now() - timedelta(days=60),
            is_read=False,
        )
        bucket, rule = triage_rules.classify(m)
        assert rule == "gcp-project-at-risk"
        assert bucket == "needs-action"

    def test_unread_automation_30d_human_sender_not_matched(self):
        """Human sender (not noreply/notifications) → rule must NOT fire."""
        m = _msg(
            sender_email="jane.doe@partner.co",
            subject="Checking in",
            received_datetime=datetime.now() - timedelta(days=45),
            is_read=False,
        )
        bucket, rule = triage_rules.classify(m)
        assert rule != "unread-automation-30d"


# ===========================================================================
# apply_triage.load_report — freshness / errors
# ===========================================================================

class TestLoadReport:
    def _write_report(self, tmp_path: Path, generated: datetime) -> Path:
        path = tmp_path / "triage.json"
        path.write_text(json.dumps({
            "schema_version": 1,
            "generated_at": generated.isoformat(),
            "backend": "applescript",
            "sample_size": 0,
            "messages": [],
            "summary": {b: 0 for b in triage_rules.BUCKETS},
        }))
        return path

    def test_fresh_report_loads(self, tmp_path):
        now = datetime.now()
        path = self._write_report(tmp_path, now - timedelta(minutes=5))
        report = apply_triage.load_report(path, now=now)
        assert report["backend"] == "applescript"
        assert report["_age_seconds"] < apply_triage.MAX_REPORT_AGE_SECONDS

    def test_too_old_refuses(self, tmp_path):
        now = datetime.now()
        path = self._write_report(tmp_path, now - timedelta(hours=2))
        with pytest.raises(apply_triage.ReportTooOldError):
            apply_triage.load_report(path, now=now)

    def test_missing_file_errors(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            apply_triage.load_report(tmp_path / "nope.json")

    def test_missing_generated_at_errors(self, tmp_path):
        path = tmp_path / "bad.json"
        path.write_text(json.dumps({"messages": []}))
        with pytest.raises(ValueError, match="generated_at"):
            apply_triage.load_report(path)


# ===========================================================================
# apply_triage.render_preview
# ===========================================================================

class TestRenderPreview:
    def _report(self, *, messages: list) -> dict:
        return {
            "schema_version": 1,
            "generated_at": datetime.now().isoformat(),
            "_age_seconds": 120,
            "backend": "applescript",
            "sample_folder": "Inbox",
            "sample_size": len(messages),
            "window_days": 30,
            "summary": {b: 0 for b in triage_rules.BUCKETS},
            "messages": messages,
        }

    def test_empty_bucket_renders(self):
        md = apply_triage.render_preview(
            self._report(messages=[]),
            bucket="safe-delete",
            report_path=Path("/tmp/x.json"),
        )
        assert "## Triage Preview — safe-delete bucket (0 messages)" in md
        assert "*(no messages in this bucket)*" in md

    def test_safe_delete_bucket_rows(self):
        md = apply_triage.render_preview(
            self._report(messages=[
                {
                    "message_id": "as_msg_7",
                    "sender_email": "noreply@github.com",
                    "subject": "[GitHub] 100% Actions minutes used",
                    "received_at": "2026-11-15T10:11:00",
                    "bucket": "safe-delete",
                    "matched_rule": "generic-noreply-old",
                },
            ]),
            bucket="safe-delete",
            report_path=Path("/tmp/x.json"),
        )
        assert "| 1 | 2026-11-15T10:11" in md
        assert "noreply@github.com" in md
        assert "generic-noreply-old" in md

    def test_preview_only_bucket_notes_no_apply(self):
        md = apply_triage.render_preview(
            self._report(messages=[]),
            bucket="needs-action",
            report_path=Path("/tmp/x.json"),
        )
        assert "Preview-only" in md


# ===========================================================================
# apply_triage.main — dry-run and apply branches
# ===========================================================================

class TestApplyTriageMain:
    def _write_fresh_report(self, tmp_path: Path, messages: list) -> Path:
        path = tmp_path / "triage.json"
        path.write_text(json.dumps({
            "schema_version": 1,
            "generated_at": datetime.now().isoformat(),
            "backend": "applescript",
            "sample_folder": "Inbox",
            "sample_size": len(messages),
            "window_days": 30,
            "summary": {b: 0 for b in triage_rules.BUCKETS},
            "messages": messages,
        }))
        return path

    def test_dry_run_does_not_touch_client(self, tmp_path, capsys):
        report = self._write_fresh_report(tmp_path, messages=[
            {
                "message_id": "as_msg_3",
                "sender_email": "n@x.com",
                "subject": "x",
                "received_at": datetime.now().isoformat(),
                "bucket": "safe-delete",
                "matched_rule": "r1",
            },
        ])
        with patch("apply_triage.OutlookClient") as mock_cls:
            rc = apply_triage.main([str(report), "--bucket", "safe-delete"])
        mock_cls.assert_not_called()
        assert rc == 0
        out = capsys.readouterr().out
        assert "Triage Preview — safe-delete" in out

    def test_report_too_old_returns_2(self, tmp_path, capsys):
        # Hand-write an old report so load_report refuses it.
        path = tmp_path / "triage.json"
        path.write_text(json.dumps({
            "generated_at": (datetime.now() - timedelta(hours=3)).isoformat(),
            "messages": [],
            "backend": "applescript",
            "summary": {},
        }))
        rc = apply_triage.main([str(path), "--bucket", "safe-delete", "--apply"])
        assert rc == 2
        err = capsys.readouterr().err
        assert "REFUSING" in err

    def test_yes_really_permanent_refused(self, tmp_path):
        rc = apply_triage.main([
            str(tmp_path / "x.json"),
            "--bucket", "safe-delete",
            "--yes-really-permanent",
        ])
        assert rc == 2  # refused before file ever loaded

    def test_preview_only_bucket_apply_noop(self, tmp_path):
        report = self._write_fresh_report(tmp_path, messages=[
            {
                "message_id": "as_msg_1",
                "sender_email": "a@m.com",
                "subject": "needs you",
                "received_at": datetime.now().isoformat(),
                "bucket": "needs-action",
                "matched_rule": "r",
            },
        ])
        with patch("apply_triage.OutlookClient") as mock_cls:
            rc = apply_triage.main([str(report), "--bucket", "needs-action", "--apply"])
        # preview-only buckets must not instantiate the client.
        mock_cls.assert_not_called()
        assert rc == 0

    def test_apply_branches_on_verified_not_status(self, tmp_path, capsys):
        """Panel F1 / Tim: status==success but verified==False must not pass clean.

        Simulate the client returning status==success, verified==False. The
        executor must count it as unverified-success and return non-zero.
        """
        report = self._write_fresh_report(tmp_path, messages=[
            {
                "message_id": "as_msg_9",
                "sender_email": "marketing@x.com",
                "subject": "promo",
                "received_at": datetime.now().isoformat(),
                "bucket": "safe-delete",
                "matched_rule": "marketing-sender",
            },
        ])
        unverified = OperationResult(
            status="success",
            operation="delete",
            message_id="as_msg_9",
            verified=False,
            before_state={"folder": "Inbox"},
            after_state={"folder": "Deleted Items", "permanent": False},
            backend="applescript",
            elapsed_ms=50,
        )
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete.return_value = unverified
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply", "--no-bulk",
            ])
        assert rc == 1, "unverified-success must be surfaced as failure (not clean exit)"
        combined = capsys.readouterr().out
        assert "UNVERIFIED-success" in combined
        client.delete.assert_called_once_with(
            "as_msg_9", permanent=False, folder_name="Inbox",
        )

    def test_apply_verified_success_returns_0(self, tmp_path, capsys):
        report = self._write_fresh_report(tmp_path, messages=[
            {
                "message_id": "as_msg_1",
                "sender_email": "x@y.com",
                "subject": "s",
                "received_at": datetime.now().isoformat(),
                "bucket": "safe-delete",
                "matched_rule": "marketing-sender",
            },
        ])
        verified = OperationResult(
            status="success",
            operation="delete",
            message_id="as_msg_1",
            verified=True,
            backend="applescript",
            elapsed_ms=10,
        )
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete.return_value = verified
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply", "--no-bulk",
            ])
        assert rc == 0

    def test_fingerprint_recovery_when_initial_fails(self, tmp_path):
        """When the initial write fails with a 'message_not_found'-style error,
        apply_triage should attempt a fingerprint fallback via search()."""
        report = self._write_fresh_report(tmp_path, messages=[
            {
                "message_id": "as_msg_100",  # stale
                "sender_email": "x@y.com",
                "subject": "unique-subject-ABC",
                "received_at": datetime.now().isoformat(),
                "bucket": "safe-delete",
                "matched_rule": "marketing-sender",
            },
        ])
        stale_fail = OperationResult(
            status="failed",
            operation="delete",
            message_id="as_msg_100",
            verified=False,
            error_detail="pre-fetch failed: message_not_found",
            backend="applescript",
            elapsed_ms=5,
        )
        recovered_ok = OperationResult(
            status="success",
            operation="delete",
            message_id="as_msg_42",
            verified=True,
            backend="applescript",
            elapsed_ms=20,
        )
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete.side_effect = [stale_fail, recovered_ok]
            client.search.return_value = [{
                "id": "as_msg_42",
                "subject": "unique-subject-ABC",
                "sender_email": "x@y.com",
            }]
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply", "--no-bulk",
            ])
        assert rc == 0
        # Two calls — first with stale id, second with recovered id.
        calls = client.delete.call_args_list
        assert calls[0].args[0] == "as_msg_100"
        assert calls[1].args[0] == "as_msg_42"


# ===========================================================================
# Bolt 5 (B5.2) — apply_triage --bulk integration
# ===========================================================================


class TestApplyTriageBulkPath:
    """`--bulk` default-on routes safe-delete through client.delete_bulk,
    move buckets through client.move_bulk — one call per batch."""

    def _write_fresh_report(self, tmp_path: Path, messages: list) -> Path:
        path = tmp_path / "triage.json"
        path.write_text(json.dumps({
            "schema_version": 1,
            "generated_at": datetime.now().isoformat(),
            "backend": "applescript",
            "sample_folder": "Inbox",
            "sample_size": len(messages),
            "window_days": 30,
            "summary": {b: 0 for b in triage_rules.BUCKETS},
            "messages": messages,
        }))
        return path

    def _three_msgs(self):
        return [
            {
                "message_id": f"as_msg_{i}",
                "sender_email": f"a{i}@x.com",
                "subject": f"s{i}",
                "received_at": datetime.now().isoformat(),
                "bucket": "safe-delete",
                "matched_rule": "r",
            }
            for i in (5, 7, 9)
        ]

    def test_bulk_default_routes_safe_delete_through_delete_bulk(self, tmp_path):
        report = self._write_fresh_report(tmp_path, messages=self._three_msgs())
        bulk_results = [
            OperationResult(
                status="success", operation="delete",
                message_id=f"as_msg_{i}", verified=True,
                after_state={"folder": "Deleted Items", "permanent": False},
                backend="applescript", elapsed_ms=5,
            )
            for i in (5, 7, 9)
        ]
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete_bulk.return_value = bulk_results
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply",
            ])
        # delete_bulk called ONCE with every message id.
        client.delete_bulk.assert_called_once()
        args, kwargs = client.delete_bulk.call_args
        assert args[0] == ["as_msg_5", "as_msg_7", "as_msg_9"]
        assert kwargs["permanent"] is False
        assert kwargs["folder_name"] == "Inbox"
        assert rc == 0
        # Per-message delete() must NOT have been called.
        client.delete.assert_not_called()

    def test_bulk_routes_move_buckets_through_move_bulk(self, tmp_path):
        msgs = self._three_msgs()
        for m in msgs:
            m["bucket"] = "resolved-ticket"
        report = self._write_fresh_report(tmp_path, messages=msgs)
        bulk_results = [
            OperationResult(
                status="success", operation="move",
                message_id=f"as_msg_{i}", verified=True,
                after_state={"folder": "Archive"},
                backend="applescript", elapsed_ms=5,
            )
            for i in (5, 7, 9)
        ]
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.move_bulk.return_value = bulk_results
            rc = apply_triage.main([
                str(report), "--bucket", "resolved-ticket", "--apply",
            ])
        client.move_bulk.assert_called_once()
        args, kwargs = client.move_bulk.call_args
        assert args[0] == [
            ("as_msg_5", "Archive"),
            ("as_msg_7", "Archive"),
            ("as_msg_9", "Archive"),
        ]
        assert kwargs["source_folder"] == "Inbox"
        assert rc == 0
        client.move.assert_not_called()

    def test_bulk_mixed_results_surface_as_non_zero(self, tmp_path):
        """Partial bulk → some rows fail → exit code 1 (per policy)."""
        report = self._write_fresh_report(tmp_path, messages=self._three_msgs())
        mixed = [
            OperationResult(status="success", operation="delete",
                            message_id="as_msg_5", verified=True,
                            backend="applescript"),
            OperationResult(status="failed", operation="delete",
                            message_id="as_msg_7", verified=False,
                            error_detail="bulk partial", backend="applescript"),
            OperationResult(status="success", operation="delete",
                            message_id="as_msg_9", verified=False,
                            error_detail="count-only", backend="applescript"),
        ]
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete_bulk.return_value = mixed
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply",
            ])
        assert rc == 1  # failed OR unverified-success → 1

    def test_bulk_exception_tallies_everyone_as_skipped(self, tmp_path):
        report = self._write_fresh_report(tmp_path, messages=self._three_msgs())
        with patch("apply_triage.OutlookClient") as mock_cls:
            client = mock_cls.return_value
            client.backend = "applescript"
            client.delete_bulk.side_effect = RuntimeError("osascript crashed")
            rc = apply_triage.main([
                str(report), "--bucket", "safe-delete", "--apply",
            ])
        assert rc == 1


# ===========================================================================
# Bolt 6 Phase B.3 — apply_triage.apply_unread_junk_gate re-route
# ===========================================================================


class TestUnreadJunkGate:
    """``apply_unread_junk_gate`` upgrades unread + high-confidence-junk
    messages from the archivable buckets (resolved-ticket /
    low-signal-client) or ``keep`` into ``safe-delete``. Mirrors Chris's
    'unread + junk = delete' heuristic (Bolt 6 Phase B.3)."""

    def _msg_rec(
        self,
        *,
        message_id: str = "as_msg_1",
        bucket: str = "keep",
        is_read: bool = False,
        llm_bucket: str | None = None,
        llm_confidence: float | None = None,
        matched_rule: str = "default",
    ) -> dict:
        rec = {
            "message_id": message_id,
            "sender_email": "x@y.com",
            "subject": "s",
            "received_at": datetime.now().isoformat(),
            "bucket": bucket,
            "is_read": is_read,
            "matched_rule": matched_rule,
            "body_preview": "",
        }
        if llm_bucket is not None:
            rec["llm_verdict"] = {
                "bucket": llm_bucket,
                "confidence": llm_confidence if llm_confidence is not None else 0.0,
                "reason": "test",
                "route": "auto",
            }
        return rec

    def _report(self, messages: list) -> dict:
        return {
            "schema_version": 1,
            "generated_at": datetime.now().isoformat(),
            "backend": "applescript",
            "sample_folder": "Inbox",
            "sample_size": len(messages),
            "window_days": 30,
            "summary": {b: 0 for b in triage_rules.BUCKETS},
            "messages": messages,
        }

    def test_keep_unread_high_confidence_junk_upgrades_to_safe_delete(self):
        """Classic shape: rules left it as keep, LLM says junk at 0.9,
        unread → upgrade."""
        rec = self._msg_rec(
            bucket="keep", is_read=False,
            llm_bucket="junk", llm_confidence=0.9,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "safe-delete"
        assert new_report["messages"][0]["matched_rule"] == "unread-junk-upgrade"
        assert new_report["messages"][0]["unread_junk_upgrade"]["from_bucket"] == "keep"
        assert new_report["unread_junk_upgrades"] == 1

    def test_resolved_ticket_unread_junk_upgrades(self):
        """resolved-ticket is in the upgradable set — unread + junk
        should move it to safe-delete."""
        rec = self._msg_rec(
            bucket="resolved-ticket", is_read=False,
            llm_bucket="junk", llm_confidence=0.95,
            matched_rule="hubspot-ticket-comment",
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "safe-delete"
        assert new_report["messages"][0]["unread_junk_upgrade"]["from_rule"] == "hubspot-ticket-comment"

    def test_low_signal_client_unread_junk_upgrades(self):
        """low-signal-client is also in the upgradable set."""
        rec = self._msg_rec(
            bucket="low-signal-client", is_read=False,
            llm_bucket="junk", llm_confidence=0.88,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "safe-delete"

    def test_read_junk_does_not_upgrade(self):
        """Read messages must NOT be upgraded (Chris's rule requires unread)."""
        rec = self._msg_rec(
            bucket="keep", is_read=True,
            llm_bucket="junk", llm_confidence=0.95,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "keep"
        assert new_report["unread_junk_upgrades"] == 0

    def test_low_confidence_junk_does_not_upgrade(self):
        """Confidence < 0.85 → no upgrade."""
        rec = self._msg_rec(
            bucket="keep", is_read=False,
            llm_bucket="junk", llm_confidence=0.70,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "keep"
        assert new_report["unread_junk_upgrades"] == 0

    def test_non_junk_llm_verdict_does_not_upgrade(self):
        """LLM says 'reference' at high confidence → NOT a delete signal."""
        rec = self._msg_rec(
            bucket="keep", is_read=False,
            llm_bucket="reference", llm_confidence=0.95,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "keep"

    def test_needs_action_bucket_protected_from_upgrade(self):
        """needs-action messages must NOT be touched — the point is to
        protect action signal from being swept to safe-delete even when
        the LLM screws up and calls them junk."""
        rec = self._msg_rec(
            bucket="needs-action", is_read=False,
            llm_bucket="junk", llm_confidence=0.95,
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "needs-action"
        assert new_report["unread_junk_upgrades"] == 0

    def test_safe_delete_bucket_unchanged(self):
        """Already-safe-delete stays put — gate only upgrades, never
        downgrades."""
        rec = self._msg_rec(
            bucket="safe-delete", is_read=False,
            llm_bucket="junk", llm_confidence=0.95,
            matched_rule="marketing-sender",
        )
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["matched_rule"] == "marketing-sender"
        assert new_report["unread_junk_upgrades"] == 0

    def test_summary_recounted_after_upgrade(self):
        """When the gate fires, the summary block must reflect the new
        bucket distribution."""
        messages = [
            self._msg_rec(message_id="a", bucket="keep", is_read=False,
                          llm_bucket="junk", llm_confidence=0.9),
            self._msg_rec(message_id="b", bucket="keep", is_read=True,
                          llm_bucket="junk", llm_confidence=0.9),
            self._msg_rec(message_id="c", bucket="needs-action", is_read=False),
        ]
        new_report = apply_triage.apply_unread_junk_gate(self._report(messages))
        assert new_report["summary"]["safe-delete"] == 1
        assert new_report["summary"]["keep"] == 1
        assert new_report["summary"]["needs-action"] == 1

    def test_message_without_llm_verdict_unchanged(self):
        """No llm_verdict attached → gate has nothing to act on; passes through."""
        rec = self._msg_rec(bucket="keep", is_read=False)
        new_report = apply_triage.apply_unread_junk_gate(self._report([rec]))
        assert new_report["messages"][0]["bucket"] == "keep"
        assert "unread_junk_upgrade" not in new_report["messages"][0]


# ===========================================================================
# Bolt 6 Phase B.1 — triage.py wiring (context, body_preview, --llm)
# ===========================================================================


class TestTriageClassifyInboxWiring:
    """Phase B.1 — triage.classify_inbox must thread sent-items context
    into every classify() call and preserve body_preview + is_read on the
    projected report so the LLM gate has what it needs.
    """

    def _client_with_inbox(
        self,
        inbox: list,
        sent: list | None = None,
    ):
        """Mock an OutlookClient. Inbox vs. Sent Items branches on folder kwarg."""
        import triage  # noqa: WPS433 — local import for patching

        sent = sent or []
        client = MagicMock()
        client.backend = "applescript"

        def _search(**kwargs):
            folder = kwargs.get("folder") or "Inbox"
            if folder == "Sent Items":
                return sent
            return inbox

        client.search.side_effect = _search
        return client

    def test_body_preview_and_is_read_in_projection(self):
        """A classified record must include body_preview + is_read so the
        downstream LLM gate (triage_llm) can read them."""
        import triage  # noqa: WPS433

        inbox = [
            {
                "id": "as_msg_1",
                "sender_email": "x@y.com",
                "sender_name": "X",
                "subject": "hello",
                "received_datetime": datetime.now() - timedelta(days=2),
                "is_read": False,
                "body_preview": "Please take a look at this.",
                "folder": "Inbox",
            },
        ]
        client = self._client_with_inbox(inbox)
        with patch("triage.OutlookClient", return_value=client):
            report = triage.classify_inbox(window_days=30, max_results=10)
        assert report["sample_size"] == 1
        rec = report["messages"][0]
        assert rec["body_preview"] == "Please take a look at this."
        assert rec["is_read"] is False

    def test_sent_context_built_once_and_threaded(self):
        """build_sent_context must be called once; behavioral rules
        (ignored-sender-cold) must see the context and fire."""
        import triage  # noqa: WPS433

        # Inbox: 3 messages from a cold sender, > 60 days old, never replied.
        # This is the ignored-sender-cold shape.
        old = datetime.now() - timedelta(days=90)
        inbox = [
            {
                "id": f"as_msg_{i}",
                "sender_email": "rep@unknown-vendor.io",
                "sender_name": "rep",
                "subject": f"ping {i}",
                "received_datetime": old,
                "is_read": True,
                "body_preview": "",
                "folder": "Inbox",
            }
            for i in range(3)
        ]
        client = self._client_with_inbox(inbox, sent=[])
        with patch("triage.OutlookClient", return_value=client):
            report = triage.classify_inbox(window_days=365, max_results=100)

        # At least one record must have matched ignored-sender-cold — the
        # volume floor (sender_counts >= 3) and the never-replied signal
        # (sent_to_counts == {}) both must come from the context.
        matched_rules = [r["matched_rule"] for r in report["messages"]]
        assert "ignored-sender-cold" in matched_rules

        # Sanity: the report's sent_context block records the scan.
        assert "sent_context" in report
        assert report["sent_context"]["days_lookback"] == 365

    def test_sent_context_handles_empty_sent_folder(self):
        """A zero-result Sent Items scan must not crash the pipeline."""
        import triage  # noqa: WPS433

        inbox = [{
            "id": "as_msg_1",
            "sender_email": "x@y.com",
            "sender_name": "X",
            "subject": "hello",
            "received_datetime": datetime.now(),
            "is_read": True,
            "body_preview": "",
            "folder": "Inbox",
        }]
        client = self._client_with_inbox(inbox, sent=[])
        with patch("triage.OutlookClient", return_value=client):
            report = triage.classify_inbox(window_days=30, max_results=10)
        assert report["sample_size"] == 1
        assert report["sent_context"]["scanned_messages"] == 0

    def test_llm_flag_calls_classify_messages_with_llm_only_for_keep(self):
        """--llm must only send keep-bucket messages to the LLM; the gate
        is additive (never overwrites bucket / matched_rule)."""
        import triage  # noqa: WPS433

        # Inbox: 2 marketing msgs (rules → safe-delete) + 1 generic keep msg.
        now = datetime.now()
        inbox = [
            {
                "id": "as_msg_1",
                "sender_email": "hello@birdeye.com",
                "sender_name": "Birdeye",
                "subject": "promo",
                "received_datetime": now,
                "is_read": True,
                "body_preview": "Monthly product update from Birdeye",
                "folder": "Inbox",
            },
            {
                "id": "as_msg_2",
                "sender_email": "person@random.com",
                "sender_name": "Person",
                "subject": "random subject",
                "received_datetime": now,
                "is_read": False,
                "body_preview": "Hey, interesting article you might like",
                "folder": "Inbox",
            },
        ]
        client = self._client_with_inbox(inbox, sent=[])

        def fake_llm(messages, **_kw):
            out = []
            for m in messages:
                new = dict(m)
                new["llm_verdict"] = {
                    "bucket": "junk",
                    "confidence": 0.9,
                    "reason": "mock",
                    "route": "auto",
                }
                out.append(new)
            return out

        with patch("triage.OutlookClient", return_value=client):
            with patch(
                "triage_llm.classify_messages_with_llm",
                side_effect=fake_llm,
            ) as mock_llm:
                report = triage.classify_inbox(
                    window_days=30, max_results=10, use_llm=True,
                    llm_limit=50,
                )

        # LLM must have been called exactly once.
        assert mock_llm.call_count == 1
        # The messages passed must all be bucket=="keep".
        passed_msgs = mock_llm.call_args.args[0]
        assert all(m["bucket"] == "keep" for m in passed_msgs)
        # One keep message in our inbox ⇒ one message sent.
        assert len(passed_msgs) == 1

        # Keep message should now carry llm_verdict; safe-delete must not.
        by_id = {m["message_id"]: m for m in report["messages"]}
        assert by_id["as_msg_2"]["bucket"] == "keep"
        assert by_id["as_msg_2"]["llm_verdict"]["bucket"] == "junk"
        assert "llm_verdict" not in by_id["as_msg_1"]
        # Report-level llm_summary present.
        assert report["llm_summary"]["verdicts_added"] == 1
        assert report["llm_summary"]["keep_candidates"] == 1

    def test_llm_flag_absent_skips_llm_entirely(self):
        """Without --llm, triage_llm must not even be imported-resolved."""
        import triage  # noqa: WPS433

        inbox = [{
            "id": "as_msg_1",
            "sender_email": "x@y.com",
            "sender_name": "X",
            "subject": "hello",
            "received_datetime": datetime.now(),
            "is_read": False,
            "body_preview": "",
            "folder": "Inbox",
        }]
        client = self._client_with_inbox(inbox, sent=[])
        with patch("triage.OutlookClient", return_value=client):
            with patch(
                "triage_llm.classify_messages_with_llm",
            ) as mock_llm:
                report = triage.classify_inbox(window_days=30, max_results=10)
        mock_llm.assert_not_called()
        assert "llm_summary" not in report
