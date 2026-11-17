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
