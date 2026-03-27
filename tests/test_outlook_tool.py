"""
Tests for outlook_tool.py — unit tests for helpers, filters, and client logic.

These tests don't require Outlook or Graph API access. They test the pure-Python
logic: date parsing, domain extraction, post-filtering, and client initialization.
"""

import re
import unittest
from datetime import datetime
from unittest.mock import MagicMock, patch

import pytest

# We need to be able to import even on Mac where win32com isn't available
import sys
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent.parent))

from outlook_tool import OutlookClient, _parse_date, _extract_domain, _escape_applescript


# =============================================================================
# HELPER TESTS
# =============================================================================

class TestEscapeApplescript:
    def test_double_quotes_escaped(self):
        assert _escape_applescript('say "hello"') == 'say \\"hello\\"'

    def test_backslash_escaped(self):
        assert _escape_applescript("path\\to\\file") == "path\\\\to\\\\file"

    def test_newlines_escaped(self):
        assert _escape_applescript("line1\nline2") == "line1\\nline2"

    def test_carriage_return_escaped(self):
        assert _escape_applescript("a\rb") == "a\\rb"

    def test_tab_escaped(self):
        assert _escape_applescript("col1\tcol2") == "col1\\tcol2"

    def test_combined_special_chars(self):
        result = _escape_applescript('He said "hi"\npath\\x')
        assert result == 'He said \\"hi\\"\\npath\\\\x'

    def test_plain_string_unchanged(self):
        assert _escape_applescript("hello world") == "hello world"

    def test_empty_string(self):
        assert _escape_applescript("") == ""

    def test_injection_attempt(self):
        # Simulates an attacker trying to break out of a quoted string
        malicious = '" & do shell script "rm -rf /" & "'
        result = _escape_applescript(malicious)
        assert '"' not in result.replace('\\"', '')  # no unescaped quotes


class TestParseDate:
    def test_iso_format(self):
        assert _parse_date("2026-03-15") == datetime(2026, 3, 15)

    def test_us_format(self):
        assert _parse_date("03/15/2026") == datetime(2026, 3, 15)

    def test_slash_iso_format(self):
        assert _parse_date("2026/03/15") == datetime(2026, 3, 15)

    def test_datetime_passthrough(self):
        dt = datetime(2026, 3, 15, 10, 30)
        assert _parse_date(dt) is dt

    def test_invalid_format_raises(self):
        with pytest.raises(ValueError, match="Cannot parse date"):
            _parse_date("March 15, 2026")

    def test_empty_string_raises(self):
        with pytest.raises(ValueError):
            _parse_date("")


class TestExtractDomain:
    def test_standard_email(self):
        assert _extract_domain("user@example.com") == "example.com"

    def test_case_insensitive(self):
        assert _extract_domain("User@EXAMPLE.COM") == "example.com"

    def test_no_at_sign(self):
        assert _extract_domain("not-an-email") == ""

    def test_empty_string(self):
        assert _extract_domain("") == ""

    def test_multiple_at_signs(self):
        assert _extract_domain("weird@nested@domain.com") == "domain.com"


# =============================================================================
# POST-FILTER TESTS
# =============================================================================

def _make_email(**kwargs):
    """Create a minimal email dict for testing filters."""
    defaults = {
        "id": "test_1",
        "subject": "Test Email",
        "sender_name": "John Doe",
        "sender_email": "john@example.com",
        "received_datetime": datetime(2026, 3, 15),
        "received_date": "2026-03-15",
        "day_of_week": "Sunday",
        "is_read": False,
        "has_attachments": False,
        "importance": "normal",
        "body_preview": "Hello, this is a test email body.",
        "to": "recipient@example.com",
        "attachments": [],
    }
    defaults.update(kwargs)
    return defaults


class TestPostFilters:
    def test_subject_contains_match(self):
        emails = [_make_email(subject="Q1 Report 2026"), _make_email(subject="Hello")]
        result = OutlookClient._apply_post_filters(emails, subject_contains="report")
        assert len(result) == 1
        assert result[0]["subject"] == "Q1 Report 2026"

    def test_subject_contains_case_insensitive(self):
        emails = [_make_email(subject="URGENT Report")]
        result = OutlookClient._apply_post_filters(emails, subject_contains="urgent")
        assert len(result) == 1

    def test_subject_matches_regex(self):
        emails = [
            _make_email(subject="Q1 2026 Report"),
            _make_email(subject="Q5 2026 Report"),
            _make_email(subject="Meeting Notes"),
        ]
        result = OutlookClient._apply_post_filters(emails, subject_matches=r"Q[1-4] \d{4}")
        assert len(result) == 1
        assert result[0]["subject"] == "Q1 2026 Report"

    def test_sender_name_filter(self):
        emails = [
            _make_email(sender_name="Jane Smith"),
            _make_email(sender_name="John Doe"),
        ]
        result = OutlookClient._apply_post_filters(emails, sender_name="jane")
        assert len(result) == 1

    def test_sender_email_exact_match(self):
        emails = [
            _make_email(sender_email="jane@example.com"),
            _make_email(sender_email="john@example.com"),
        ]
        result = OutlookClient._apply_post_filters(emails, sender_email="jane@example.com")
        assert len(result) == 1

    def test_sender_email_case_insensitive(self):
        emails = [_make_email(sender_email="Jane@Example.COM")]
        result = OutlookClient._apply_post_filters(emails, sender_email="jane@example.com")
        assert len(result) == 1

    def test_sender_domain_filter(self):
        emails = [
            _make_email(sender_email="a@example.com"),
            _make_email(sender_email="b@other.com"),
        ]
        result = OutlookClient._apply_post_filters(emails, sender_domains={"example.com"})
        assert len(result) == 1

    def test_sender_domains_multiple(self):
        emails = [
            _make_email(sender_email="a@example.com"),
            _make_email(sender_email="b@partner.org"),
            _make_email(sender_email="c@unknown.net"),
        ]
        result = OutlookClient._apply_post_filters(
            emails, sender_domains={"example.com", "partner.org"}
        )
        assert len(result) == 2

    def test_body_contains(self):
        emails = [
            _make_email(body_preview="Please review the attached report."),
            _make_email(body_preview="Meeting at 3pm tomorrow."),
        ]
        result = OutlookClient._apply_post_filters(emails, body_contains="report")
        assert len(result) == 1

    def test_to_contains(self):
        emails = [
            _make_email(to="team@example.com"),
            _make_email(to="personal@me.com"),
        ]
        result = OutlookClient._apply_post_filters(emails, to_contains="team@")
        assert len(result) == 1

    def test_combined_filters(self):
        emails = [
            _make_email(subject="Q1 Report", sender_email="jane@example.com"),
            _make_email(subject="Q1 Report", sender_email="jane@other.com"),
            _make_email(subject="Hello", sender_email="jane@example.com"),
        ]
        result = OutlookClient._apply_post_filters(
            emails,
            subject_contains="report",
            sender_domains={"example.com"},
        )
        assert len(result) == 1
        assert result[0]["sender_email"] == "jane@example.com"

    def test_no_filters_returns_all(self):
        emails = [_make_email(), _make_email(), _make_email()]
        result = OutlookClient._apply_post_filters(emails)
        assert len(result) == 3

    def test_empty_input_returns_empty(self):
        result = OutlookClient._apply_post_filters([], subject_contains="anything")
        assert result == []


# =============================================================================
# CLIENT INITIALIZATION TESTS
# =============================================================================

class TestClientInit:
    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", False)
    @patch("outlook_tool.HAS_GRAPH", False)
    def test_no_backend_raises(self):
        with pytest.raises(RuntimeError, match="No email backend available"):
            OutlookClient()

    @patch("outlook_tool.HAS_WIN32", True)
    def test_win32_backend_selected(self):
        client = OutlookClient()
        assert client.backend == "win32com"

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool._AppleScriptBackend")
    def test_applescript_backend_selected(self, mock_as_cls):
        mock_as_cls.return_value = MagicMock()
        client = OutlookClient()
        assert client.backend == "applescript"

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool.HAS_GRAPH", True)
    @patch("outlook_tool._AppleScriptBackend")
    def test_applescript_preferred_over_graph(self, mock_as_cls):
        """AppleScript should win over Graph API on Mac."""
        mock_as_cls.return_value = MagicMock()
        client = OutlookClient()
        assert client.backend == "applescript"

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", False)
    @patch("outlook_tool.HAS_GRAPH", True)
    @patch("outlook_tool.msal", create=True)
    def test_graph_backend_selected(self, mock_msal):
        mock_msal.SerializableTokenCache.return_value = MagicMock(
            has_state_changed=False
        )
        mock_msal.PublicClientApplication.return_value = MagicMock()
        client = OutlookClient()
        assert client.backend == "graph"

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool.HAS_GRAPH", True)
    @patch("outlook_tool.msal", create=True)
    def test_force_graph_backend(self, mock_msal):
        """backend= kwarg should override auto-detection."""
        mock_msal.SerializableTokenCache.return_value = MagicMock(
            has_state_changed=False
        )
        mock_msal.PublicClientApplication.return_value = MagicMock()
        client = OutlookClient(backend="graph")
        assert client.backend == "graph"


# =============================================================================
# SEARCH PARAMETER VALIDATION
# =============================================================================

class TestSearchValidation:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_date_string_parsed(self):
        """Verify date strings get parsed before hitting the backend."""
        client = OutlookClient()
        # Mock the win32 search to capture what it receives
        with patch.object(client, "_search_win32", return_value=[]) as mock_search:
            client.search(date_from="2026-03-01", date_to="2026-03-15")
            call_kwargs = mock_search.call_args
            assert call_kwargs[1]["date_from"] == datetime(2026, 3, 1)
            assert call_kwargs[1]["date_to"] == datetime(2026, 3, 15)

    @patch("outlook_tool.HAS_WIN32", True)
    def test_domain_merging(self):
        """Verify sender_domain and sender_domains merge correctly."""
        client = OutlookClient()
        with patch.object(client, "_search_win32", return_value=[
            _make_email(sender_email="a@example.com"),
            _make_email(sender_email="b@partner.org"),
            _make_email(sender_email="c@other.net"),
        ]):
            results = client.search(
                sender_domain="example.com",
                sender_domains=["partner.org"],
            )
            assert len(results) == 2


# =============================================================================
# SEND VALIDATION
# =============================================================================

class TestSendValidation:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_missing_attachment_raises(self):
        client = OutlookClient()
        with pytest.raises(FileNotFoundError, match="Attachment not found"):
            client.send(
                to="test@example.com",
                subject="Test",
                body="Body",
                attachments=["/nonexistent/file.pdf"],
            )

    @patch("outlook_tool.HAS_WIN32", True)
    def test_to_string_converted_to_list(self):
        client = OutlookClient()
        with patch.object(client, "_send_win32", return_value=True) as mock_send:
            client.send(to="single@example.com", subject="Test", body="Body")
            call_args = mock_send.call_args[0]
            assert call_args[0] == ["single@example.com"]  # to is now a list


# ============================================================================
# GET EVENTS TESTS
# ============================================================================


class TestGetEvents(unittest.TestCase):
    """Tests for the calendar events API."""

    SAMPLE_EVENTS = [
        {
            "id": "evt_1",
            "subject": "Team Standup",
            "start_datetime": datetime(2026, 3, 27, 9, 0),
            "end_datetime": datetime(2026, 3, 27, 9, 30),
            "start_date": "2026-03-27",
            "end_date": "2026-03-27",
            "location": "Conference Room A",
            "organizer_name": "Jane Smith",
            "organizer_email": "jane@example.com",
            "is_all_day": False,
            "status": "busy",
            "body_preview": "Daily standup meeting.",
            "attendees": [
                {"name": "Bob", "email": "bob@example.com", "status": "accepted"},
            ],
        },
        {
            "id": "evt_2",
            "subject": "Quarterly Review",
            "start_datetime": datetime(2026, 3, 27, 14, 0),
            "end_datetime": datetime(2026, 3, 27, 15, 0),
            "start_date": "2026-03-27",
            "end_date": "2026-03-27",
            "location": "",
            "organizer_name": "CEO",
            "organizer_email": "ceo@example.com",
            "is_all_day": False,
            "status": "tentative",
            "body_preview": "Q1 review.",
            "attendees": [],
        },
        {
            "id": "evt_3",
            "subject": "Company Holiday",
            "start_datetime": datetime(2026, 3, 28, 0, 0),
            "end_datetime": datetime(2026, 3, 28, 23, 59),
            "start_date": "2026-03-28",
            "end_date": "2026-03-28",
            "location": "",
            "organizer_name": "",
            "organizer_email": "",
            "is_all_day": True,
            "status": "free",
            "body_preview": "",
            "attendees": [],
        },
    ]

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_delegates_to_backend(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=self.SAMPLE_EVENTS
        ) as mock:
            results = client.get_events(date_from="2026-03-27", date_to="2026-03-28")
            mock.assert_called_once()
            assert len(results) == 3

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_subject_filter(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=self.SAMPLE_EVENTS
        ):
            results = client.get_events(
                date_from="2026-03-27", subject_contains="standup"
            )
            assert len(results) == 1
            assert results[0]["subject"] == "Team Standup"

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_subject_filter_case_insensitive(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=self.SAMPLE_EVENTS
        ):
            results = client.get_events(
                date_from="2026-03-27", subject_contains="QUARTERLY"
            )
            assert len(results) == 1
            assert results[0]["subject"] == "Quarterly Review"

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_max_results(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=self.SAMPLE_EVENTS
        ):
            results = client.get_events(date_from="2026-03-27", max_results=2)
            assert len(results) == 2

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_defaults_to_today(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=[]
        ) as mock:
            client.get_events()
            call_args = mock.call_args[0]
            # date_from should be today at midnight
            assert call_args[0].date() == datetime.now().date()

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_default_end_is_7_days(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=[]
        ) as mock:
            client.get_events(date_from="2026-03-27")
            call_args = mock.call_args[0]
            # date_to should be 7 days after date_from
            assert call_args[1] == datetime(2026, 4, 3)

    @patch("outlook_tool.HAS_WIN32", True)
    def test_get_events_event_dict_structure(self):
        client = OutlookClient()
        with patch.object(
            client, "_get_events_win32", return_value=self.SAMPLE_EVENTS
        ):
            results = client.get_events(date_from="2026-03-27")
            evt = results[0]
            # Check all required fields exist
            required_fields = [
                "id", "subject", "start_datetime", "end_datetime",
                "start_date", "end_date", "location", "organizer_name",
                "organizer_email", "is_all_day", "status", "body_preview",
                "attendees",
            ]
            for field in required_fields:
                assert field in evt, f"Missing field: {field}"
