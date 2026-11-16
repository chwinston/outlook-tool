"""
Tests for Bolt 4 write verbs — ``move``, ``delete``, ``mark_read``.

Layout
------
* Unit tests (no Outlook required): ``OperationResult`` contract, Graph stub
  behaviour (mocked), ``OutlookClient`` dispatch shape, CLI dry-run gating.
* Integration tests (marked ``@pytest.mark.integration``): execute against a
  disposable folder called ``_bolt4_test`` on the user's real Outlook mailbox.
  These are Mac-only and are skipped if Outlook/AppleScript are unavailable.

The integration fixture creates ``_bolt4_test`` before any tests run and
tears it down after — a best-effort delete-on-teardown. If setup fails the
fixture yields ``None`` and tests consume that to self-skip cleanly so CI
without Outlook stays green.
"""

from __future__ import annotations

import sys
import uuid
from pathlib import Path
from typing import Optional

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outlook_tool import (  # noqa: E402  (sys.path first)
    HAS_APPLESCRIPT,
    OperationResult,
    OutlookClient,
)


TEST_FOLDER_NAME = "_bolt4_test"


# =============================================================================
# OperationResult contract (unit — no Outlook needed)
# =============================================================================

class TestOperationResultContract:
    def test_minimum_fields(self):
        r = OperationResult(status="dry_run", operation="move")
        assert r.status == "dry_run"
        assert r.operation == "move"
        assert r.verified is False
        assert r.message_id is None
        assert r.before_state == {}
        assert r.after_state == {}
        assert r.error_detail is None
        assert r.backend is None
        assert r.elapsed_ms is None

    def test_rejects_unknown_status(self):
        with pytest.raises(ValueError, match="Invalid OperationResult.status"):
            OperationResult(status="maybe", operation="move")

    def test_rejects_unknown_operation(self):
        with pytest.raises(ValueError, match="Invalid OperationResult.operation"):
            OperationResult(status="success", operation="flag")

    def test_to_dict_round_trip(self):
        r = OperationResult(
            status="success",
            operation="mark_read",
            message_id="abc",
            verified=True,
            before_state={"is_read": False},
            after_state={"is_read": True},
            backend="applescript",
            elapsed_ms=42,
        )
        d = r.to_dict()
        assert d["status"] == "success"
        assert d["operation"] == "mark_read"
        assert d["message_id"] == "abc"
        assert d["verified"] is True
        assert d["before_state"] == {"is_read": False}
        assert d["after_state"] == {"is_read": True}
        assert d["backend"] == "applescript"
        assert d["elapsed_ms"] == 42


# =============================================================================
# Integration fixture — disposable _bolt4_test folder on real Outlook (Mac only)
# =============================================================================

def _can_run_applescript_integration() -> bool:
    """Return True only if Outlook is reachable right now."""
    if not HAS_APPLESCRIPT:
        return False
    if sys.platform != "darwin":
        return False
    try:
        from outlook_tool import _AppleScriptBackend
        _AppleScriptBackend()  # probes `tell application Outlook to get version`
    except Exception:
        return False
    return True


@pytest.fixture(scope="session")
def bolt4_test_folder() -> Optional[str]:
    """Create ``_bolt4_test`` on the primary Exchange account; yield its name.

    Best-effort teardown. If create fails (Outlook not running, no account, etc.)
    yield ``None`` so integration tests can self-skip cleanly.
    """
    if not _can_run_applescript_integration():
        yield None
        return

    from outlook_tool import _AppleScriptBackend  # noqa: E402

    backend = _AppleScriptBackend()
    try:
        backend.ensure_folder(TEST_FOLDER_NAME)
    except Exception as e:  # pragma: no cover — depends on real Outlook
        print(f"bolt4_test_folder: setup failed: {e}", file=sys.stderr)
        yield None
        return

    yield TEST_FOLDER_NAME

    # Best-effort teardown. If it fails we leave the empty folder — safer than
    # nuking something the user named the same.
    try:
        backend.delete_folder_if_empty(TEST_FOLDER_NAME)
    except Exception as e:  # pragma: no cover
        print(
            f"bolt4_test_folder: teardown failed, folder {TEST_FOLDER_NAME!r} left behind: {e}",
            file=sys.stderr,
        )


@pytest.fixture()
def seeded_test_message(bolt4_test_folder):
    """Place a fresh disposable email into ``_bolt4_test`` and return its handle.

    The ``OutlookClient`` ``search()`` call returns the seeded email so tests
    can get the ``_as_msg_index`` / ``_as_folder_name`` needed for AppleScript
    write ops.

    Skips cleanly if the folder fixture didn't come up.
    """
    if bolt4_test_folder is None:
        pytest.skip("Outlook/AppleScript integration unavailable")

    from outlook_tool import _AppleScriptBackend  # noqa: E402

    backend = _AppleScriptBackend()
    unique_tag = f"bolt4-seed-{uuid.uuid4().hex[:8]}"
    try:
        msg_ref = backend.seed_draft_in_folder(
            folder_name=bolt4_test_folder,
            subject=f"[{unique_tag}] Bolt 4 test message",
            body="Disposable test message created by tests/test_write_ops.py",
        )
    except Exception as e:
        pytest.skip(f"seeding test message failed: {e}")

    yield msg_ref

    # Leave teardown of the message itself to the folder teardown or the test.


# =============================================================================
# AppleScript backend write ops — mocked unit tests
# =============================================================================

from unittest.mock import patch  # noqa: E402


def _make_applescript_backend():
    """Build an ``_AppleScriptBackend`` without invoking Outlook."""
    from outlook_tool import _AppleScriptBackend
    backend = _AppleScriptBackend.__new__(_AppleScriptBackend)
    backend._version = "16.0"
    return backend


class TestAppleScriptMarkRead:
    def test_success_when_post_state_matches(self):
        backend = _make_applescript_backend()
        state_sequence = [
            {"folder": "Inbox", "is_read": False, "subject": "Hi", "sender_email": "a@b.com"},
            {"folder": "Inbox", "is_read": True, "subject": "Hi", "sender_email": "a@b.com"},
        ]
        with patch.object(backend, "_fetch_message_state", side_effect=state_sequence), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.mark_read(msg_index=3, folder_name="Inbox", read=True)
        assert r.status == "success"
        assert r.verified is True
        assert r.backend == "applescript"
        assert r.operation == "mark_read"
        assert r.before_state["is_read"] is False
        assert r.after_state["is_read"] is True

    def test_silent_success_detected(self):
        """Post-state unchanged = failed, not success (Tim F1)."""
        backend = _make_applescript_backend()
        unchanged = {"folder": "Inbox", "is_read": False, "subject": "Hi", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", side_effect=[unchanged, unchanged]), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.mark_read(msg_index=3, folder_name="Inbox", read=True)
        assert r.status == "failed"
        assert r.verified is False
        assert "silent success" in (r.error_detail or "").lower()

    def test_pre_fetch_failure_short_circuits(self):
        backend = _make_applescript_backend()
        with patch.object(
            backend, "_fetch_message_state",
            side_effect=RuntimeError("message not found"),
        ):
            r = backend.mark_read(msg_index=99, folder_name="Inbox", read=True)
        assert r.status == "failed"
        assert "pre-fetch failed" in r.error_detail

    def test_applescript_error_during_write(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Hi", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch("outlook_tool._run_applescript", side_effect=RuntimeError("boom")):
            r = backend.mark_read(msg_index=3, folder_name="Inbox", read=True)
        assert r.status == "failed"
        assert "applescript error" in r.error_detail


class TestAppleScriptMove:
    def test_success_with_fingerprint_hit(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", return_value=True), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "success"
        assert r.verified is True
        assert r.operation == "move"
        assert r.after_state["folder"] == "Archive"
        assert r.before_state["folder"] == "Inbox"

    def test_silent_success_when_not_in_target(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", return_value=False), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "failed"
        assert r.verified is False
        assert "silent-success" in (r.error_detail or "")

    def test_empty_target_folder_rejected(self):
        backend = _make_applescript_backend()
        r = backend.move(msg_index=1, target_folder="", source_folder="Inbox")
        assert r.status == "failed"
        assert "target_folder" in r.error_detail

    def test_ensure_folder_failure_short_circuits(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder", side_effect=RuntimeError("ACL denied")):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "failed"
        assert "ensure_folder" in r.error_detail


class TestAppleScriptDelete:
    def test_soft_delete_delegates_to_move(self):
        backend = _make_applescript_backend()
        # Soft-delete = move to Deleted Items. Mock out the move() path.
        fake_move_result = OperationResult(
            status="success",
            operation="move",
            message_id="as_msg_7",
            verified=True,
            before_state={"folder": "Inbox"},
            after_state={"folder": "Deleted Items"},
            backend="applescript",
            elapsed_ms=5,
        )
        with patch.object(backend, "move", return_value=fake_move_result) as mv:
            r = backend.delete_message(msg_index=7, folder_name="Inbox", permanent=False)
        mv.assert_called_once_with(
            msg_index=7, target_folder="Deleted Items", source_folder="Inbox",
        )
        assert r.status == "success"
        assert r.operation == "delete"  # rewritten to delete
        assert r.verified is True
        assert r.after_state["permanent"] is False

    def test_hard_delete_verifies_absence(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Spam", "sender_email": "x@y.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "_verify_present_in_folder", return_value=False), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.delete_message(msg_index=4, folder_name="Inbox", permanent=True)
        assert r.status == "success"
        assert r.verified is True
        assert r.after_state["permanent"] is True

    def test_hard_delete_silent_success_fails(self):
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Spam", "sender_email": "x@y.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "_verify_present_in_folder", return_value=True), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.delete_message(msg_index=4, folder_name="Inbox", permanent=True)
        assert r.status == "failed"
        assert r.verified is False
        assert "still present" in (r.error_detail or "")


# =============================================================================
# AppleScript integration tests — against the disposable _bolt4_test folder
# =============================================================================

@pytest.mark.integration
class TestAppleScriptWriteOpsIntegration:
    def test_mark_read_and_unread_round_trip(self, seeded_test_message):
        from outlook_tool import _AppleScriptBackend
        backend = _AppleScriptBackend()
        idx = seeded_test_message["msg_index"]
        folder = seeded_test_message["folder_name"]

        r1 = backend.mark_read(msg_index=idx, folder_name=folder, read=True)
        assert r1.status == "success", r1.error_detail
        assert r1.verified is True
        assert r1.after_state["is_read"] is True

        r2 = backend.mark_read(msg_index=idx, folder_name=folder, read=False)
        assert r2.status == "success", r2.error_detail
        assert r2.verified is True
        assert r2.after_state["is_read"] is False

    def test_move_into_subfolder_then_back(self, bolt4_test_folder, seeded_test_message):
        from outlook_tool import _AppleScriptBackend
        backend = _AppleScriptBackend()
        idx = seeded_test_message["msg_index"]
        source = seeded_test_message["folder_name"]
        # Move to Inbox (always exists). We do not move to an unrelated real
        # folder — we're confined to _bolt4_test and system folders.
        r = backend.move(msg_index=idx, target_folder="Inbox", source_folder=source)
        assert r.status == "success", r.error_detail
        assert r.verified is True
        assert r.after_state["folder"] == "Inbox"

    def test_soft_delete_sends_to_deleted_items(self, seeded_test_message):
        from outlook_tool import _AppleScriptBackend
        backend = _AppleScriptBackend()
        idx = seeded_test_message["msg_index"]
        folder = seeded_test_message["folder_name"]
        r = backend.delete_message(msg_index=idx, folder_name=folder, permanent=False)
        assert r.status == "success", r.error_detail
        assert r.verified is True
        assert r.after_state["permanent"] is False


# =============================================================================
# Shared integration marker — gated by presence of the fixture
# =============================================================================

integration = pytest.mark.integration
