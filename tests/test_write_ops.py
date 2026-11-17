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
    def test_success_with_both_source_absence_and_target_presence(self):
        # Verify is called twice: source-check first (returns False = not
        # present = move succeeded), then target-check (returns True = found
        # in Archive = fully verified).
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return False  # not in source → move succeeded
            if folder_name == "Archive":
                return True  # confirmed in target
            raise AssertionError(f"unexpected folder {folder_name!r}")

        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "success"
        assert r.verified is True
        assert r.operation == "move"
        assert r.after_state["folder"] == "Archive"
        assert r.before_state["folder"] == "Inbox"

    def test_unverified_success_when_target_scan_misses(self):
        # Large-mailbox case: move succeeded (gone from source) but target-side
        # fingerprint scan misses (e.g. Deleted Items has 5000+ messages).
        # Should return success, verified=False — the UNVERIFIED-success tier.
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return False  # not in source
            return False  # scan missed in target

        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "success"
        assert r.verified is False
        assert "source-absence confirms" in (r.error_detail or "")

    def test_silent_success_when_message_still_in_source(self):
        # Real silent-success: AppleScript returned OK but message is still
        # in source folder → the move didn't take effect.
        backend = _make_applescript_backend()
        before = {"folder": "Inbox", "is_read": False, "subject": "Report", "sender_email": "jane@e.com"}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return True  # still there → move failed
            return False

        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "failed"
        assert r.verified is False
        assert "silent-success" in (r.error_detail or "")
        assert "still in" in (r.error_detail or "")

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
# COM backend write ops — mocked unit tests (win32com unavailable on Mac)
# =============================================================================

from unittest.mock import MagicMock  # noqa: E402


class TestCOMMarkRead:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_mark_read_success(self):
        client = OutlookClient()
        msg_before = MagicMock(Subject="Hi", SenderEmailAddress="a@b.com",
                               UnRead=True, EntryID="eid1")
        msg_before.Parent.Name = "Inbox"
        msg_after = MagicMock(Subject="Hi", SenderEmailAddress="a@b.com",
                              UnRead=False, EntryID="eid1")
        msg_after.Parent.Name = "Inbox"
        with patch.object(
            client, "_get_win32_message_by_entry_id",
            side_effect=[msg_before, msg_after],
        ):
            r = client._mark_read_win32("eid1", read=True)
        assert r.status == "success"
        assert r.verified is True
        assert r.backend == "win32com"
        assert r.before_state["is_read"] is False
        assert r.after_state["is_read"] is True

    @patch("outlook_tool.HAS_WIN32", True)
    def test_mark_read_silent_success_detected(self):
        """Simulate Outlook silently refusing the write — before and after both
        report unread even though we wrote UnRead=False."""
        client = OutlookClient()
        msg_before = MagicMock(Subject="Hi", SenderEmailAddress="a@b.com",
                               UnRead=True, EntryID="eid1")
        msg_before.Parent.Name = "Inbox"
        # The post-write re-fetch returns a *fresh* mock whose UnRead is still True.
        msg_after = MagicMock(Subject="Hi", SenderEmailAddress="a@b.com",
                              UnRead=True, EntryID="eid1")
        msg_after.Parent.Name = "Inbox"
        with patch.object(
            client, "_get_win32_message_by_entry_id",
            side_effect=[msg_before, msg_after],
        ):
            r = client._mark_read_win32("eid1", read=True)
        assert r.status == "failed"
        assert r.verified is False


class TestCOMMove:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_move_captures_returned_item(self):
        """COM Move() invalidates the original handle — we must use what it returns."""
        client = OutlookClient()
        original = MagicMock(Subject="S", SenderEmailAddress="s@x.com",
                             UnRead=False, EntryID="old_eid")
        original.Parent.Name = "Inbox"
        moved = MagicMock(Subject="S", SenderEmailAddress="s@x.com",
                          UnRead=False, EntryID="new_eid")
        moved.Parent.Name = "Archive"
        original.Move.return_value = moved
        target_folder = MagicMock(Name="Archive")

        with patch.object(client, "_get_win32_message_by_entry_id",
                          return_value=original), \
             patch.object(client, "_resolve_win32_folder",
                          return_value=target_folder):
            r = client._move_win32("old_eid", "Archive")
        original.Move.assert_called_once_with(target_folder)
        assert r.status == "success"
        assert r.verified is True
        # EntryID is rewritten to the moved item's.
        assert r.message_id == "new_eid"
        assert r.after_state["folder"] == "Archive"

    @patch("outlook_tool.HAS_WIN32", True)
    def test_move_returns_none_is_failure(self):
        client = OutlookClient()
        original = MagicMock(Subject="S", SenderEmailAddress="s@x.com",
                             UnRead=False, EntryID="eid")
        original.Parent.Name = "Inbox"
        original.Move.return_value = None
        target_folder = MagicMock(Name="Archive")
        with patch.object(client, "_get_win32_message_by_entry_id",
                          return_value=original), \
             patch.object(client, "_resolve_win32_folder",
                          return_value=target_folder):
            r = client._move_win32("eid", "Archive")
        assert r.status == "failed"
        assert "returned None" in r.error_detail

    @patch("outlook_tool.HAS_WIN32", True)
    def test_move_empty_folder_rejected(self):
        client = OutlookClient()
        r = client._move_win32("eid", "")
        assert r.status == "failed"
        assert "target_folder" in r.error_detail


class TestCOMDelete:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_soft_delete_goes_through_move(self):
        client = OutlookClient()
        with patch.object(
            client, "_move_win32",
            return_value=OperationResult(
                status="success", operation="move", message_id="new_eid",
                verified=True,
                before_state={"folder": "Inbox"},
                after_state={"folder": "Deleted Items"},
                backend="win32com", elapsed_ms=3,
            ),
        ) as mv:
            r = client._delete_win32("old_eid", permanent=False)
        mv.assert_called_once_with("old_eid", "Deleted Items")
        assert r.status == "success"
        assert r.operation == "delete"
        assert r.after_state["permanent"] is False

    @patch("outlook_tool.HAS_WIN32", True)
    def test_hard_delete_verifies_entry_id_gone(self):
        client = OutlookClient()
        msg = MagicMock(Subject="X", SenderEmailAddress="x@y.com",
                        UnRead=False, EntryID="eid")
        msg.Parent.Name = "Inbox"
        with patch.object(
            client, "_get_win32_message_by_entry_id",
            side_effect=[msg, Exception("not found")],
        ):
            r = client._delete_win32("eid", permanent=True)
        msg.Delete.assert_called_once()
        assert r.status == "success"
        assert r.verified is True
        assert r.after_state["permanent"] is True

    @patch("outlook_tool.HAS_WIN32", True)
    def test_hard_delete_silent_success_when_entry_still_resolves(self):
        client = OutlookClient()
        msg = MagicMock(Subject="X", SenderEmailAddress="x@y.com",
                        UnRead=False, EntryID="eid")
        msg.Parent.Name = "Inbox"
        with patch.object(
            client, "_get_win32_message_by_entry_id",
            side_effect=[msg, msg],  # 2nd call still succeeds = silent-success
        ):
            r = client._delete_win32("eid", permanent=True)
        assert r.status == "failed"
        assert r.verified is False


# =============================================================================
# Graph backend write ops — scope-gated, mocked HTTP
# =============================================================================

def _make_graph_client(scopes):
    """Build an OutlookClient with a stubbed _GraphBackend so no network fires."""
    client = OutlookClient.__new__(OutlookClient)
    client.backend = "graph"
    client._graph = MagicMock()
    client._graph.scopes = scopes
    # Delegate has_write_scope to the real impl for realism.
    from outlook_tool import _GraphBackend
    client._graph.has_write_scope = lambda: _GraphBackend.has_write_scope(client._graph)
    client._applescript = None
    client._win32_msg_cache = {}
    return client


class TestGraphScopeDetection:
    def test_read_only_scopes_cannot_write(self):
        c = _make_graph_client(["Mail.Read"])
        r = c._mark_read_graph("msgid", read=True)
        assert r.status == "unsupported"
        assert "Mail.ReadWrite" in r.error_detail
        assert r.backend == "graph"

    def test_mail_send_alone_still_cannot_write(self):
        c = _make_graph_client(["Mail.Read", "Mail.Send"])
        r = c._move_graph("msgid", "Archive")
        assert r.status == "unsupported"

    def test_mail_readwrite_granted_proceeds(self):
        c = _make_graph_client(["Mail.ReadWrite"])
        fake_resp = MagicMock(status_code=200)
        c._graph._api_request = MagicMock(return_value=fake_resp)
        r = c._mark_read_graph("msgid", read=True)
        assert r.status == "success"
        assert r.backend == "graph"
        assert r.verified is True

    def test_mail_readwrite_all_also_granted(self):
        c = _make_graph_client(["Mail.ReadWrite.All"])
        fake_resp = MagicMock(status_code=200)
        c._graph._api_request = MagicMock(return_value=fake_resp)
        r = c._mark_read_graph("msgid", read=True)
        assert r.status == "success"


class TestGraph403VsScopeMissing:
    """Panel F2 — scope-missing and 403 permission-denied must not be conflated."""

    def test_403_with_write_scope_is_failed_not_unsupported(self):
        c = _make_graph_client(["Mail.ReadWrite"])

        class FakeResp:
            status_code = 403
            text = '{"error":{"code":"ErrorAccessDenied"}}'

        class HTTPErrorLike(Exception):
            def __init__(self, resp):
                self.response = resp

        c._graph._api_request = MagicMock(
            side_effect=HTTPErrorLike(FakeResp()),
        )
        r = c._mark_read_graph("msgid", read=True)
        assert r.status == "failed"  # NOT unsupported
        assert "403" in r.error_detail
        assert "permission denied" in r.error_detail.lower()

    def test_401_surfaces_as_failed(self):
        c = _make_graph_client(["Mail.ReadWrite"])

        class FakeResp:
            status_code = 401
            text = "token expired"

        class HTTPErrorLike(Exception):
            def __init__(self, resp):
                self.response = resp

        c._graph._api_request = MagicMock(
            side_effect=HTTPErrorLike(FakeResp()),
        )
        r = c._move_graph("msgid", "Archive")
        assert r.status == "failed"
        assert "401" in r.error_detail


class TestGraphMove:
    def test_move_calls_graph_move_endpoint(self):
        c = _make_graph_client(["Mail.ReadWrite"])
        fake_resp = MagicMock(status_code=201)
        c._graph._api_request = MagicMock(return_value=fake_resp)
        r = c._move_graph("mid", "archive")
        assert r.status == "success"
        call_args = c._graph._api_request.call_args
        assert call_args[0][0] == "post"
        assert "/messages/mid/move" in call_args[0][1]
        assert call_args[1]["json"] == {"destinationId": "archive"}

    def test_move_empty_target_rejected(self):
        c = _make_graph_client(["Mail.ReadWrite"])
        r = c._move_graph("mid", "")
        assert r.status == "failed"
        assert "target_folder" in r.error_detail


class TestGraphDelete:
    def test_soft_delete_uses_move_endpoint_with_deleteditems(self):
        c = _make_graph_client(["Mail.ReadWrite"])
        fake_resp = MagicMock(status_code=201)
        c._graph._api_request = MagicMock(return_value=fake_resp)
        r = c._delete_graph("mid", permanent=False)
        assert r.status == "success"
        call = c._graph._api_request.call_args
        assert "/messages/mid/move" in call[0][1]
        assert call[1]["json"] == {"destinationId": "deleteditems"}
        assert r.after_state["permanent"] is False

    def test_hard_delete_uses_delete_verb(self):
        c = _make_graph_client(["Mail.ReadWrite"])
        fake_resp = MagicMock(status_code=204)
        c._graph._api_request = MagicMock(return_value=fake_resp)
        r = c._delete_graph("mid", permanent=True)
        assert r.status == "success"
        call = c._graph._api_request.call_args
        assert call[0][0] == "delete"
        assert "/messages/mid" in call[0][1] and "/move" not in call[0][1]
        assert r.after_state["permanent"] is True


# =============================================================================
# OutlookClient dispatch tests (yes_really_permanent, backend routing)
# =============================================================================

class TestOutlookClientDispatch:
    @patch("outlook_tool.HAS_WIN32", True)
    def test_mark_read_routes_to_win32(self):
        c = OutlookClient()
        with patch.object(
            c, "_mark_read_win32",
            return_value=OperationResult(status="success", operation="mark_read",
                                         message_id="eid", backend="win32com"),
        ) as m:
            r = c.mark_read("eid", read=False)
        m.assert_called_once_with("eid", read=False)
        assert r.status == "success"

    @patch("outlook_tool.HAS_WIN32", True)
    def test_move_routes_to_win32(self):
        c = OutlookClient()
        with patch.object(
            c, "_move_win32",
            return_value=OperationResult(status="success", operation="move",
                                         message_id="new_eid", backend="win32com"),
        ) as m:
            c.move("eid", "Archive")
        m.assert_called_once_with("eid", "Archive")

    @patch("outlook_tool.HAS_WIN32", True)
    def test_delete_default_is_soft(self):
        c = OutlookClient()
        with patch.object(
            c, "_delete_win32",
            return_value=OperationResult(status="success", operation="delete",
                                         message_id="eid", backend="win32com"),
        ) as m:
            c.delete("eid")
        m.assert_called_once_with("eid", permanent=False)

    @patch("outlook_tool.HAS_WIN32", True)
    def test_delete_permanent_requires_double_flag(self):
        c = OutlookClient()
        with pytest.raises(ValueError, match="yes_really_permanent"):
            c.delete("eid", permanent=True)  # missing yes_really_permanent

    @patch("outlook_tool.HAS_WIN32", True)
    def test_delete_permanent_with_double_flag_proceeds(self):
        c = OutlookClient()
        with patch.object(
            c, "_delete_win32",
            return_value=OperationResult(status="success", operation="delete",
                                         message_id="eid", backend="win32com"),
        ) as m:
            c.delete("eid", permanent=True, yes_really_permanent=True)
        m.assert_called_once_with("eid", permanent=True)

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool._AppleScriptBackend")
    def test_mark_read_routes_to_applescript_with_parsed_index(self, mock_as_cls):
        mock_backend = MagicMock()
        mock_backend.mark_read.return_value = OperationResult(
            status="success", operation="mark_read",
            message_id="as_msg_7", backend="applescript",
        )
        mock_as_cls.return_value = mock_backend
        c = OutlookClient()
        c.mark_read("as_msg_7", read=True, folder_name="Inbox")
        mock_backend.mark_read.assert_called_once_with(
            msg_index=7, folder_name="Inbox", read=True,
        )

    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool._AppleScriptBackend")
    def test_move_routes_to_applescript_with_source_folder(self, mock_as_cls):
        mock_backend = MagicMock()
        mock_backend.move.return_value = OperationResult(
            status="success", operation="move",
            message_id="as_msg_3", backend="applescript",
        )
        mock_as_cls.return_value = mock_backend
        c = OutlookClient()
        c.move("as_msg_3", "Archive", source_folder="Inbox")
        mock_backend.move.assert_called_once_with(
            msg_index=3, target_folder="Archive", source_folder="Inbox",
        )


class TestParseAsMsgIndex:
    def test_strips_prefix(self):
        from outlook_tool import _parse_as_msg_index
        assert _parse_as_msg_index("as_msg_42") == 42

    def test_bare_numeric_ok(self):
        from outlook_tool import _parse_as_msg_index
        assert _parse_as_msg_index("7") == 7

    def test_rejects_junk(self):
        from outlook_tool import _parse_as_msg_index
        with pytest.raises(ValueError, match="numeric string"):
            _parse_as_msg_index("win32_abcd")


# =============================================================================
# Shared integration marker — gated by presence of the fixture
# =============================================================================

integration = pytest.mark.integration
