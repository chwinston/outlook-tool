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


def _make_applescript_backend(
    verify_delay_ms: int = 0,
    verify_use_count_check: bool = False,
):
    """Build an ``_AppleScriptBackend`` without invoking Outlook.

    Default ``verify_delay_ms=0`` and ``verify_use_count_check=False`` keep
    unit tests fast and independent of Bolt 5 count-check code paths unless
    the test explicitly opts in.
    """
    from outlook_tool import _AppleScriptBackend
    backend = _AppleScriptBackend.__new__(_AppleScriptBackend)
    backend._version = "16.0"
    backend.verify_delay_ms = verify_delay_ms
    backend.verify_use_count_check = verify_use_count_check
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
# Bolt 5 (B5.3) — verify-step hardening
# =============================================================================


class TestVerifyHardening:
    """B5.3 — settle delay + count-delta verify defeat Exchange sync-lag false-failures."""

    def test_settle_delay_is_called_before_verify_on_move(self):
        """_sleep_verify_delay must fire between AppleScript write and verify."""
        backend = _make_applescript_backend(verify_delay_ms=25)
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_verify_present_in_folder", return_value=False), \
             patch.object(backend, "_sleep_verify_delay") as sleep_mock, \
             patch("outlook_tool._run_applescript", return_value=""):
            backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        sleep_mock.assert_called_once()

    def test_settle_delay_is_called_on_hard_delete(self):
        backend = _make_applescript_backend(verify_delay_ms=25)
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "_verify_present_in_folder", return_value=False), \
             patch.object(backend, "_sleep_verify_delay") as sleep_mock, \
             patch("outlook_tool._run_applescript", return_value=""):
            backend.delete_message(msg_index=4, folder_name="Inbox", permanent=True)
        sleep_mock.assert_called_once()

    def test_settle_delay_is_called_on_mark_read(self):
        backend = _make_applescript_backend(verify_delay_ms=25)
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}
        after = {"folder": "Inbox", "is_read": True, "subject": "S", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", side_effect=[before, after]), \
             patch.object(backend, "_sleep_verify_delay") as sleep_mock, \
             patch("outlook_tool._run_applescript", return_value=""):
            backend.mark_read(msg_index=1, folder_name="Inbox", read=True)
        sleep_mock.assert_called_once()

    def test_count_drop_overrides_fingerprint_false_failure(self):
        """Hypothesis 2 fix: if source count dropped, treat as success even
        when the fingerprint scan still sees the message (cache lag)."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}
        # Count went 100 → 99 (the move actually dropped the source), but
        # the fingerprint scan still sees the message (sync lag).
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 99]), \
             patch.object(backend, "_verify_present_in_folder", return_value=True), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "success", r.error_detail
        # No present_in_target confirmation + count-drop only → unverified-success.
        # (we assert move_succeeded, the verified bit depends on target presence)
        assert "unconfirmed" in (r.error_detail or "").lower() or r.verified

    def test_count_stable_with_fingerprint_present_still_fails(self):
        """If the count didn't drop AND the fingerprint still sees the
        message, the move really didn't take effect."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 100]), \
             patch.object(backend, "_verify_present_in_folder", return_value=True), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "failed"
        assert r.verified is False
        assert "silent-success" in (r.error_detail or "")

    def test_count_check_disabled_falls_back_to_fingerprint(self):
        """With count-check off, verify behaves like Bolt 4 — fingerprint only."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=False,
        )
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return False  # not in source → move succeeded
            return True

        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count") as count_mock, \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        # Never call count when check is disabled.
        count_mock.assert_not_called()
        assert r.status == "success"
        assert r.verified is True

    def test_count_check_unavailable_falls_through_to_fingerprint(self):
        """If _folder_message_count returns -1 (resolve failed), fall back."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        before = {"folder": "Inbox", "is_read": False, "subject": "S", "sender_email": "a@b.com"}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return False  # fingerprint says: not in source → move succeeded
            return True  # confirmed in target

        with patch.object(backend, "_fetch_message_state", return_value=before), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", return_value=-1), \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            r = backend.move(msg_index=1, target_folder="Archive", source_folder="Inbox")
        assert r.status == "success"
        assert r.verified is True


# =============================================================================
# Bolt 5 (B5.2) — AppleScript bulk-move primitive (Shape A)
# =============================================================================


class TestAppleScriptMoveBulk:
    """move_bulk — Shape A: single tell block, refs-before-moves, count-delta verify."""

    def test_empty_moves_returns_empty(self):
        backend = _make_applescript_backend()
        assert backend.move_bulk([], source_folder="Inbox") == []

    def test_happy_path_count_drop_equals_batch_size(self):
        """Count dropped by N for N moves → every row success (unverified-success tier)."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre = {"folder": "Inbox", "is_read": False, "subject": "s", "sender_email": "a@b.com"}
        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 97]), \
             patch("outlook_tool._run_applescript", return_value="") as run_mock:
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Archive"), (9, "Deleted Items")],
                source_folder="Inbox",
            )
        assert len(results) == 3
        for r in results:
            assert r.status == "success", r.error_detail
            assert r.operation == "move"
            assert r.backend == "applescript"
        # One AppleScript call for the whole batch — Shape A promise.
        assert run_mock.call_count == 1

    def test_tell_block_captures_refs_before_moving(self):
        """The generated script must bind every m1..mN BEFORE issuing any move."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre = {"folder": "Inbox", "is_read": False, "subject": "s", "sender_email": "a@b.com"}
        captured = {}

        def capture(script, timeout=60):
            captured["script"] = script
            return ""

        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 98]), \
             patch("outlook_tool._run_applescript", side_effect=capture):
            backend.move_bulk(
                [(5, "Archive"), (7, "Archive")],
                source_folder="Inbox",
            )

        script = captured["script"]
        # The three `set m<i>` lines must come before any `move m<i>` line.
        set_positions = [script.index(f"set m{i}") for i in (1, 2)]
        move_positions = [script.index(f"move m{i}") for i in (1, 2)]
        assert max(set_positions) < min(move_positions), \
            "Shape A violated — refs must bind BEFORE any move fires"

    def test_partial_drop_falls_back_to_per_message_verify(self):
        """If count dropped by k < N, resolve via fingerprint verify per row."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        states = [
            {"folder": "Inbox", "subject": "A", "sender_email": "a@x.com", "is_read": False},
            {"folder": "Inbox", "subject": "B", "sender_email": "b@x.com", "is_read": False},
            {"folder": "Inbox", "subject": "C", "sender_email": "c@x.com", "is_read": False},
        ]

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            # A and B moved; C still in Inbox.
            if folder_name == "Inbox":
                return subject == "C"  # only C still there
            return subject in {"A", "B"}  # A, B landed in target

        with patch.object(backend, "_fetch_message_state", side_effect=states), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 98]), \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Archive"), (9, "Archive")],
                source_folder="Inbox",
            )
        # Expected drop was 3, got 2 → per-message fallback.
        assert results[0].status == "success"  # A
        assert results[0].verified is True
        assert results[1].status == "success"  # B
        assert results[1].verified is True
        assert results[2].status == "failed"    # C still in Inbox
        assert "bulk partial" in (results[2].error_detail or "")

    def test_script_error_marks_whole_batch_failed(self):
        """A tell-block exception is the one case we don't trust partial results."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre = {"folder": "Inbox", "subject": "s", "sender_email": "a@b.com", "is_read": False}
        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 100]), \
             patch("outlook_tool._run_applescript", side_effect=RuntimeError("boom")):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Archive")],
                source_folder="Inbox",
            )
        assert all(r.status == "failed" for r in results)
        assert all("applescript bulk move error" in (r.error_detail or "")
                   for r in results)

    def test_ensure_folder_failure_on_one_target_isolates_bad_rows(self):
        """Rows whose target folder can't be ensured are failed; others proceed."""
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre = {"folder": "Inbox", "subject": "s", "sender_email": "a@b.com", "is_read": False}

        def ensure_side_effect(name):
            if name == "Bad Folder":
                raise RuntimeError("ACL denied")

        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder", side_effect=ensure_side_effect), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 99]), \
             patch("outlook_tool._run_applescript", return_value=""):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Bad Folder"), (9, "Archive")],
                source_folder="Inbox",
            )
        assert results[0].status == "success"
        assert results[1].status == "failed"
        assert "ACL denied" in (results[1].error_detail or "")
        assert results[2].status == "success"

    def test_pre_fetch_failure_on_one_row_isolates_it(self):
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre_ok = {"folder": "Inbox", "subject": "s", "sender_email": "a@b.com", "is_read": False}

        def fetch_side(idx, folder):
            if idx == 7:
                raise RuntimeError("message_not_found")
            return pre_ok

        with patch.object(backend, "_fetch_message_state", side_effect=fetch_side), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 98]), \
             patch("outlook_tool._run_applescript", return_value=""):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Archive"), (9, "Archive")],
                source_folder="Inbox",
            )
        assert results[0].status == "success"
        assert results[1].status == "failed"
        assert "pre-fetch failed" in (results[1].error_detail or "")
        assert results[2].status == "success"

    def test_count_check_disabled_uses_per_message_verify(self):
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=False,
        )
        pre = {"folder": "Inbox", "subject": "s", "sender_email": "a@b.com", "is_read": False}

        def verify_side_effect(folder_name, subject, sender_email, scan_limit):
            if folder_name == "Inbox":
                return False  # gone from source
            return True  # landed in target

        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count") as count_mock, \
             patch.object(backend, "_verify_present_in_folder", side_effect=verify_side_effect), \
             patch("outlook_tool._run_applescript", return_value=""):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "Archive")],
                source_folder="Inbox",
            )
        # Count never queried when disabled.
        count_mock.assert_not_called()
        for r in results:
            assert r.status == "success"
            assert r.verified is True

    def test_empty_target_folder_rejected_per_row(self):
        backend = _make_applescript_backend(
            verify_delay_ms=0, verify_use_count_check=True,
        )
        pre = {"folder": "Inbox", "subject": "s", "sender_email": "a@b.com", "is_read": False}
        with patch.object(backend, "_fetch_message_state", return_value=pre), \
             patch.object(backend, "ensure_folder"), \
             patch.object(backend, "_folder_message_count", side_effect=[100, 99]), \
             patch("outlook_tool._run_applescript", return_value=""):
            results = backend.move_bulk(
                [(5, "Archive"), (7, "")],
                source_folder="Inbox",
            )
        assert results[0].status == "success"
        assert results[1].status == "failed"
        assert "non-empty" in (results[1].error_detail or "")


class TestOutlookClientBulkDispatch:
    """OutlookClient.move_bulk + delete_bulk — dispatch and double-flag."""

    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool._AppleScriptBackend.__init__", return_value=None)
    def test_move_bulk_routes_to_applescript_backend(self, _init):
        client = OutlookClient()
        client._applescript = MagicMock()
        client._applescript.move_bulk.return_value = [
            OperationResult(status="success", operation="move", message_id="as_msg_5",
                            backend="applescript"),
        ]
        results = client.move_bulk(
            [("as_msg_5", "Archive")],
            source_folder="Inbox",
        )
        client._applescript.move_bulk.assert_called_once()
        args, kwargs = client._applescript.move_bulk.call_args
        # First positional arg is the parsed-index list of tuples.
        assert args[0] == [(5, "Archive")]
        assert kwargs["source_folder"] == "Inbox"
        assert len(results) == 1

    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool._AppleScriptBackend.__init__", return_value=None)
    def test_delete_bulk_default_is_soft_delete(self, _init):
        client = OutlookClient()
        client._applescript = MagicMock()
        client._applescript.move_bulk.return_value = [
            OperationResult(status="success", operation="move", message_id="as_msg_5",
                            verified=True, backend="applescript",
                            after_state={"folder": "Deleted Items"}),
        ]
        results = client.delete_bulk(["as_msg_5"], folder_name="Inbox")
        # Routed through move_bulk with target "Deleted Items".
        args, _kwargs = client._applescript.move_bulk.call_args
        assert args[0] == [(5, "Deleted Items")]
        # Operation label is "delete" on the returned result.
        assert results[0].operation == "delete"
        assert results[0].after_state.get("permanent") is False

    @patch("outlook_tool.HAS_APPLESCRIPT", True)
    @patch("outlook_tool.HAS_WIN32", False)
    @patch("outlook_tool._AppleScriptBackend.__init__", return_value=None)
    def test_delete_bulk_permanent_requires_double_flag(self, _init):
        client = OutlookClient()
        client._applescript = MagicMock()
        with pytest.raises(ValueError, match="yes_really_permanent"):
            client.delete_bulk(["as_msg_5"], permanent=True)


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
