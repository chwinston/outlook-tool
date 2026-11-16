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
# Shared integration marker — gated by presence of the fixture
# =============================================================================

integration = pytest.mark.integration
