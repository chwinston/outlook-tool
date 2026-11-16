"""
CLI tests for ``outlook-tool move / delete / mark-read``.

Covers the shape, not the real backend. Backend behaviour is tested
exhaustively in ``tests/test_write_ops.py``.

Dry-run is the default — every test that does NOT pass ``--apply`` must
exit 0 WITHOUT instantiating an ``OutlookClient``. We prove that by
patching ``cli.OutlookClient`` to raise if called.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import cli  # noqa: E402
from outlook_tool import OperationResult  # noqa: E402


def _run_cli(argv, monkeypatch, outlook_client_factory=None):
    """Invoke cli.main() with a fake argv. Return (stdout, stderr, exit_code).

    If outlook_client_factory is given, it is called to produce the
    OutlookClient used by cmd_* handlers. If omitted, OutlookClient() will
    RAISE — any test that doesn't provide one is asserting the dry-run
    path never touches the client.
    """
    monkeypatch.setattr(sys, "argv", ["outlook-tool"] + argv)
    captured_out = io.StringIO()
    captured_err = io.StringIO()
    monkeypatch.setattr(sys, "stdout", captured_out)
    monkeypatch.setattr(sys, "stderr", captured_err)

    if outlook_client_factory is None:
        def _raise(*a, **k):
            raise RuntimeError("OutlookClient was constructed during a dry-run")
        monkeypatch.setattr(cli, "OutlookClient", _raise)
    else:
        monkeypatch.setattr(cli, "OutlookClient", outlook_client_factory)

    exit_code = 0
    try:
        cli.main()
    except SystemExit as e:
        exit_code = e.code if isinstance(e.code, int) else 1
    return captured_out.getvalue(), captured_err.getvalue(), exit_code


# =============================================================================
# move
# =============================================================================

class TestMoveCLI:
    def test_dry_run_default_does_not_call_backend(self, monkeypatch):
        out, err, code = _run_cli(["move", "as_msg_7", "Archive"], monkeypatch)
        assert code == 0
        assert "DRY RUN" in err
        assert "Archive" in err
        assert "[DRY-RUN]" in out

    def test_dry_run_json_emits_dry_run_status(self, monkeypatch):
        out, err, code = _run_cli(["move", "as_msg_7", "Archive", "--json"], monkeypatch)
        assert code == 0
        assert '"status": "dry_run"' in out

    def test_apply_invokes_client_move(self, monkeypatch):
        mock_client = MagicMock()
        mock_client.move.return_value = OperationResult(
            status="success", operation="move",
            message_id="as_msg_7", verified=True, backend="applescript",
        )
        factory = MagicMock(return_value=mock_client)
        out, err, code = _run_cli(
            ["move", "as_msg_7", "Archive", "--apply"],
            monkeypatch, outlook_client_factory=factory,
        )
        mock_client.move.assert_called_once_with(
            "as_msg_7", "Archive", source_folder=None,
        )
        assert code == 0

    def test_apply_failure_exits_1(self, monkeypatch):
        mock_client = MagicMock()
        mock_client.move.return_value = OperationResult(
            status="failed", operation="move",
            message_id="as_msg_7", backend="applescript",
            error_detail="silent-success",
        )
        out, err, code = _run_cli(
            ["move", "as_msg_7", "Archive", "--apply"],
            monkeypatch, outlook_client_factory=MagicMock(return_value=mock_client),
        )
        assert code == 1

    def test_apply_unsupported_exits_2(self, monkeypatch):
        mock_client = MagicMock()
        mock_client.move.return_value = OperationResult(
            status="unsupported", operation="move",
            message_id="mid", backend="graph",
            error_detail="scope missing",
        )
        out, err, code = _run_cli(
            ["move", "mid", "Archive", "--apply"],
            monkeypatch, outlook_client_factory=MagicMock(return_value=mock_client),
        )
        assert code == 2


# =============================================================================
# delete
# =============================================================================

class TestDeleteCLI:
    def test_dry_run_soft_delete_default(self, monkeypatch):
        out, err, code = _run_cli(["delete", "as_msg_7"], monkeypatch)
        assert code == 0
        assert "SOFT DELETE" in err
        assert "[DRY-RUN]" in out

    def test_permanent_without_yes_really_permanent_refused(self, monkeypatch):
        out, err, code = _run_cli(
            ["delete", "as_msg_7", "--permanent", "--apply"], monkeypatch,
        )
        assert code == 2
        assert "REFUSED" in err
        assert "yes-really-permanent" in err

    def test_permanent_with_double_flag_dry_runs_cleanly(self, monkeypatch):
        out, err, code = _run_cli(
            ["delete", "as_msg_7", "--permanent", "--yes-really-permanent"],
            monkeypatch,
        )
        assert code == 0
        assert "HARD DELETE" in err
        assert "[DRY-RUN]" in out

    def test_permanent_with_double_flag_and_apply_invokes_backend(self, monkeypatch):
        mock_client = MagicMock()
        mock_client.delete.return_value = OperationResult(
            status="success", operation="delete",
            message_id="as_msg_7", verified=True, backend="applescript",
            after_state={"permanent": True},
        )
        out, err, code = _run_cli(
            ["delete", "as_msg_7", "--permanent", "--yes-really-permanent", "--apply"],
            monkeypatch, outlook_client_factory=MagicMock(return_value=mock_client),
        )
        mock_client.delete.assert_called_once_with(
            "as_msg_7",
            permanent=True, yes_really_permanent=True, folder_name=None,
        )
        assert code == 0


# =============================================================================
# mark-read
# =============================================================================

class TestMarkReadCLI:
    def test_dry_run_default_read(self, monkeypatch):
        out, err, code = _run_cli(["mark-read", "as_msg_7"], monkeypatch)
        assert code == 0
        assert "mark READ" in err
        assert "[DRY-RUN]" in out

    def test_dry_run_unread_flip(self, monkeypatch):
        out, err, code = _run_cli(
            ["mark-read", "as_msg_7", "--unread"], monkeypatch,
        )
        assert code == 0
        assert "mark UNREAD" in err

    def test_apply_invokes_client(self, monkeypatch):
        mock_client = MagicMock()
        mock_client.mark_read.return_value = OperationResult(
            status="success", operation="mark_read",
            message_id="as_msg_7", verified=True, backend="applescript",
        )
        out, err, code = _run_cli(
            ["mark-read", "as_msg_7", "--apply", "--folder", "Inbox"],
            monkeypatch, outlook_client_factory=MagicMock(return_value=mock_client),
        )
        mock_client.mark_read.assert_called_once_with(
            "as_msg_7", read=True, folder_name="Inbox",
        )
        assert code == 0
