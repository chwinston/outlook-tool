"""Tests for the LLM-backed triage classifier — Bolt 6 Phase A2.

All API calls are mocked — no network, no cost, no content leak. Tests cover:

* JSON-verdict parsing (happy path, malformed, missing fields)
* ``limit`` cap (N > limit → only ``limit`` are classified)
* confidence thresholds (auto / downgrade / unclear-force)
* eligibility filters (missing body_preview, missing subject)
* CLI entry point round-trip (input JSON → output JSON)
* ``--limit`` CLI flag honored
* prompt caching flag present on API call
* API error handled gracefully (returns messages unchanged, logs warning)
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

_REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_REPO_ROOT))
sys.path.insert(0, str(_REPO_ROOT / "scripts"))

import triage_llm  # noqa: E402


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _msg(
    *,
    message_id: str = "as_msg_1",
    sender_email: str = "someone@example.com",
    subject: str = "hello",
    body_preview: str = "please take a look",
    bucket: str = "keep",
    matched_rule: str = "",
) -> dict:
    """A triage-report-shaped message (what ``scripts/triage.py`` emits)."""
    return {
        "message_id": message_id,
        "sender_email": sender_email,
        "sender_name": "Someone",
        "subject": subject,
        "received_at": "2026-11-16T10:00:00",
        "folder": "Inbox",
        "bucket": bucket,
        "matched_rule": matched_rule,
        "body_preview": body_preview,
    }


def _mock_response(text: str) -> MagicMock:
    """Shape the mock the way the Anthropic SDK returns."""
    block = MagicMock()
    block.text = text
    response = MagicMock()
    response.content = [block]
    return response


def _mock_client(*verdicts_json: str) -> MagicMock:
    """Build a mock Anthropic client that returns the given verdicts in order."""
    client = MagicMock()
    client.messages.create.side_effect = [
        _mock_response(v) for v in verdicts_json
    ]
    return client


# ---------------------------------------------------------------------------
# Unit tests — parsing + gating
# ---------------------------------------------------------------------------

class TestParseVerdict:
    def test_parses_valid_json(self):
        v = triage_llm._parse_verdict(
            '{"bucket": "junk", "confidence": 0.9, "reason": "cold outreach"}'
        )
        assert v == {"bucket": "junk", "confidence": 0.9, "reason": "cold outreach"}

    def test_tolerates_surrounding_prose(self):
        v = triage_llm._parse_verdict(
            'Here is the verdict:\n'
            '{"bucket": "action", "confidence": 0.75, "reason": "direct ask"}\n'
            'Thanks!'
        )
        assert v["bucket"] == "action"
        assert v["confidence"] == 0.75

    def test_handles_malformed_json(self):
        v = triage_llm._parse_verdict("not json at all")
        assert v == {
            "bucket": "unclear",
            "confidence": 0.0,
            "reason": "parse-failed",
        }

    def test_handles_truncated_json(self):
        v = triage_llm._parse_verdict('{"bucket": "junk", "confidence": 0.9')
        assert v["bucket"] == "unclear"
        assert v["confidence"] == 0.0

    def test_unknown_bucket_becomes_unclear(self):
        v = triage_llm._parse_verdict(
            '{"bucket": "spam", "confidence": 0.95, "reason": "noise"}'
        )
        # "spam" isn't a valid bucket → fall back to unclear
        assert v["bucket"] == "unclear"

    def test_confidence_clamped_to_unit_interval(self):
        v = triage_llm._parse_verdict(
            '{"bucket": "junk", "confidence": 1.5, "reason": "very sure"}'
        )
        assert v["confidence"] == 1.0

        v2 = triage_llm._parse_verdict(
            '{"bucket": "junk", "confidence": -0.5, "reason": "very unsure"}'
        )
        assert v2["confidence"] == 0.0


class TestConfidenceGate:
    def test_high_confidence_routes_to_auto(self):
        gated = triage_llm._apply_confidence_gate(
            {"bucket": "junk", "confidence": 0.95, "reason": "clear"},
            threshold=0.85,
        )
        assert gated["route"] == "auto"
        assert gated["bucket"] == "junk"

    def test_mid_confidence_downgrades_to_needs_action(self):
        gated = triage_llm._apply_confidence_gate(
            {"bucket": "junk", "confidence": 0.75, "reason": "maybe"},
            threshold=0.85,
        )
        # bucket preserved, but routing says operator should review
        assert gated["bucket"] == "junk"
        assert gated["route"] == "needs-action"

    def test_low_confidence_forces_unclear(self):
        gated = triage_llm._apply_confidence_gate(
            {"bucket": "junk", "confidence": 0.5, "reason": "unsure"},
            threshold=0.85,
        )
        assert gated["bucket"] == "unclear"
        assert gated["route"] == "needs-action"


# ---------------------------------------------------------------------------
# Eligibility tests — skip messages missing required fields
# ---------------------------------------------------------------------------

class TestEligibility:
    def test_skips_message_missing_body_preview(self):
        client = MagicMock()
        msgs = [_msg(body_preview="")]
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        assert "llm_verdict" not in out[0]
        # No API call made.
        client.messages.create.assert_not_called()

    def test_skips_message_missing_subject(self):
        client = MagicMock()
        msgs = [_msg(subject="")]
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        assert "llm_verdict" not in out[0]
        client.messages.create.assert_not_called()

    def test_skips_message_missing_sender(self):
        client = MagicMock()
        msgs = [_msg(sender_email="")]
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        assert "llm_verdict" not in out[0]
        client.messages.create.assert_not_called()

    def test_does_not_classify_non_keep_bucket(self):
        client = MagicMock()
        msgs = [_msg(bucket="safe-delete", matched_rule="some-rule")]
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        assert "llm_verdict" not in out[0]
        client.messages.create.assert_not_called()
        # Original bucket + matched_rule preserved.
        assert out[0]["bucket"] == "safe-delete"
        assert out[0]["matched_rule"] == "some-rule"


# ---------------------------------------------------------------------------
# End-to-end classifier behavior (mocked API)
# ---------------------------------------------------------------------------

class TestClassifyMessages:
    def test_respects_limit(self):
        # 300 messages but limit=100 → only 100 classified.
        msgs = [_msg(message_id=f"as_msg_{i}") for i in range(300)]
        verdicts = [
            '{"bucket": "junk", "confidence": 0.9, "reason": "r"}'
        ] * 100
        client = _mock_client(*verdicts)

        out = triage_llm.classify_messages_with_llm(msgs, limit=100, client=client)

        verdict_count = sum(1 for m in out if "llm_verdict" in m)
        assert verdict_count == 100
        assert client.messages.create.call_count == 100
        # Messages after the limit are present but un-augmented.
        assert len(out) == 300
        assert "llm_verdict" not in out[200]

    def test_respects_confidence_threshold(self):
        # Model says 0.6 confidence → gate forces unclear + needs-action.
        msgs = [_msg()]
        client = _mock_client(
            '{"bucket": "junk", "confidence": 0.6, "reason": "unsure"}'
        )
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        v = out[0]["llm_verdict"]
        assert v["bucket"] == "unclear"
        assert v["route"] == "needs-action"

    def test_handles_malformed_response(self):
        msgs = [_msg()]
        client = _mock_client("this is not JSON")
        out = triage_llm.classify_messages_with_llm(msgs, client=client)
        v = out[0]["llm_verdict"]
        assert v["bucket"] == "unclear"
        assert v["confidence"] == 0.0

    def test_api_error_handled_gracefully(self, caplog):
        # If the API raises, the message passes through un-augmented, and
        # subsequent messages still get classified. No exception propagates.
        msgs = [_msg(message_id="m1"), _msg(message_id="m2")]
        client = MagicMock()
        client.messages.create.side_effect = [
            RuntimeError("boom"),
            _mock_response(
                '{"bucket": "junk", "confidence": 0.9, "reason": "ok"}'
            ),
        ]
        with caplog.at_level("WARNING"):
            out = triage_llm.classify_messages_with_llm(msgs, client=client)

        assert "llm_verdict" not in out[0]
        assert out[1]["llm_verdict"]["bucket"] == "junk"
        assert any("llm classify failed" in r.message for r in caplog.records)

    def test_input_messages_not_mutated(self):
        msgs = [_msg()]
        snapshot = json.dumps(msgs, sort_keys=True)
        client = _mock_client(
            '{"bucket": "junk", "confidence": 0.9, "reason": "clear"}'
        )
        triage_llm.classify_messages_with_llm(msgs, client=client)
        assert json.dumps(msgs, sort_keys=True) == snapshot

    def test_prompt_caching_headers_present(self):
        # The system prompt must be sent as a content block with
        # cache_control=ephemeral so the API caches the shared prefix.
        msgs = [_msg()]
        client = _mock_client(
            '{"bucket": "junk", "confidence": 0.9, "reason": "r"}'
        )
        triage_llm.classify_messages_with_llm(msgs, client=client)

        kwargs = client.messages.create.call_args.kwargs
        assert "system" in kwargs
        system = kwargs["system"]
        assert isinstance(system, list)
        assert system[0]["type"] == "text"
        assert system[0]["cache_control"] == {"type": "ephemeral"}
        assert "Classify messages into exactly one bucket" in system[0]["text"]

    def test_user_prompt_shape(self):
        msgs = [_msg(
            sender_email="rep@gartner.com",
            subject="Quick sync?",
            body_preview="Hey Chris, do you have 15 min next week to chat?",
        )]
        client = _mock_client(
            '{"bucket": "junk", "confidence": 0.9, "reason": "cold"}'
        )
        triage_llm.classify_messages_with_llm(msgs, client=client)

        user_content = client.messages.create.call_args.kwargs["messages"][0]
        assert user_content["role"] == "user"
        text = user_content["content"]
        assert "From: rep@gartner.com" in text
        assert "Subject: Quick sync?" in text
        assert "Body preview: Hey Chris" in text

    def test_model_passthrough(self):
        msgs = [_msg()]
        client = _mock_client(
            '{"bucket": "junk", "confidence": 0.9, "reason": "r"}'
        )
        triage_llm.classify_messages_with_llm(
            msgs, client=client, model="claude-haiku-4-5-20251001"
        )
        kwargs = client.messages.create.call_args.kwargs
        assert kwargs["model"] == "claude-haiku-4-5-20251001"


# ---------------------------------------------------------------------------
# Integration — Anthropic client is lazily built from env var
# ---------------------------------------------------------------------------

class TestClientConstruction:
    def test_build_client_requires_api_key(self, monkeypatch):
        monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
        with pytest.raises(RuntimeError, match="ANTHROPIC_API_KEY not set"):
            triage_llm._build_client(api_key=None)

    def test_build_client_accepts_explicit_key(self):
        with patch("anthropic.Anthropic") as mock_cls:
            triage_llm._build_client(api_key="sk-test-123")
            mock_cls.assert_called_once_with(api_key="sk-test-123")

    def test_build_client_falls_back_to_env(self, monkeypatch):
        monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-env-456")
        with patch("anthropic.Anthropic") as mock_cls:
            triage_llm._build_client(api_key=None)
            mock_cls.assert_called_once_with(api_key="sk-env-456")

    def test_classify_fails_open_when_client_build_fails(self, monkeypatch, caplog):
        # No API key → _build_client raises → remaining messages flow through.
        monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
        msgs = [_msg(message_id="a"), _msg(message_id="b")]
        with caplog.at_level("WARNING"):
            out = triage_llm.classify_messages_with_llm(msgs)
        assert all("llm_verdict" not in m for m in out)
        assert len(out) == 2


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

class TestCLI:
    def test_cli_reads_input_writes_output(self, tmp_path):
        in_path = tmp_path / "in.json"
        out_path = tmp_path / "out.json"
        in_report = {
            "schema_version": 1,
            "generated_at": "2026-11-17T10:00:00",
            "messages": [_msg(message_id="x", bucket="keep")],
            "summary": {"keep": 1},
        }
        in_path.write_text(json.dumps(in_report))

        with patch.object(triage_llm, "classify_messages_with_llm") as mock_fn:
            mock_fn.return_value = [
                {
                    **_msg(message_id="x", bucket="keep"),
                    "llm_verdict": {
                        "bucket": "junk",
                        "confidence": 0.9,
                        "reason": "test",
                        "route": "auto",
                    },
                }
            ]
            rc = triage_llm.main([
                "--input", str(in_path),
                "--output", str(out_path),
                "--quiet",
            ])

        assert rc == 0
        assert out_path.exists()
        written = json.loads(out_path.read_text())
        assert written["messages"][0]["llm_verdict"]["bucket"] == "junk"
        assert written["llm_summary"]["verdicts_added"] == 1

    def test_cli_limit_flag_passed_through(self, tmp_path):
        in_path = tmp_path / "in.json"
        out_path = tmp_path / "out.json"
        in_report = {"messages": [], "summary": {}}
        in_path.write_text(json.dumps(in_report))

        with patch.object(triage_llm, "classify_messages_with_llm") as mock_fn:
            mock_fn.return_value = []
            triage_llm.main([
                "--input", str(in_path),
                "--output", str(out_path),
                "--limit", "42",
                "--quiet",
            ])

        kwargs = mock_fn.call_args.kwargs
        assert kwargs["limit"] == 42

    def test_cli_missing_input_returns_error(self, tmp_path):
        rc = triage_llm.main([
            "--input", str(tmp_path / "nope.json"),
            "--output", str(tmp_path / "out.json"),
            "--quiet",
        ])
        assert rc == 2

    def test_cli_confidence_threshold_flag(self, tmp_path):
        in_path = tmp_path / "in.json"
        out_path = tmp_path / "out.json"
        in_path.write_text(json.dumps({"messages": [], "summary": {}}))

        with patch.object(triage_llm, "classify_messages_with_llm") as mock_fn:
            mock_fn.return_value = []
            triage_llm.main([
                "--input", str(in_path),
                "--output", str(out_path),
                "--confidence-threshold", "0.7",
                "--quiet",
            ])
        kwargs = mock_fn.call_args.kwargs
        assert kwargs["confidence_threshold"] == 0.7
