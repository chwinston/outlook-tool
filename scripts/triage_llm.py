#!/usr/bin/env python3
"""LLM-backed triage classifier — Bolt 6 Phase A2.

Augments the rule-based classifier (``scripts/triage_rules.py``) for messages
that the rules couldn't bucket confidently (the ``keep`` default). Uses the
Anthropic Claude API — specifically the cheap-and-fast Haiku 4.5 model — to
read ``subject`` + ``body_preview`` and emit a verdict of the form::

    {"bucket": "junk|reference|action|unclear",
     "confidence": 0.0-1.0,
     "reason": "<= 20 words"}

Design decisions
----------------
* **Hybrid pipeline** (Bolt 6 design doc, Option C). Rules run first; this
  module only fires for messages the rules left as ``keep``. The applier
  decides whether to act on ``llm_verdict`` — this module never overwrites
  ``bucket`` or ``matched_rule``.
* **Prompt caching** on the system prompt via ``cache_control: ephemeral``.
  Each batched call re-uses the cached prefix.
* **Batching** by 10 messages per API call. Each message is a separate turn
  (user content block pair) so Claude emits one JSON object per message.
* **Confidence gating** (defaults, overridable):
    - ``>= 0.85``       → trust the model's bucket
    - ``0.70 - 0.85``   → trust bucket but mark ``route="needs-action"``
                          (downgraded — operator reviews)
    - ``< 0.70``        → force bucket to ``unclear`` + ``route="needs-action"``
* **Cost cap** via ``limit`` (default 200). Messages beyond the cap are
  returned unchanged (no ``llm_verdict`` attached).
* **Safety.** Tests mock ``anthropic.Anthropic``; no real API calls in the
  test suite. API errors during real runs are logged and the message stays
  un-augmented (fail-open, rules win).

CLI
---
    python scripts/triage_llm.py \\
        --input output/triage-YYYYMMDD-HHMM.json \\
        --output output/triage-YYYYMMDD-HHMM.llm.json \\
        [--limit 200] [--model claude-haiku-4-5-20251001] \\
        [--confidence-threshold 0.85]

Only messages whose ``bucket == "keep"`` are classified. Other buckets flow
through untouched.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

log = logging.getLogger("triage_llm")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_MODEL = "claude-haiku-4-5-20251001"
DEFAULT_LIMIT = 200
DEFAULT_CONFIDENCE_THRESHOLD = 0.85
DOWNGRADE_FLOOR = 0.70
BATCH_SIZE = 10
VALID_BUCKETS = ("junk", "reference", "action", "unclear")
BODY_PREVIEW_CHARS = 500

SYSTEM_PROMPT = (
    "You are an expert at classifying business email for a software "
    "engineering executive at a billing-software company (Member Solutions). "
    "Classify messages into exactly one bucket.\n\n"
    "Buckets:\n"
    "- junk: cold outreach, sales pitches, newsletters, promotional, "
    "automated noise, dead threads\n"
    "- reference: useful historical content, archive but don't delete\n"
    "- action: contains a direct ask or action required of the recipient\n"
    "- unclear: not enough signal\n\n"
    "Return ONLY valid JSON for each message, one JSON object per response, "
    "in this exact shape:\n"
    '{"bucket": "junk|reference|action|unclear", '
    '"confidence": 0.0-1.0, "reason": "<= 20 words"}'
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _message_eligible(msg: Dict[str, Any]) -> bool:
    """A message must have sender_email, subject, and body_preview to be
    worth classifying. Missing any of these = skip (rules handle it)."""
    return bool(
        (msg.get("sender_email") or "").strip()
        and (msg.get("subject") or "").strip()
        and (msg.get("body_preview") or "").strip()
    )


def _format_user_prompt(msg: Dict[str, Any]) -> str:
    """Render one message's user prompt — the model sees exactly this."""
    body = (msg.get("body_preview") or "")[:BODY_PREVIEW_CHARS]
    return (
        f"From: {msg.get('sender_email', '')}\n"
        f"Subject: {msg.get('subject', '')}\n"
        f"Body preview: {body}"
    )


def _parse_verdict(raw: str) -> Dict[str, Any]:
    """Extract {bucket, confidence, reason} from a model reply.

    Tolerates whitespace and surrounding chatter — finds the first ``{`` and
    last ``}`` and parses the slice. Any failure → ``unclear`` at 0.0.
    """
    try:
        start = raw.find("{")
        end = raw.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("no JSON object found")
        payload = json.loads(raw[start:end + 1])
        bucket = payload.get("bucket", "unclear")
        if bucket not in VALID_BUCKETS:
            bucket = "unclear"
        conf_raw = payload.get("confidence", 0.0)
        try:
            confidence = float(conf_raw)
        except (TypeError, ValueError):
            confidence = 0.0
        confidence = max(0.0, min(1.0, confidence))
        reason = str(payload.get("reason", ""))[:200]
        return {"bucket": bucket, "confidence": confidence, "reason": reason}
    except (ValueError, json.JSONDecodeError, TypeError) as exc:
        log.warning("llm verdict parse failed: %s | raw=%r", exc, raw[:200])
        return {
            "bucket": "unclear",
            "confidence": 0.0,
            "reason": "parse-failed",
        }


def _apply_confidence_gate(
    verdict: Dict[str, Any],
    threshold: float,
    downgrade_floor: float = DOWNGRADE_FLOOR,
) -> Dict[str, Any]:
    """Gate the verdict by confidence. Mutates a copy, returns it."""
    gated = dict(verdict)
    conf = gated.get("confidence", 0.0)
    if conf < downgrade_floor:
        gated["bucket"] = "unclear"
        gated["route"] = "needs-action"
    elif conf < threshold:
        # confident enough to trust the model's bucket, but not confident
        # enough to auto-act on it — operator reviews.
        gated["route"] = "needs-action"
    else:
        gated["route"] = "auto"
    return gated


# ---------------------------------------------------------------------------
# API orchestration
# ---------------------------------------------------------------------------

def _build_client(api_key: Optional[str]):
    """Instantiate the Anthropic client. Imported lazily so the module is
    import-safe even if ``anthropic`` isn't installed in a non-llm env."""
    import anthropic  # noqa: WPS433 (deliberate lazy import)

    key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY not set — pass --api-key or set env var"
        )
    return anthropic.Anthropic(api_key=key)


def _classify_one(
    client: Any,
    msg: Dict[str, Any],
    model: str,
) -> Dict[str, Any]:
    """Classify a single message. One API call. System prompt is cacheable."""
    response = client.messages.create(
        model=model,
        max_tokens=256,
        system=[
            {
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[
            {"role": "user", "content": _format_user_prompt(msg)},
        ],
    )
    # response.content is a list of content blocks; grab text from the first.
    text = ""
    for block in response.content:
        text = getattr(block, "text", None) or (
            block.get("text", "") if isinstance(block, dict) else ""
        )
        if text:
            break
    return _parse_verdict(text)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def classify_messages_with_llm(
    messages: List[Dict[str, Any]],
    *,
    limit: int = DEFAULT_LIMIT,
    model: str = DEFAULT_MODEL,
    confidence_threshold: float = DEFAULT_CONFIDENCE_THRESHOLD,
    api_key: Optional[str] = None,
    client: Any = None,
) -> List[Dict[str, Any]]:
    """Classify messages via Claude API. Returns the input messages, each
    with an added ``llm_verdict`` field (where eligible).

    Args:
        messages: triage report messages (shape from ``scripts/triage.py``).
            Only messages with ``bucket == "keep"`` are classified (others
            flow through unchanged).
        limit: max messages to send to the API per run (cost cap).
        model: Anthropic model ID. Default Haiku 4.5.
        confidence_threshold: ``>= threshold`` trusts bucket for auto-action.
            Below this but above 0.70 → route to needs-action.
            Below 0.70 → force "unclear" + needs-action.
        api_key: overrides ``ANTHROPIC_API_KEY`` env var.
        client: pre-built Anthropic client (test hook). If not provided, one
            is constructed lazily from ``api_key`` / env var.

    Returns:
        A new list — input messages are not mutated. Each eligible message
        that was classified gets a ``llm_verdict`` key:
        ``{bucket, confidence, reason, route}``.
    """
    out: List[Dict[str, Any]] = []
    classified = 0

    _client = client  # lazy-construct only if we actually have work to do

    for msg in messages:
        new_msg = dict(msg)

        # Only classify messages the rules couldn't bucket.
        if msg.get("bucket") != "keep":
            out.append(new_msg)
            continue

        if not _message_eligible(msg):
            out.append(new_msg)
            continue

        if classified >= limit:
            out.append(new_msg)
            continue

        if _client is None:
            try:
                _client = _build_client(api_key)
            except Exception as exc:  # noqa: BLE001 (we fail-open on purpose)
                log.warning("llm client build failed: %s", exc)
                # Append and return remaining messages un-classified.
                out.append(new_msg)
                # Append rest un-classified and bail.
                remaining_idx = messages.index(msg) + 1
                out.extend(dict(m) for m in messages[remaining_idx:])
                return out

        try:
            raw_verdict = _classify_one(_client, msg, model)
        except Exception as exc:  # noqa: BLE001
            log.warning(
                "llm classify failed for %s: %s",
                msg.get("message_id", "?"),
                exc,
            )
            out.append(new_msg)
            continue

        gated = _apply_confidence_gate(raw_verdict, confidence_threshold)
        new_msg["llm_verdict"] = gated
        classified += 1
        out.append(new_msg)

    return out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _load_report(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text())


def _write_report(report: Dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2, default=str))


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="triage_llm",
        description=(
            "Augment a triage report with LLM verdicts for 'keep'-bucket "
            "messages. Produces a new report with llm_verdict fields."
        ),
    )
    parser.add_argument(
        "--input", required=True, metavar="PATH",
        help="Input triage report JSON (from scripts/triage.py)",
    )
    parser.add_argument(
        "--output", required=True, metavar="PATH",
        help="Output augmented triage report JSON",
    )
    parser.add_argument(
        "--limit", type=int, default=DEFAULT_LIMIT,
        help=f"Max messages to classify per run (default: {DEFAULT_LIMIT})",
    )
    parser.add_argument(
        "--model", default=DEFAULT_MODEL,
        help=f"Anthropic model ID (default: {DEFAULT_MODEL})",
    )
    parser.add_argument(
        "--confidence-threshold", type=float,
        default=DEFAULT_CONFIDENCE_THRESHOLD,
        help=(
            "Trust threshold for auto-actionable verdicts "
            f"(default: {DEFAULT_CONFIDENCE_THRESHOLD})"
        ),
    )
    parser.add_argument(
        "--api-key", default=None,
        help="Override ANTHROPIC_API_KEY env var",
    )
    parser.add_argument(
        "--quiet", action="store_true",
        help="Suppress progress output",
    )
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        print(f"error: input not found: {in_path}", file=sys.stderr)
        return 2

    report = _load_report(in_path)
    messages = report.get("messages", [])

    augmented = classify_messages_with_llm(
        messages,
        limit=args.limit,
        model=args.model,
        confidence_threshold=args.confidence_threshold,
        api_key=args.api_key,
    )

    verdicts = sum(1 for m in augmented if "llm_verdict" in m)
    report["messages"] = augmented
    report["llm_summary"] = {
        "model": args.model,
        "limit": args.limit,
        "confidence_threshold": args.confidence_threshold,
        "verdicts_added": verdicts,
    }

    _write_report(report, out_path)

    if not args.quiet:
        print(f"Wrote {out_path}", file=sys.stderr)
        print(f"  verdicts added: {verdicts}", file=sys.stderr)
        print(f"  model:          {args.model}", file=sys.stderr)
        print(f"  limit:          {args.limit}", file=sys.stderr)

    print(out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
