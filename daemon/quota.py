from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional

from db import get_connection, init_db

# Conservative daily caps strictly inside free tiers (0 billing risk)
DAILY_FREE_LIMITS = {
    "groq": {
        "max_requests": 1000,
        "max_tokens": 100000,
    },
    "gemini": {
        "max_requests": 250,
        "max_tokens": 250000,
    },
}

WARN_THRESHOLD_PCT = 80.0
BLOCK_THRESHOLD_PCT = 95.0


def _today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def check_quota(provider: str, db_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Evaluates today's usage for a provider against conservative free-tier caps.
    Returns allowed=False if usage >= 95%, warning=True if usage >= 80%.
    """
    provider = provider.lower().strip()
    limits = DAILY_FREE_LIMITS.get(provider, {"max_requests": 250, "max_tokens": 100000})
    init_db(db_path)

    today = _today_str()
    with get_connection(db_path) as conn:
        row = conn.execute(
            """
            SELECT
                COUNT(*) AS requests_used,
                COALESCE(SUM(total_tokens), 0) AS tokens_used
            FROM quota_ledger
            WHERE date = ? AND provider = ?
            """,
            (today, provider),
        ).fetchone()

    req_used = int(row["requests_used"])
    tok_used = int(row["tokens_used"])

    req_pct = (req_used / limits["max_requests"]) * 100.0 if limits["max_requests"] else 0.0
    tok_pct = (tok_used / limits["max_tokens"]) * 100.0 if limits["max_tokens"] else 0.0
    usage_pct = round(max(req_pct, tok_pct), 2)

    allowed = usage_pct < BLOCK_THRESHOLD_PCT
    warning = WARN_THRESHOLD_PCT <= usage_pct < BLOCK_THRESHOLD_PCT

    return {
        "provider": provider,
        "date": today,
        "allowed": allowed,
        "warning": warning,
        "usage_pct": usage_pct,
        "requests_used": req_used,
        "requests_limit": limits["max_requests"],
        "tokens_used": tok_used,
        "tokens_limit": limits["max_tokens"],
    }


def record_usage(
    provider: str,
    model: str,
    prompt_tokens: int = 0,
    completion_tokens: int = 0,
    total_tokens: Optional[int] = None,
    db_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """Logs a completed LLM request and its token consumption into quota_ledger."""
    provider = provider.lower().strip()
    init_db(db_path)

    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    ts_str = now.isoformat(timespec="seconds")
    total = total_tokens if total_tokens is not None else (prompt_tokens + completion_tokens)

    with get_connection(db_path) as conn:
        conn.execute(
            """
            INSERT INTO quota_ledger (date, timestamp, provider, model, prompt_tokens, completion_tokens, total_tokens)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (date_str, ts_str, provider, model, int(prompt_tokens), int(completion_tokens), int(total)),
        )

    return check_quota(provider, db_path=db_path)


def get_daily_quota_summary(db_path: Optional[Path] = None) -> Dict[str, Any]:
    """Returns daily free-tier status across all configured providers."""
    return {
        provider: check_quota(provider, db_path=db_path)
        for provider in DAILY_FREE_LIMITS
    }

