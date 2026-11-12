import os
from pathlib import Path
from typing import Tuple, Dict, Any
import httpx
from dotenv import load_dotenv

from quota import check_quota, record_usage, get_daily_quota_summary
from notify import notify_windows

# Load .env from project root
env_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path=env_path)

SYSTEM_PROMPT = (
    "You are Rohit Ganesh Patil's senior engineering mentor and candid technical sparring partner. "
    "You know his context: Final-year B.E. AIML student at SIES GST (8.06 CGPA, 10th 83.83%, 12th 58.33% "
    "bypassing 60%-filtered mass recruiters in favor of product startups/off-campus roles), AI-augmented builder "
    "(Hisab Hero, Sentinel Grid, Bio-Nexus, SomnoSense), and open-sea endurance swimmer. "
    "Never flatter, apologize, or use corporate filler. Point out architectural flaws directly, explain why they "
    "fail in production, and give the simplest working alternative."
)


GROQ_MODELS = ("openai/gpt-oss-120b", "qwen/qwen3.8-27b")
GEMINI_MODEL = "gemini-3.8-flash"


async def query_groq(
    prompt: str,
    api_key: str,
    system_prompt: str = SYSTEM_PROMPT,
    max_tokens: int = 1500,
    model_name: str = GROQ_MODELS[0],
) -> Tuple[str, Dict[str, int]]:
    """Calls Groq API with the specified free model and returns (answer, token_usage)."""
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.4,
        "max_tokens": max_tokens
    }
    async with httpx.AsyncClient(timeout=20.0) as client:
        resp = await client.post(url, headers=headers, json=payload)
        resp.raise_for_status()
        data = resp.json()
        usage = data.get("usage", {})
        tokens = {
            "prompt": int(usage.get("prompt_tokens", 0)),
            "completion": int(usage.get("completion_tokens", 0)),
            "total": int(usage.get("total_tokens", 0)),
        }
        return data["choices"][0]["message"]["content"], tokens


async def query_gemini(
    prompt: str,
    api_key: str,
    system_prompt: str = SYSTEM_PROMPT,
    max_tokens: int = 1500,
    model_name: str = GEMINI_MODEL,
) -> Tuple[str, Dict[str, int]]:
    """Calls Google Gemini API via x-goog-api-key header (never URL query param)."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent"
    headers = {
        "x-goog-api-key": api_key,
        "Content-Type": "application/json",
    }
    payload = {
        "systemInstruction": {
            "parts": [{"text": system_prompt}]
        },
        "contents": [
            {"parts": [{"text": prompt}]}
        ],
        "generationConfig": {
            "temperature": 0.4,
            "maxOutputTokens": max_tokens
        }
    }
    async with httpx.AsyncClient(timeout=25.0) as client:
        resp = await client.post(url, headers=headers, json=payload)
        resp.raise_for_status()
        data = resp.json()
        meta = data.get("usageMetadata", {})
        tokens = {
            "prompt": int(meta.get("promptTokenCount", 0)),
            "completion": int(meta.get("candidatesTokenCount", 0)),
            "total": int(meta.get("totalTokenCount", 0)),
        }
        return data["candidates"][0]["content"]["parts"][0]["text"], tokens


async def query_mentor(
    query_text: str,
    system_prompt: str = SYSTEM_PROMPT,
    max_tokens: int = 1500,
) -> Dict[str, Any]:
    """Queries Groq first, falls back to Gemini 2.5 Flash, enforcing ₹0 daily quota caps."""
    query_text = query_text.strip()
    if not query_text:
        return {"status": "error", "message": "Query cannot be empty."}

    load_dotenv(dotenv_path=env_path, override=True)
    groq_key = os.environ.get("GROQ_API_KEY", "").strip()
    gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()

    has_groq = bool(groq_key and not groq_key.startswith("your_"))
    has_gemini = bool(gemini_key and not gemini_key.startswith("your_"))

    if not has_groq and not has_gemini:
        guidance = (
            "Chitran Mentor Sparring Partner is connected, but no active API key was detected.\n\n"
            "To enable real-time sub-second responses, please paste your free Groq or Gemini API key into:\n"
            f"  {env_path}\n\n"
            "Configuration variables:\n"
            "  GROQ_API_KEY=gsk_...\n"
            "  GEMINI_API_KEY=AIza...\n\n"
            "Once saved in .env, Chitran reloads it automatically on your next 'cn ask' trigger."
        )
        return {"status": "success", "model": "offline-guide", "answer": guidance}

    blocked_reasons = []

    # 1. Try Groq (openai/gpt-oss-120b -> qwen/qwen3.8-27b)
    if has_groq:
        q_status = check_quota("groq")
        if q_status["allowed"]:
            for model_name in GROQ_MODELS:
                try:
                    answer, tokens = await query_groq(
                        query_text,
                        groq_key,
                        system_prompt=system_prompt,
                        max_tokens=max_tokens,
                        model_name=model_name,
                    )
                    post_status = record_usage(
                        "groq",
                        model_name,
                        prompt_tokens=tokens["prompt"],
                        completion_tokens=tokens["completion"],
                        total_tokens=tokens["total"] or None,
                    )
                    if post_status["warning"]:
                        notify_windows(
                            "Chitran — Quota Warning",
                            f"Groq free daily usage at {post_status['usage_pct']}%. Hard stop triggers at 95%."
                        )
                        answer = f"[Quota Warning: Groq daily free usage at {post_status['usage_pct']}%]\n\n{answer}"
                    return {
                        "status": "success",
                        "model": f"groq/{model_name}",
                        "quota_pct": post_status["usage_pct"],
                        "answer": answer,
                    }
                except Exception as e:
                    print(f"[Router] Groq {model_name} failed: {e}")
        else:
            blocked_reasons.append(f"Groq ({q_status['usage_pct']}% used)")

    # 2. Try Gemini (gemini-3.8-flash)
    if has_gemini:
        g_status = check_quota("gemini")
        if g_status["allowed"]:
            try:
                answer, tokens = await query_gemini(
                    query_text,
                    gemini_key,
                    system_prompt=system_prompt,
                    max_tokens=max_tokens,
                    model_name=GEMINI_MODEL,
                )
                post_status = record_usage(
                    "gemini",
                    GEMINI_MODEL,
                    prompt_tokens=tokens["prompt"],
                    completion_tokens=tokens["completion"],
                    total_tokens=tokens["total"] or None,
                )
                if post_status["warning"]:
                    notify_windows(
                        "Chitran — Quota Warning",
                        f"Gemini free daily usage at {post_status['usage_pct']}%. Hard stop triggers at 95%."
                    )
                    answer = f"[Quota Warning: Gemini daily free usage at {post_status['usage_pct']}%]\n\n{answer}"
                return {
                    "status": "success",
                    "model": f"google/{GEMINI_MODEL}",
                    "quota_pct": post_status["usage_pct"],
                    "answer": answer,
                }
            except Exception as e:
                print(f"[Router] Gemini {GEMINI_MODEL} failed: {e}")
        else:
            blocked_reasons.append(f"Gemini ({g_status['usage_pct']}% used)")

    # 3. Hard circuit breaker message if caps hit
    if blocked_reasons:
        summary = get_daily_quota_summary()
        msg = (
            "₹0 Quota Circuit Breaker Activated\n\n"
            f"Blocked providers (>=95% daily free cap): {', '.join(blocked_reasons)}.\n"
            "Chitran paused LLM API calls for today to guarantee zero billing or account lockouts.\n\n"
            f"Today's Ledger:\n"
            f"  Groq:   {summary['groq']['requests_used']}/{summary['groq']['requests_limit']} reqs | "
            f"{summary['groq']['tokens_used']}/{summary['groq']['tokens_limit']} tokens\n"
            f"  Gemini: {summary['gemini']['requests_used']}/{summary['gemini']['requests_limit']} reqs | "
            f"{summary['gemini']['tokens_used']}/{summary['gemini']['tokens_limit']} tokens"
        )
        notify_windows("Chitran — Quota Cap Reached", "API calls paused at 95% daily free limit.")
        return {"status": "quota_blocked", "model": "circuit-breaker", "answer": msg}

    return {
        "status": "error",
        "model": "error",
        "answer": "Configured API providers failed to respond. Check your .env keys or internet connection."
    }
