"""
Sidekick OS — Step 6.1: Local SQLite Analytics Aggregator (`daemon/analytics.py`)

Computes Day / Week / Month metrics strictly from local `data/sidekick.db`:
  - Focus vs. distraction vs. neutral timeline (hourly for Day; daily for Week/Month)
  - Top applications & category share
  - Completed & pending goals list
  - Hydration streaks (`water_reminder` acknowledgments), eye breaks & YouTube nudges
  - Live Groq/Gemini free-tier quota gauges + period token totals
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from db import get_connection, init_db
from quota import get_daily_quota_summary

RANGE_DAYS = {
    "day": 1,
    "week": 7,
    "month": 30,
}


def normalize_range(range_key: str) -> str:
    val = (range_key or "day").strip().lower()
    if val in ("day", "today", "1d"):
        return "day"
    if val in ("week", "7d", "weekly"):
        return "week"
    if val in ("month", "30d", "monthly"):
        return "month"
    return "day"


def _resolve_date_window(range_norm: str, reference_date: Optional[str] = None) -> Tuple[date, date]:
    end_dt = date.fromisoformat(reference_date) if reference_date else datetime.now().date()
    span = RANGE_DAYS[range_norm]
    start_dt = end_dt - timedelta(days=span - 1)
    return start_dt, end_dt


def _calculate_hydration_streak(conn: Any, end_dt: date) -> int:
    """Count consecutive days up to end_dt (or yesterday if today not yet logged) with a water ack."""
    rows = conn.execute(
        """
        SELECT DISTINCT date
        FROM habit_events
        WHERE event_type = 'water_reminder'
          AND action_taken IN ('acknowledged', 'completed', 'done')
          AND date <= ?
        ORDER BY date DESC
        """,
        (end_dt.isoformat(),),
    ).fetchall()

    ack_dates = {r["date"] for r in rows}
    if not ack_dates:
        return 0

    cursor = end_dt
    if cursor.isoformat() not in ack_dates:
        yesterday = end_dt - timedelta(days=1)
        if yesterday.isoformat() not in ack_dates:
            return 0
        cursor = yesterday

    streak = 0
    while cursor.isoformat() in ack_dates:
        streak += 1
        cursor -= timedelta(days=1)
    return streak


def _build_timeline(
    conn: Any,
    range_norm: str,
    start_dt: date,
    end_dt: date,
) -> List[Dict[str, Any]]:
    start_str = start_dt.isoformat()
    end_str = end_dt.isoformat()

    if range_norm == "day":
        rows = conn.execute(
            """
            SELECT
                SUBSTR(timestamp, 12, 2) AS bucket,
                category,
                COALESCE(SUM(duration_sec), 0) AS total_sec
            FROM activity_logs
            WHERE date = ?
            GROUP BY bucket, category
            """,
            (end_str,),
        ).fetchall()

        hour_map: Dict[int, Dict[str, int]] = {}
        for r in rows:
            try:
                hr = int(r["bucket"])
            except (TypeError, ValueError):
                continue
            slot = hour_map.setdefault(hr, {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0})
            cat = (r["category"] or "neutral").lower()
            slot[cat if cat in slot else "neutral"] += int(r["total_sec"])

        if hour_map:
            min_hr = min(8, min(hour_map.keys()))
            max_hr = max(21, max(hour_map.keys()))
        else:
            min_hr, max_hr = 8, 21

        timeline = []
        for hr in range(min_hr, max_hr + 1):
            slot = hour_map.get(hr, {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0})
            timeline.append({
                "label": f"{hr:02d}:00",
                "focus_min": round(slot["focus"] / 60.0, 1),
                "distraction_min": round(slot["distraction"] / 60.0, 1),
                "neutral_min": round(slot["neutral"] / 60.0, 1),
            })
        return timeline

    rows = conn.execute(
        """
        SELECT
            date AS bucket,
            category,
            COALESCE(SUM(duration_sec), 0) AS total_sec
        FROM activity_logs
        WHERE date BETWEEN ? AND ?
        GROUP BY date, category
        ORDER BY date ASC
        """,
        (start_str, end_str),
    ).fetchall()

    day_map: Dict[str, Dict[str, int]] = {}
    for r in rows:
        d_str = r["bucket"]
        slot = day_map.setdefault(d_str, {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0})
        cat = (r["category"] or "neutral").lower()
        slot[cat if cat in slot else "neutral"] += int(r["total_sec"])

    timeline = []
    cur = start_dt
    while cur <= end_dt:
        iso = cur.isoformat()
        slot = day_map.get(iso, {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0})
        timeline.append({
            "label": cur.strftime("%b %d") if range_norm == "week" else cur.strftime("%m-%d"),
            "date": iso,
            "focus_min": round(slot["focus"] / 60.0, 1),
            "distraction_min": round(slot["distraction"] / 60.0, 1),
            "neutral_min": round(slot["neutral"] / 60.0, 1),
        })
        cur += timedelta(days=1)
    return timeline


def get_analytics_summary(
    range_key: str = "day",
    db_path: Optional[Path] = None,
    reference_date: Optional[str] = None,
) -> Dict[str, Any]:
    """Return full Day / Week / Month analytics snapshot for the dashboard UI."""
    init_db(db_path)
    range_norm = normalize_range(range_key)
    start_dt, end_dt = _resolve_date_window(range_norm, reference_date=reference_date)
    start_str, end_str = start_dt.isoformat(), end_dt.isoformat()

    with get_connection(db_path) as conn:
        # 1. Category time totals
        cat_rows = conn.execute(
            """
            SELECT category, COALESCE(SUM(duration_sec), 0) AS total_sec
            FROM activity_logs
            WHERE date BETWEEN ? AND ?
            GROUP BY category
            """,
            (start_str, end_str),
        ).fetchall()
        cat_sec = {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0}
        for r in cat_rows:
            c = (r["category"] or "neutral").lower()
            cat_sec[c if c in cat_sec else "neutral"] += int(r["total_sec"])

        focus_min = round(cat_sec["focus"] / 60.0, 1)
        distract_min = round(cat_sec["distraction"] / 60.0, 1)
        neutral_min = round(cat_sec["neutral"] / 60.0, 1)
        idle_min = round(cat_sec["idle"] / 60.0, 1)
        active_fd = cat_sec["focus"] + cat_sec["distraction"]
        focus_score = round((cat_sec["focus"] / active_fd) * 100.0, 1) if active_fd > 0 else 0.0

        # 2. Timeline buckets
        timeline = _build_timeline(conn, range_norm, start_dt, end_dt)

        # 3. Top applications breakdown (excluding pure idle)
        app_rows = conn.execute(
            """
            SELECT
                app_name,
                category,
                COALESCE(SUM(duration_sec), 0) AS total_sec
            FROM activity_logs
            WHERE date BETWEEN ? AND ? AND category != 'idle'
            GROUP BY app_name, category
            ORDER BY total_sec DESC
            """,
            (start_str, end_str),
        ).fetchall()

        app_agg: Dict[str, Dict[str, Any]] = {}
        total_active_app_sec = 0
        for r in app_rows:
            name = r["app_name"] or "Unknown"
            sec = int(r["total_sec"])
            cat = (r["category"] or "neutral").lower()
            total_active_app_sec += sec
            entry = app_agg.setdefault(name, {"app_name": name, "total_sec": 0, "by_cat": {}})
            entry["total_sec"] += sec
            entry["by_cat"][cat] = entry["by_cat"].get(cat, 0) + sec

        app_breakdown = []
        for entry in sorted(app_agg.values(), key=lambda x: x["total_sec"], reverse=True)[:8]:
            dom_cat = max(entry["by_cat"].items(), key=lambda kv: kv[1])[0]
            app_breakdown.append({
                "app_name": entry["app_name"],
                "minutes": round(entry["total_sec"] / 60.0, 1),
                "category": dom_cat,
                "share_pct": round((entry["total_sec"] / total_active_app_sec) * 100.0, 1)
                if total_active_app_sec > 0
                else 0.0,
            })

        # 4. Goals in period
        goal_rows = conn.execute(
            """
            SELECT id, date, title, deadline_time, status, created_at, completed_at
            FROM goals
            WHERE date BETWEEN ? AND ?
            ORDER BY date DESC, id DESC
            """,
            (start_str, end_str),
        ).fetchall()
        goals = [dict(g) for g in goal_rows]
        goals_completed = sum(1 for g in goals if g["status"] == "completed")
        goals_total = len(goals)
        completion_rate = round((goals_completed / goals_total) * 100.0, 1) if goals_total > 0 else 0.0

        # 5. Habits & hydration streak
        habit_rows = conn.execute(
            """
            SELECT event_type, action_taken, COUNT(*) AS cnt
            FROM habit_events
            WHERE date BETWEEN ? AND ?
            GROUP BY event_type, action_taken
            """,
            (start_str, end_str),
        ).fetchall()
        habit_counts: Dict[str, Dict[str, int]] = {}
        for r in habit_rows:
            et = r["event_type"]
            act = r["action_taken"]
            habit_counts.setdefault(et, {})[act] = int(r["cnt"])

        water_map = habit_counts.get("water_reminder", {})
        water_acks = water_map.get("acknowledged", 0) + water_map.get("completed", 0) + water_map.get("done", 0)
        water_triggers = sum(water_map.values())
        eye_map = habit_counts.get("eye_break", {})
        eye_acks = eye_map.get("acknowledged", 0) + eye_map.get("completed", 0) + eye_map.get("done", 0)
        eye_triggers = sum(eye_map.values())
        yt_nudges = sum(habit_counts.get("youtube_nudge", {}).values()) + sum(
            habit_counts.get("distraction_nudge", {}).values()
        )
        mail_briefings = sum(habit_counts.get("gmail_briefing", {}).values())
        hydration_streak = _calculate_hydration_streak(conn, end_dt)

        # 6. Period token summary from quota_ledger
        period_q_rows = conn.execute(
            """
            SELECT
                provider,
                COUNT(*) AS requests,
                COALESCE(SUM(total_tokens), 0) AS tokens
            FROM quota_ledger
            WHERE date BETWEEN ? AND ?
            GROUP BY provider
            """,
            (start_str, end_str),
        ).fetchall()
        period_quota = {
            r["provider"]: {"requests": int(r["requests"]), "tokens": int(r["tokens"])}
            for r in period_q_rows
        }

    daily_quota = get_daily_quota_summary(db_path=db_path)

    return {
        "range": range_norm,
        "start_date": start_str,
        "end_date": end_str,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "summary": {
            "focus_minutes": focus_min,
            "distraction_minutes": distract_min,
            "neutral_minutes": neutral_min,
            "idle_minutes": idle_min,
            "focus_score_pct": focus_score,
            "goals_completed": goals_completed,
            "goals_total": goals_total,
            "completion_rate_pct": completion_rate,
            "hydration_acks": water_acks,
            "hydration_triggers": water_triggers,
            "hydration_streak_days": hydration_streak,
            "eye_breaks_acks": eye_acks,
            "eye_breaks_triggers": eye_triggers,
            "youtube_nudges": yt_nudges,
            "mail_briefings": mail_briefings,
        },
        "timeline": timeline,
        "app_breakdown": app_breakdown,
        "goals": goals[:25],
        "quota": {
            "daily": daily_quota,
            "period": period_quota,
        },
    }

