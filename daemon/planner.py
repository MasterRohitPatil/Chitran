import re
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple

from db import get_connection, init_db
from state import is_paused, set_mascot_state

TIME_TOKEN_RE = re.compile(
    r"^\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\s*$",
    re.IGNORECASE,
)


def today_iso() -> str:
    return date.today().isoformat()


def normalize_time_str(expr: str) -> Optional[str]:
    """Normalizes '21:30', '9:30 pm', or '9pm' into 24-hour 'HH:MM'."""
    if not expr:
        return None
    m = TIME_TOKEN_RE.match(expr.strip())
    if not m:
        return None

    hour = int(m.group(1))
    minute = int(m.group(2) or "0")
    meridiem = (m.group(3) or "").lower()

    if minute < 0 or minute > 59:
        return None

    if meridiem:
        if hour < 1 or hour > 12:
            return None
        if meridiem == "am":
            hour = 0 if hour == 12 else hour
        elif meridiem == "pm":
            hour = 12 if hour == 12 else hour + 12
    else:
        if hour < 0 or hour > 23:
            return None

    return f"{hour:02d}:{minute:02d}"


def parse_goal_input(raw: str) -> Tuple[str, Optional[str]]:
    """
    Parses '<task> @ <HH:MM>' (or trailing 'at/by <time>').
    Returns (clean_title, normalized_HH_MM_or_None).
    """
    text = (raw or "").strip()
    if not text:
        return "", None

    # 1. Explicit '@' separator takes highest precedence
    at_match = re.search(r"@\s*([^@]+)$", text)
    if at_match:
        candidate_time = at_match.group(1).strip()
        normalized = normalize_time_str(candidate_time)
        clean_title = text[: at_match.start()].strip(" -:,")
        return clean_title, normalized

    # 2. Fallback: trailing 'by <time>' or 'at <time>' ONLY if it parses as a valid time
    trailing_match = re.search(
        r"\s+(?:by|at)\s+(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)\s*$",
        text,
        re.IGNORECASE,
    )
    if trailing_match:
        normalized = normalize_time_str(trailing_match.group(1))
        if normalized:
            clean_title = text[: trailing_match.start()].strip(" -:,")
            return clean_title, normalized

    return text, None


def add_goal(
    title: str,
    deadline_time: str,
    target_date: Optional[str] = None,
) -> Dict[str, Any]:
    """Inserts a new daily goal into SQLite and returns the created record."""
    clean_title = title.strip()
    if not clean_title:
        raise ValueError("Goal title cannot be empty.")

    normalized_time = normalize_time_str(deadline_time)
    if not normalized_time:
        raise ValueError(f"Invalid deadline time '{deadline_time}'. Use HH:MM (e.g. 21:30 or 9:30 pm).")

    init_db()
    day = target_date or today_iso()
    now_iso = datetime.now().isoformat(timespec="seconds")

    with get_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO goals (date, title, deadline_time, status, created_at)
            VALUES (?, ?, ?, 'pending', ?)
            """,
            (day, clean_title, normalized_time, now_iso),
        )
        goal_id = cur.lastrowid

    return {
        "id": goal_id,
        "date": day,
        "title": clean_title,
        "deadline_time": normalized_time,
        "status": "pending",
        "created_at": now_iso,
        "completed_at": None,
    }


def get_today_goals(target_date: Optional[str] = None) -> List[Dict[str, Any]]:
    """Returns all goals for the given date ordered by status (pending first) and deadline."""
    init_db()
    day = target_date or today_iso()
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, date, title, deadline_time, status, created_at, completed_at
            FROM goals
            WHERE date = ?
            ORDER BY
                CASE status WHEN 'pending' THEN 0 ELSE 1 END,
                COALESCE(deadline_time, '23:59') ASC,
                id ASC
            """,
            (day,),
        ).fetchall()
    return [dict(r) for r in rows]


def get_pending_goals(target_date: Optional[str] = None) -> List[Dict[str, Any]]:
    return [g for g in get_today_goals(target_date) if g["status"] == "pending"]


def has_goals_today(target_date: Optional[str] = None) -> bool:
    init_db()
    day = target_date or today_iso()
    with get_connection() as conn:
        row = conn.execute("SELECT COUNT(1) AS cnt FROM goals WHERE date = ?", (day,)).fetchone()
    return bool(row and row["cnt"] > 0)


def refresh_mascot_goal_state(target_date: Optional[str] = None) -> List[Dict[str, Any]]:
    """Checks for overdue pending goals today and updates Byte's state unless paused."""
    now_hhmm = datetime.now().strftime("%H:%M")
    pending = get_pending_goals(target_date)
    overdue = [
        g for g in pending
        if g.get("deadline_time") and g["deadline_time"] < now_hhmm
    ]
    if not is_paused():
        set_mascot_state("alert" if overdue else "focused")
    return overdue


def mark_goal_done(
    identifier: str = "",
    target_date: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Marks a pending goal completed by numeric ID ('3' / '#3'), keyword substring,
    or auto-selects when exactly one pending goal exists today.
    """
    init_db()
    day = target_date or today_iso()
    ident = (identifier or "").strip()
    now_iso = datetime.now().isoformat(timespec="seconds")

    pending = get_pending_goals(day)
    if not pending:
        return {
            "status": "empty",
            "message": "No pending goals remaining for today.",
            "goals": get_today_goals(day),
        }

    matched: Optional[Dict[str, Any]] = None
    id_match = re.match(r"^#?(\d+)$", ident)

    if not ident:
        if len(pending) == 1:
            matched = pending[0]
        else:
            return {
                "status": "ambiguous",
                "message": f"{len(pending)} pending goals found. Pick one in the Planner HUD or run 'cn done <id>'.",
                "goals": pending,
            }
    elif id_match:
        target_id = int(id_match.group(1))
        for g in pending:
            if g["id"] == target_id:
                matched = g
                break
        if not matched:
            return {
                "status": "not_found",
                "message": f"No pending goal with ID #{target_id} found today.",
                "goals": get_today_goals(day),
            }
    else:
        needle = ident.lower()
        matches = [g for g in pending if needle in g["title"].lower()]
        if len(matches) == 1:
            matched = matches[0]
        elif len(matches) > 1:
            return {
                "status": "ambiguous",
                "message": f"Multiple goals match '{ident}'. Pick by ID in the Planner HUD.",
                "goals": matches,
            }
        else:
            return {
                "status": "not_found",
                "message": f"No pending goal matching '{ident}' found today.",
                "goals": get_today_goals(day),
            }

    with get_connection() as conn:
        conn.execute(
            """
            UPDATE goals
            SET status = 'completed', completed_at = ?
            WHERE id = ?
            """,
            (now_iso, matched["id"]),
        )

    matched["status"] = "completed"
    matched["completed_at"] = now_iso
    refresh_mascot_goal_state(day)

    remaining_pending = len(get_pending_goals(day))
    return {
        "status": "success",
        "goal": matched,
        "remaining_pending": remaining_pending,
        "goals": get_today_goals(day),
        "message": f"Completed #{matched['id']}: {matched['title']}",
    }


def handle_goal_command(raw_arg: str) -> Dict[str, Any]:
    """Routes `sk goal` commands from daemon /command endpoint."""
    raw = (raw_arg or "").strip()
    if not raw or raw.lower() in ("list", "show", "today", "view"):
        goals = get_today_goals()
        return {
            "status": "open_planner",
            "goals": goals,
            "message": f"{len(goals)} goal(s) logged for today.",
        }

    title, deadline = parse_goal_input(raw)
    if not title:
        return {
            "status": "open_planner",
            "goals": get_today_goals(),
            "message": "Enter a goal title along with a deadline.",
        }

    if not deadline:
        return {
            "status": "needs_deadline",
            "title": title,
            "goals": get_today_goals(),
            "message": f"Deadline missing for '{title}'. Opening inline HUD prompt.",
        }

    goal = add_goal(title, deadline)
    refresh_mascot_goal_state()
    return {
        "status": "success",
        "goal": goal,
        "goals": get_today_goals(),
        "message": f"Goal #{goal['id']} locked: {goal['title']} @ {goal['deadline_time']}",
    }

