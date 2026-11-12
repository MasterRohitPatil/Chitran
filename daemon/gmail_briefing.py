"""
Sidekick OS — Step 4.4: Twice-Daily Smart Gmail Briefing (`sk mail`)

Reads today's received emails for `rohit.newchallenger@gmail.com` silently via
Chrome DevTools Protocol (`BrowserManager` on port 9222, background tab, zero
stolen focus), summarizes important items cleanly using Groq/Gemini (or a zero-cost
deterministic fallback when offline / quota-capped), speaks the summary through
headphones (`voice.py`), shifts Byte into `mail` (Mint-Teal envelope) mode, and
displays the Floating Dark Cyber-HUD (`hud.py`).

Smart Schedule Rules:
  - Morning slot (10:00 AM+): triggers automatically once per day at 10:00 AM, or
    immediately on the first active laptop wake-up after 10:00 AM.
  - Evening slot (10:30 PM+): triggers automatically once per day after 10:30 PM
    ONLY when the user is actively at the laptop (`idle_sec < 180`).
  - Zero duplicate runs per day across daemon restarts (`data/mail_briefing_state.json`).
"""

from __future__ import annotations

import asyncio
import json
import re
from datetime import datetime, time as dt_time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from browser import browser_manager, is_cdp_alive
from db import get_connection, init_db
from hud import spawn_hud_process
from monitor import get_idle_seconds, hold_mascot_state
from planner import today_iso
from router import query_mentor
from state import is_paused, load_system_state
from voice import speak_background


PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATE_FILE = PROJECT_ROOT / "data" / "mail_briefing_state.json"

TARGET_EMAIL = "rohit.newchallenger@gmail.com"
MORNING_TIME = dt_time(10, 0)
EVENING_TIME = dt_time(22, 30)
ACTIVE_IDLE_LIMIT_SEC = 180.0
SCHEDULER_POLL_SEC = 60

URGENT_KEYWORDS = (
    "interview",
    "shortlist",
    "assessment",
    "coding test",
    "hackerrank",
    "assignment",
    "offer",
    "internship",
    "application",
    "deadline",
    "urgent",
    "action required",
    "verify",
    "security alert",
    "otp",
    "meeting",
    "schedule",
    "exam",
    "sies",
    "recruiter",
    "hr",
)


# ---------------------------------------------------------------------------
# Schedule State Persistence (prevents duplicate briefings per day)
# ---------------------------------------------------------------------------

def load_briefing_state(state_path: Optional[Path] = None) -> Dict[str, Any]:
    path = state_path or STATE_FILE
    today = today_iso()
    default_state: Dict[str, Any] = {
        "date": today,
        "target_email": TARGET_EMAIL,
        "morning_done": False,
        "evening_done": False,
        "last_run_iso": None,
        "last_slot": None,
        "last_status": None,
        "last_today_count": 0,
        "last_spoken": "",
    }
    if not path.exists():
        return default_state
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("date") != today:
            return default_state
        return {**default_state, **data}
    except Exception:
        return default_state


def save_briefing_state(data: Dict[str, Any], state_path: Optional[Path] = None) -> None:
    path = state_path or STATE_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def determine_due_slot(
    now_dt: Optional[datetime] = None,
    idle_sec: Optional[float] = None,
    state_data: Optional[Dict[str, Any]] = None,
) -> Optional[str]:
    """
    Decide whether 'morning' (10:00 AM+ wake/active) or 'evening' (10:30 PM+ active)
    briefing should fire right now, or None if not due.
    """
    if is_paused():
        return None

    now = now_dt or datetime.now()
    idle = float(idle_sec if idle_sec is not None else get_idle_seconds())
    st = state_data if state_data is not None else load_briefing_state()
    if st.get("date") != now.date().isoformat():
        st = {
            "date": now.date().isoformat(),
            "morning_done": False,
            "evening_done": False,
        }

    cur_time = now.time()

    # 10:30 PM+: user must be actively at the laptop
    if cur_time >= EVENING_TIME:
        if not st.get("evening_done", False) and idle < ACTIVE_IDLE_LIMIT_SEC:
            return "evening"
        return None

    # 10:00 AM .. 10:29 PM: fires at 10:00 AM or first wake-up/activity after 10:00 AM
    if cur_time >= MORNING_TIME:
        if not st.get("morning_done", False) and idle < 600.0:
            return "morning"

    return None


# ---------------------------------------------------------------------------
# Pure Parsing, Priority Classification & Summarization Helpers
# ---------------------------------------------------------------------------

def is_email_from_today(
    time_text: str,
    time_title: str = "",
    ref_dt: Optional[datetime] = None,
) -> bool:
    """Detect whether a Gmail inbox row's timestamp represents today's date."""
    now = ref_dt or datetime.now()
    txt = (time_text or "").strip()
    title = (time_title or "").strip()

    # Gmail displays clock-only strings ("9:45 AM", "14:20", "00:08") strictly for today's mail
    if re.match(r"^\d{1,2}:\d{2}(?:\s*[AP]M)?$", txt, re.IGNORECASE):
        return True

    if title:
        month_abbr = now.strftime("%b").lower()
        month_full = now.strftime("%B").lower()
        day_no_zero = str(now.day)
        day_padded = f"{now.day:02d}"
        year_str = str(now.year)
        t_low = title.lower()
        date_iso = now.date().isoformat()
        if date_iso in t_low:
            return True
        if year_str in t_low and (month_abbr in t_low or month_full in t_low):
            if re.search(rf"\b(?:{day_no_zero}|{day_padded})\b", t_low):
                return True

    return False


def is_urgent_email(email_item: Dict[str, Any]) -> bool:
    blob = f"{email_item.get('sender', '')} {email_item.get('subject', '')} {email_item.get('snippet', '')}".lower()
    return any(kw in blob for kw in URGENT_KEYWORDS)


def format_deterministic_briefing(
    emails: List[Dict[str, Any]],
    slot: str = "manual",
    target_email: str = TARGET_EMAIL,
    unread_total: int = 0,
) -> Tuple[str, str]:
    """
    Build `(spoken_summary, markdown_hud)` with zero LLM cost.
    Never fails even if offline or daily free LLM quota is capped.
    """
    slot_title = {
        "morning": "Morning Gmail Briefing (10:00 AM)",
        "evening": "Nightly Gmail Briefing (10:30 PM)",
    }.get(slot, "Live Gmail Briefing")

    count = len(emails)
    if count == 0:
        spoken = (
            f"Gmail check complete for {target_email}. No new emails arrived today"
            + (f", though {unread_total} older unread threads remain in your inbox." if unread_total else ".")
        )
        md = (
            f"## {slot_title}\n"
            f"**Account:** `{target_email}`\n\n"
            f"- **Received Today:** 0 emails\n"
            + (f"- **Unread Inbox Threads (older):** {unread_total}\n" if unread_total else "")
            + "\nAll caught up for today."
        )
        return spoken, md

    urgent = [e for e in emails if is_urgent_email(e)]
    normal = [e for e in emails if not is_urgent_email(e)]
    unread_today = sum(1 for e in emails if e.get("unread"))

    lines = [
        f"## {slot_title}",
        f"**Account:** `{target_email}` · **Today:** {count} email(s) ({unread_today} unread)\n",
    ]

    if urgent:
        lines.append("### Priority / Actionable")
        for item in urgent[:6]:
            mark = "**[UNREAD]** " if item.get("unread") else ""
            time_str = item.get("time_text") or ""
            subj = item.get("subject") or "(No Subject)"
            sender = item.get("sender") or "Unknown Sender"
            snip = (item.get("snippet") or "").strip()
            lines.append(f"- {mark}**{subj}** — *{sender}* ({time_str})")
            if snip:
                lines.append(f"  > {snip[:140]}")
        lines.append("")

    if normal:
        lines.append("### Other Received Today")
        for item in normal[:10]:
            mark = "**[UNREAD]** " if item.get("unread") else ""
            time_str = item.get("time_text") or ""
            subj = item.get("subject") or "(No Subject)"
            sender = item.get("sender") or "Unknown Sender"
            lines.append(f"- {mark}{subj} — *{sender}* ({time_str})")

    top_highlights = urgent[:2] if urgent else emails[:2]
    spoken_parts = [f"Rohit, you received {count} email{'s' if count != 1 else ''} today"]
    if unread_today:
        spoken_parts[0] += f", with {unread_today} unread"
    spoken_parts[0] += "."
    if urgent:
        spoken_parts.append(f"{len(urgent)} look{'s' if len(urgent) == 1 else ''} high priority.")
    for item in top_highlights:
        spoken_parts.append(f"From {item.get('sender', 'someone')}: {item.get('subject', 'no subject')}.")

    return " ".join(spoken_parts), "\n".join(lines)


async def summarize_emails_smart(
    emails: List[Dict[str, Any]],
    slot: str = "manual",
    target_email: str = TARGET_EMAIL,
    unread_total: int = 0,
) -> Tuple[str, str, str]:
    """
    Return `(spoken_summary, markdown_hud, model_used)`.
    Uses Groq/Gemini free tier within quota caps for crisp natural briefing,
    falling back deterministically to zero-LLM formatter if unavailable.
    """
    fallback_spoken, fallback_md = format_deterministic_briefing(
        emails, slot=slot, target_email=target_email, unread_total=unread_total
    )
    if not emails:
        return fallback_spoken, fallback_md, "deterministic"

    compact_rows = [
        {
            "sender": e.get("sender", ""),
            "subject": e.get("subject", ""),
            "snippet": (e.get("snippet", "") or "")[:120],
            "time": e.get("time_text", ""),
            "unread": bool(e.get("unread")),
        }
        for e in emails[:18]
    ]
    prompt = (
        f"Summarize today's Gmail inbox ({len(emails)} emails for {target_email}, slot={slot}).\n"
        "Return strict JSON with two string keys ONLY:\n"
        '{"spoken": "2 short natural sentences to read aloud in headphones (mention urgent items or interviews first)", '
        '"markdown": "Clean markdown bullets grouped into ### Priority / Actionable and ### Updates"}\n\n'
        f"Emails JSON:\n{json.dumps(compact_rows, ensure_ascii=False)}"
    )

    mail_sys = (
        "You summarize Rohit's Gmail inbox concisely. Highlight placement drives, "
        "interviews, deadlines, and security alerts first. Output strict valid JSON."
    )
    raw_answer: Optional[str] = None
    model_used = "deterministic"

    try:
        llm_res = await query_mentor(prompt, system_prompt=mail_sys, max_tokens=700)
        if llm_res.get("status") == "success" and llm_res.get("model") != "offline-guide":
            raw_answer = llm_res.get("answer")
            model_used = str(llm_res.get("model") or "llm")
    except Exception as e:
        print(f"[Mail] LLM summary fallback ({e})")

    if raw_answer:
        try:
            m = re.search(r"\{.*\}", raw_answer, re.DOTALL)
            parsed = json.loads(m.group(0) if m else raw_answer)
            spoken = str(parsed.get("spoken") or "").strip() or fallback_spoken
            md = str(parsed.get("markdown") or "").strip() or fallback_md
            return spoken, md, model_used
        except Exception:
            # Treat plain markdown output gracefully
            return fallback_spoken, raw_answer.strip(), model_used

    return fallback_spoken, fallback_md, "deterministic"


# ---------------------------------------------------------------------------
# Chrome CDP Silent Gmail Reader (port 9222)
# ---------------------------------------------------------------------------

_EXTRACT_GMAIL_JS = r"""
() => {
    const url = window.location.href || "";
    const bodyText = (document.body && document.body.innerText) ? document.body.innerText : "";
    const accountNodes = Array.from(document.querySelectorAll('a[aria-label*="Google Account"], a[aria-label*="@gmail.com"], [aria-label*="@gmail.com"]'));
    let activeAccount = "";
    for (const el of accountNodes) {
        const lbl = el.getAttribute("aria-label") || "";
        const m = lbl.match(/[a-zA-Z0-9._%+-]+@gmail\.com/i);
        if (m) { activeAccount = m[0].toLowerCase(); break; }
    }

    const signedOut = url.includes("accounts.google.com")
        || (!url.includes("mail.google.com/mail"))
        || (bodyText.includes("Sign in") && bodyText.includes("Create account"));

    const rows = Array.from(document.querySelectorAll("tr.zA"));
    const items = [];
    for (const row of rows.slice(0, 30)) {
        const senderEl = row.querySelector(".yW span[email], .yW span[name], .yW");
        const sender = senderEl
            ? ((senderEl.getAttribute("name") || senderEl.innerText || senderEl.getAttribute("email") || "").trim())
            : "";
        const senderEmail = senderEl ? (senderEl.getAttribute("email") || "").trim() : "";

        const subjEl = row.querySelector(".bog, .y6 span[id], .y6");
        const subject = subjEl ? subjEl.innerText.trim() : "";

        const snipEl = row.querySelector(".y2");
        let snippet = snipEl ? snipEl.innerText.trim() : "";
        snippet = snippet.replace(/^\s*[-–—]\s*/, "");

        const timeEl = row.querySelector(".xW span[title], .xW span[aria-label], .xW span");
        const timeText = timeEl ? timeEl.innerText.trim() : "";
        const timeTitle = timeEl ? ((timeEl.getAttribute("title") || timeEl.getAttribute("aria-label") || "").trim()) : "";

        const unread = row.classList.contains("zE");
        if (subject || sender) {
            items.push({
                sender: sender || senderEmail || "Unknown",
                sender_email: senderEmail,
                subject: subject || "(No Subject)",
                snippet,
                time_text: timeText,
                time_title: timeTitle,
                unread
            });
        }
    }

    return {
        url,
        signed_out: signedOut,
        active_account: activeAccount,
        rows: items
    };
}
"""


async def fetch_today_emails_via_cdp(
    target_email: str = TARGET_EMAIL,
    auto_launch_chrome: bool = True,
) -> Dict[str, Any]:
    """
    Open Gmail in a silent background CDP tab (`background=True`), inspect today's
    received emails for `target_email`, and close the temporary tab cleanly.
    """
    if not auto_launch_chrome and not await is_cdp_alive(timeout_sec=1.0):
        return {
            "status": "chrome_offline",
            "message": "Sidekick Chrome (CDP port 9222) is not open yet.",
            "today_emails": [],
            "unread_total": 0,
        }

    # Check both /u/0/ and fallback /u/1/ only if primary account belongs to another Gmail
    for slot_idx in (0, 1):
        gmail_url = f"https://mail.google.com/mail/u/{slot_idx}/#inbox"
        page = None
        try:
            page, _ = await browser_manager.create_background_tab(gmail_url)
            try:
                await page.wait_for_selector("tr.zA, [role='main'], input[type='email']", timeout=10000)
            except Exception:
                await asyncio.sleep(1.5)

            data = await page.evaluate(_EXTRACT_GMAIL_JS)
        except Exception as e:
            if page:
                try:
                    await page.close()
                except Exception:
                    pass
            return {
                "status": "error",
                "message": f"Chrome CDP could not load Gmail inbox: {e}",
                "today_emails": [],
                "unread_total": 0,
            }

        try:
            await page.close()
        except Exception:
            pass

        if data.get("signed_out"):
            return {
                "status": "needs_signin",
                "message": (
                    f"Please sign in to {target_email} inside Sidekick Chrome (port 9222) once "
                    "so background Gmail briefings stay authenticated."
                ),
                "today_emails": [],
                "unread_total": 0,
            }

        active_acct = (data.get("active_account") or "").lower()
        if active_acct and target_email.lower() not in active_acct and slot_idx == 0:
            # Might be logged into multiple Google accounts; check /u/1/ next
            continue

        all_rows = data.get("rows") or []
        now_dt = datetime.now()
        today_rows = [
            row for row in all_rows
            if is_email_from_today(row.get("time_text", ""), row.get("time_title", ""), ref_dt=now_dt)
        ]
        unread_total = sum(1 for row in all_rows if row.get("unread"))
        return {
            "status": "success",
            "account": active_acct or target_email,
            "today_emails": today_rows,
            "recent_rows": all_rows[:10],
            "unread_total": unread_total,
        }

    return {
        "status": "needs_signin",
        "message": f"Could not locate signed-in Gmail session for {target_email} in Sidekick Chrome.",
        "today_emails": [],
        "unread_total": 0,
    }


def _record_mail_habit(action_taken: str, day: Optional[str] = None) -> None:
    init_db()
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO habit_events (date, timestamp, event_type, action_taken) VALUES (?, ?, ?, ?)",
            (day or today_iso(), datetime.now().isoformat(timespec="seconds"), "gmail_briefing", action_taken),
        )


# ---------------------------------------------------------------------------
# End-to-End Orchestrator & Daemon Background Scheduler
# ---------------------------------------------------------------------------

async def run_gmail_briefing(
    trigger: str = "manual",
    arg: str = "",
    spawn_hud: bool = True,
) -> Dict[str, Any]:
    """
    Fetch today's Gmail messages via background CDP, summarize, update Byte (`mail`),
    speak via headphones, and optionally spawn the Dark Floating HUD.
    """
    if trigger != "manual" and is_paused():
        return {"status": "paused", "message": "Sidekick monitoring is paused."}

    auto_launch = trigger == "manual"
    hold_mascot_state("mail", duration_sec=40.0)

    fetched = await fetch_today_emails_via_cdp(
        target_email=TARGET_EMAIL,
        auto_launch_chrome=auto_launch,
    )
    status = fetched.get("status")

    if status != "success":
        msg = fetched.get("message", "Unable to read Gmail right now.")
        md = (
            "## Gmail Briefing — Sign-In / Browser Required\n\n"
            f"**Target Account:** `{TARGET_EMAIL}`\n\n"
            f"{msg}\n\n"
            "### Quick Fix\n"
            "1. Run `launch_chrome_cdp.bat` (or press `Alt + Space` → `cn mail`).\n"
            f"2. Log into **{TARGET_EMAIL}** once inside Sidekick Chrome.\n"
            "3. Re-run **`cn mail`** anytime for instant briefing."
        )
        if trigger == "manual":
            speak_background(msg)
        if spawn_hud and trigger != "manual":
            spawn_hud_process("Gmail Briefing", TARGET_EMAIL, md)
        return {
            "status": status,
            "message": msg,
            "markdown": md,
            "spoken": msg,
            "today_count": 0,
        }

    today_emails = fetched.get("today_emails", [])
    unread_total = int(fetched.get("unread_total", 0))
    spoken, md, model_used = await summarize_emails_smart(
        today_emails,
        slot=trigger,
        target_email=TARGET_EMAIL,
        unread_total=unread_total,
    )

    # Persist daily briefing state so 10:00 AM / 10:30 PM never double-trigger
    st = load_briefing_state()
    now_iso = datetime.now().isoformat(timespec="seconds")
    if trigger == "morning":
        st["morning_done"] = True
    elif trigger == "evening":
        st["evening_done"] = True
        if datetime.now().time() >= EVENING_TIME:
            st["morning_done"] = True
    st["last_run_iso"] = now_iso
    st["last_slot"] = trigger
    st["last_status"] = "success"
    st["last_today_count"] = len(today_emails)
    st["last_spoken"] = spoken
    save_briefing_state(st)

    _record_mail_habit(f"{trigger}:{len(today_emails)}_emails")
    speak_background(spoken)
    if spawn_hud:
        spawn_hud_process(
            f"Gmail Briefing ({len(today_emails)} today)",
            f"{TARGET_EMAIL} · {model_used}",
            md,
        )

    return {
        "status": "success",
        "slot": trigger,
        "account": TARGET_EMAIL,
        "today_count": len(today_emails),
        "unread_total": unread_total,
        "model": model_used,
        "spoken": spoken,
        "markdown": md,
    }


def get_mail_briefing_status() -> Dict[str, Any]:
    """Return current twice-daily Gmail briefing state for API / dashboard visibility."""
    st = load_briefing_state()
    due = determine_due_slot(state_data=st)
    return {
        **st,
        "due_slot_now": due,
        "schedule": {"morning": "10:00+", "evening_active": "22:30+"},
    }


async def mail_scheduler_loop() -> None:
    """Checks every 60s whether the 10:00 AM wake-up or 10:30 PM active briefing is due."""
    print(f"[Mail] Smart Gmail scheduler online for {TARGET_EMAIL} (10:00 AM wake / 10:30 PM active).")
    # Small startup delay so server finishes binding before opening CDP tabs
    await asyncio.sleep(6)
    while True:
        try:
            due_slot = determine_due_slot()
            if due_slot:
                print(f"[Mail] Triggering scheduled {due_slot} Gmail briefing...")
                res = await run_gmail_briefing(trigger=due_slot, spawn_hud=True)
                # If Chrome isn't open/signed-in yet during scheduled check, mark attempted or retry next boot?
                # If chrome_offline, do NOT burn the slot; wait until user opens Chrome or triggers sk mail!
                if res.get("status") in ("chrome_offline", "needs_signin"):
                    print(f"[Mail] Scheduled {due_slot} skipped ({res.get('status')}): will retry once Chrome CDP is signed in.")
                    # Wait longer before next retry if Chrome offline
                    await asyncio.sleep(300)
        except asyncio.CancelledError:
            print("[Mail] Smart Gmail scheduler stopped.")
            raise
        except Exception as e:
            print(f"[Mail] Scheduler loop error (recovered): {e}")
        await asyncio.sleep(SCHEDULER_POLL_SEC)


if __name__ == "__main__":
    print(json.dumps(get_mail_briefing_status(), indent=2))

