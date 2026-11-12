"""
Sidekick OS — Step 4.3: Local Active Window Monitor, YouTube Classifier & Break Nudges

100% local privacy: window titles and activity categories are stored strictly inside
`data/sidekick.db` (`activity_logs`, `habit_events`) and are never uploaded to any cloud.

Core responsibilities:
  - Poll Win32 `GetForegroundWindow` + `GetLastInputInfo` every 15 seconds.
  - Classify foreground activity locally (`focus`, `distraction`, `neutral`, `idle`),
    using a context-aware study vs. distraction classifier for YouTube tabs.
  - Fire smart, cooldown-protected voice + non-intrusive Cyber-HUD nudges:
      * Accountability nudge when drifting on non-study YouTube while a goal is pending
      * 45-minute continuous work Hydration reminder (`hydration` Cyber-Slime)
      * 60-minute continuous work 20-20-20 Eye-Break reminder
  - Auto-reset break timers cleanly when the PC is idle for >= 5 minutes, and
    respect `sk pause` / `sk mute` at all times.
"""

from __future__ import annotations

import asyncio
import ctypes
import ctypes.wintypes as wintypes
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple

# Ensure windowless pythonw.exe subprocesses never fail on None stdout/stderr
if sys.stdout is None or sys.stderr is None:
    _log_dir = Path(__file__).resolve().parent.parent / "data"
    _log_dir.mkdir(parents=True, exist_ok=True)
    _log_file = open(_log_dir / "daemon.log", "a", encoding="utf-8", buffering=1)
    if sys.stdout is None:
        sys.stdout = _log_file
    if sys.stderr is None:
        sys.stderr = _log_file

from db import get_connection, init_db
from planner import get_pending_goals, refresh_mascot_goal_state, today_iso
from state import is_paused, load_system_state, set_mascot_state
from voice import speak_background


# ---------------------------------------------------------------------------
# Timing & Cooldown Constants
# ---------------------------------------------------------------------------

POLL_INTERVAL_SEC = 15
IDLE_PAUSE_SEC = 180          # 3 mins without input -> pause continuous work timers
IDLE_RESET_SEC = 300          # 5 mins away -> natural break, auto-reset break timers
YOUTUBE_NUDGE_STREAK_SEC = 120
DISTRACTION_NUDGE_STREAK_SEC = 180
DISTRACTION_NUDGE_COOLDOWN_SEC = 600
WATER_REMINDER_SEC = 45 * 60
EYE_BREAK_SEC = 60 * 60
HYDRATION_MASCOT_HOLD_SEC = 45


# ---------------------------------------------------------------------------
# Deterministic Local App & YouTube Classifier Rules
# ---------------------------------------------------------------------------

BROWSER_APPS = {
    "chrome.exe",
    "msedge.exe",
    "brave.exe",
    "firefox.exe",
    "opera.exe",
    "vivaldi.exe",
    "arc.exe",
}

FOCUS_APPS = {
    "code.exe",
    "cursor.exe",
    "antigravity.exe",
    "antigravity ide.exe",
    "pycharm64.exe",
    "idea64.exe",
    "devenv.exe",
    "windowsterminal.exe",
    "wt.exe",
    "powershell.exe",
    "pwsh.exe",
    "cmd.exe",
    "winword.exe",
    "excel.exe",
    "powerpnt.exe",
    "pbidesktop.exe",
    "tableau.exe",
    "obsidian.exe",
    "notion.exe",
    "acrobat.exe",
    "acrord32.exe",
    "sumatrapdf.exe",
    "postman.exe",
    "docker desktop.exe",
    "mysqlworkbench.exe",
}

NEUTRAL_APPS = {
    "explorer.exe",
    "flow.launcher.exe",
    "searchhost.exe",
    "shellexperiencehost.exe",
    "taskmgr.exe",
    "applicationframehost.exe",
    "lockapp.exe",
    "lockscreen",
}

YOUTUBE_STUDY_PHRASES = (
    "python", "machine learning", "deep learning", "neural network", "artificial intelligence",
    "llm", "rag", "langchain", "pytorch", "tensorflow", "scikit", "pandas", "numpy",
    "sql", "postgres", "sqlite", "power bi", "tableau", "excel", "statistics",
    "probability", "linear algebra", "calculus", "dsa", "leetcode", "algorithm",
    "data structure", "system design", "backend", "fastapi", "django", "flask", "react",
    "typescript", "javascript", "docker", "kubernetes", "aws", "azure", "gcp", "linux",
    "git ", "github", "interview", "placement", "aptitude", "gate", "lecture", "tutorial",
    "course", "bootcamp", "stanford", "mit ", "cs50", "3blue1brown", "karpathy",
    "statquest", "freecodecamp", "explained", "documentation", "project", "coding",
    "programming", "data science", "data analyst", "analytics", "aiml", "roadmap",
    "dbms", "operating system", "computer network", "oop",
)

YOUTUBE_DISTRACTION_SIGNALS = (
    "#shorts",
    "youtube shorts",
    "official music video",
    "official video",
    "lyrical",
    "roast",
    "stand up",
    "stand-up",
    "comedy",
    "prank",
    "vlog",
    "reaction",
    "highlights",
    "gameplay",
    "walkthrough",
    "trailer",
    "teaser",
    "ipl",
    "cricket",
    "podcast clips",
    "meme",
)

BROWSER_FOCUS_HINTS = (
    "github",
    "colab",
    "stackoverflow",
    "leetcode",
    "hackerrank",
    "kaggle",
    "huggingface",
    "localhost",
    "127.0.0.1",
    "claude",
    "chatgpt",
    "gemini",
    "google docs",
    "google sheets",
    "notion",
    "coursera",
    "udemy",
    "nptel",
    "arxiv",
    "fastapi",
    "python",
    "chitran",
    "ditto",
    "sidekick",
    "jupyter",
    "wellfound",
    "cutshort",
    "linkedin",
    "internshala",
)

BROWSER_DISTRACTION_HINTS = (
    "instagram",
    "facebook",
    "reddit",
    "twitter",
    " / x",
    "x.com",
    "netflix",
    "prime video",
    "hotstar",
    "jiocinema",
    "twitch",
    "9gag",
    "reels",
    "cricbuzz",
)

GOAL_STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "into", "complete",
    "finish", "start", "today", "goal", "task", "work", "make", "build",
    "learn", "revise", "prepare", "check", "update", "write",
}


# ---------------------------------------------------------------------------
# Native Win32 Foreground Window + Idle Time Probes
# ---------------------------------------------------------------------------

class _LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("dwTime", wintypes.DWORD),
    ]


def get_idle_seconds() -> float:
    """Return seconds elapsed since the user's last keyboard or mouse input."""
    if sys.platform != "win32":
        return 0.0
    try:
        lii = _LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(_LASTINPUTINFO)
        if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            return 0.0
        now_tick = ctypes.windll.kernel32.GetTickCount64() & 0xFFFFFFFF
        idle_ms = (now_tick - lii.dwTime) & 0xFFFFFFFF
        return max(0.0, idle_ms / 1000.0)
    except Exception:
        return 0.0


def get_foreground_window_info() -> Tuple[str, str]:
    """Return `(process_basename, window_title)` for the active foreground window."""
    if sys.platform != "win32":
        return ("unknown", "")
    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return ("lockscreen", "Windows Lock Screen")

        title_len = user32.GetWindowTextLengthW(hwnd)
        title = ""
        if title_len > 0:
            buf = ctypes.create_unicode_buffer(title_len + 1)
            user32.GetWindowTextW(hwnd, buf, title_len + 1)
            title = buf.value.strip()

        pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

        app_name = "unknown"
        if pid.value:
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            h_proc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
            if h_proc:
                try:
                    exe_buf = ctypes.create_unicode_buffer(512)
                    size = wintypes.DWORD(512)
                    if kernel32.QueryFullProcessImageNameW(h_proc, 0, exe_buf, ctypes.byref(size)):
                        app_name = os.path.basename(exe_buf.value).lower()
                finally:
                    kernel32.CloseHandle(h_proc)

        return (app_name or "unknown", title[:240])
    except Exception:
        return ("unknown", "")


# ---------------------------------------------------------------------------
# Classifier Helpers
# ---------------------------------------------------------------------------

def extract_goal_keywords(goals: Iterable[Dict[str, Any]]) -> Set[str]:
    """Extract informative keywords from today's pending goal titles."""
    keywords: Set[str] = set()
    for g in goals:
        title = (g.get("title") or "").lower()
        for tok in re.findall(r"[a-z0-9]{3,}", title):
            if tok not in GOAL_STOPWORDS:
                keywords.add(tok)
    return keywords


def is_youtube_window(app_name: str, window_title: str) -> bool:
    app_l = (app_name or "").lower()
    title_l = (window_title or "").lower()
    return (app_l in BROWSER_APPS or app_l == "") and ("youtube" in title_l)


def classify_youtube_title(
    window_title: str,
    goal_keywords: Optional[Set[str]] = None,
) -> str:
    """Classify a YouTube window title locally as 'focus' (study) or 'distraction'."""
    clean = re.sub(r"^\(\d+\)\s*", "", (window_title or "").strip()).lower()
    clean = re.sub(r"\s*-\s*youtube.*$", "", clean).strip()

    for bad in YOUTUBE_DISTRACTION_SIGNALS:
        if bad in clean:
            return "distraction"

    for good in YOUTUBE_STUDY_PHRASES:
        if good in clean:
            return "focus"

    if goal_keywords:
        words = set(re.findall(r"[a-z0-9]{3,}", clean))
        if words & goal_keywords:
            return "focus"

    # Bare YouTube home feed / unknown clip defaults to distraction
    return "distraction"


def classify_activity(
    app_name: str,
    window_title: str,
    idle_sec: float = 0.0,
    goal_keywords: Optional[Set[str]] = None,
) -> str:
    """Categorize local activity into 'focus', 'distraction', 'neutral', or 'idle'."""
    app_l = (app_name or "unknown").lower()
    title_l = (window_title or "").lower()

    if idle_sec >= IDLE_PAUSE_SEC or app_l == "lockscreen":
        return "idle"

    if "chitran" in title_l or "sidekick os" in title_l or "ditto" == title_l or "byte" == title_l:
        return "neutral"

    if app_l in FOCUS_APPS:
        return "focus"

    if app_l in BROWSER_APPS:
        if is_youtube_window(app_l, title_l):
            return classify_youtube_title(title_l, goal_keywords=goal_keywords)
        for bad in BROWSER_DISTRACTION_HINTS:
            if bad in title_l:
                return "distraction"
        for good in BROWSER_FOCUS_HINTS:
            if good in title_l:
                return "focus"
        return "neutral"

    if app_l in NEUTRAL_APPS:
        return "neutral"

    if any(g in app_l for g in ("steam", "valorant", "riot", "epicgames")):
        return "distraction"

    return "neutral"


# ---------------------------------------------------------------------------
# Compact Non-Focus-Stealing Dark Cyber-HUD Nudge
# ---------------------------------------------------------------------------

def _record_habit_event(event_type: str, action_taken: str = "triggered", day: Optional[str] = None) -> None:
    init_db()
    target_day = day or today_iso()
    now_iso = datetime.now().isoformat(timespec="seconds")
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO habit_events (date, timestamp, event_type, action_taken) VALUES (?, ?, ?, ?)",
            (target_day, now_iso, event_type, action_taken),
        )


def spawn_nudge_hud(event_type: str, title: str, message: str) -> None:
    """Spawn a lightweight non-blocking Cyber-HUD nudge card above the taskbar corner."""
    try:
        monitor_py = Path(__file__).resolve()
        subprocess.Popen(
            [sys.executable, str(monitor_py), "--nudge", event_type, title, message],
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
    except Exception as e:
        print(f"[Monitor] Failed to spawn nudge HUD: {e}")


def show_nudge_card(event_type: str, title: str, message: str) -> None:
    """Render a sleek compact dark-mode corner nudge that dismisses cleanly."""
    import tkinter as tk
    from win_ui import apply_win11_dwm_style, enable_high_dpi

    enable_high_dpi()
    accent_map = {
        "youtube_nudge": "#F43F5E",
        "distraction_nudge": "#F43F5E",
        "water_reminder": "#38BDF8",
        "eye_break": "#10B981",
    }
    accent = accent_map.get(event_type, "#22D3EE")

    root = tk.Tk()
    root.title(f"Ditto Nudge — {title}")
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.configure(bg="#18181B")

    outer = tk.Frame(root, bg=accent, padx=1, pady=1)
    outer.pack(fill="both", expand=True)
    card = tk.Frame(outer, bg="#121216", padx=16, pady=12)
    card.pack(fill="both", expand=True)

    tk.Frame(card, bg=accent, height=2).pack(fill="x", pady=(0, 8))
    tk.Label(
        card,
        text=title,
        bg="#121216",
        fg=accent,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        card,
        text=message,
        bg="#121216",
        fg="#E4E4E7",
        font=("Segoe UI", 10),
        justify="left",
        wraplength=330,
        anchor="w",
    ).pack(fill="x", pady=(4, 10))

    btn_row = tk.Frame(card, bg="#121216")
    btn_row.pack(fill="x")

    def acknowledge():
        try:
            _record_habit_event(event_type, action_taken="acknowledged")
        except Exception:
            pass
        root.destroy()

    ack_label = "Drank Water ✓" if event_type == "water_reminder" else ("Back on Track ✓" if "nudge" in event_type else "Rested Eyes ✓")
    tk.Button(
        btn_row,
        text=ack_label,
        command=acknowledge,
        bg=accent,
        fg="#09090B",
        activebackground="#E4E4E7",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=4,
        cursor="hand2",
    ).pack(side="right")

    tk.Button(
        btn_row,
        text="Dismiss",
        command=root.destroy,
        bg="#27272A",
        fg="#A1A1AA",
        activebackground="#3F3F46",
        font=("Segoe UI", 9),
        relief="flat",
        padx=10,
        pady=4,
        cursor="hand2",
    ).pack(side="right", padx=(0, 8))

    root.update_idletasks()
    w = max(360, root.winfo_reqwidth())
    h = root.winfo_reqheight()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{sw - w - 20}+{max(30, sh - h - 72)}")
    apply_win11_dwm_style(root, dark_titlebar=False, rounded=True)
    root.lift()
    try:
        ctypes.windll.user32.MessageBeep(0x00000040)
    except Exception:
        pass

    root.bind("<Escape>", lambda _: root.destroy())
    # Stay visible long enough to drink water / rest eyes and click confirm on return
    auto_close_ms = 180000 if event_type in ("water_reminder", "eye_break") else 90000
    root.after(auto_close_ms, root.destroy)
    root.mainloop()


# ---------------------------------------------------------------------------
# Stateful Monitor Engine
# ---------------------------------------------------------------------------

class ActivityMonitor:
    """Tracks continuous active focus/distraction/hydration/eye timers every 15s."""

    def __init__(self, poll_sec: int = POLL_INTERVAL_SEC):
        self.poll_sec = poll_sec
        self.water_timer_sec: int = 0
        self.eye_timer_sec: int = 0
        self.distraction_streak_sec: int = 0
        self.last_distraction_nudge_ts: float = 0.0
        self.hydration_hold_until_ts: float = 0.0
        self.alert_hold_until_ts: float = 0.0
        self.current_app: str = ""
        self.current_title: str = ""
        self.current_category: str = "neutral"
        self.last_idle_sec: float = 0.0

    def tick(
        self,
        now_ts: Optional[float] = None,
        window_info: Optional[Tuple[str, str]] = None,
        idle_sec: Optional[float] = None,
        notify: bool = True,
        has_active_tasks: bool = False,
    ) -> Dict[str, Any]:
        """
        Execute one 15s monitoring step: inspect active window + idle state,
        log locally to SQLite, update break timers, and fire smart nudges.
        """
        init_db()
        now_dt = datetime.fromtimestamp(now_ts) if now_ts is not None else datetime.now()
        epoch_now = now_ts if now_ts is not None else now_dt.timestamp()
        day = now_dt.date().isoformat()
        now_iso = now_dt.isoformat(timespec="seconds")

        app_name, title = window_info if window_info is not None else get_foreground_window_info()
        idle = float(idle_sec if idle_sec is not None else get_idle_seconds())

        pending_goals = get_pending_goals(day)
        goal_keywords = extract_goal_keywords(pending_goals)
        category = classify_activity(app_name, title, idle_sec=idle, goal_keywords=goal_keywords)

        self.current_app = app_name
        self.current_title = title
        self.current_category = category
        self.last_idle_sec = idle

        # Always log local-only activity record when not paused
        paused = is_paused()
        if not paused:
            with get_connection() as conn:
                conn.execute(
                    """
                    INSERT INTO activity_logs (date, timestamp, app_name, window_title, category, duration_sec)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (day, now_iso, app_name, title[:240], category, self.poll_sec),
                )

        events_fired: List[str] = []

        if paused:
            return {
                "category": category,
                "paused": True,
                "events_fired": events_fired,
            }

        # Auto-reset continuous break timers when user has stepped away (>= 5 mins idle)
        if idle >= IDLE_RESET_SEC or category == "idle":
            if idle >= IDLE_RESET_SEC:
                self.water_timer_sec = 0
                self.eye_timer_sec = 0
            self.distraction_streak_sec = 0
            return {
                "category": category,
                "idle_sec": idle,
                "water_timer_sec": self.water_timer_sec,
                "eye_timer_sec": self.eye_timer_sec,
                "events_fired": events_fired,
            }

        # Active user: advance hydration & eye-break continuous clocks
        self.water_timer_sec += self.poll_sec
        self.eye_timer_sec += self.poll_sec

        # 1) Distraction / YouTube accountability tracking
        if category == "distraction":
            self.distraction_streak_sec += self.poll_sec
            on_youtube = is_youtube_window(app_name, title)
            threshold = YOUTUBE_NUDGE_STREAK_SEC if on_youtube else DISTRACTION_NUDGE_STREAK_SEC

            if (
                pending_goals
                and self.distraction_streak_sec >= threshold
                and (epoch_now - self.last_distraction_nudge_ts) >= DISTRACTION_NUDGE_COOLDOWN_SEC
            ):
                top_goal = pending_goals[0]
                dl_txt = f" (due {top_goal['deadline_time']})" if top_goal.get("deadline_time") else ""
                event_name = "youtube_nudge" if on_youtube else "distraction_nudge"
                nudge_title = "Ditto Focus Check!" if on_youtube else "Distraction Alert!"
                msg = f"You've drifted for {self.distraction_streak_sec // 60} mins while '{top_goal['title']}'{dl_txt} is still pending."
                _record_habit_event(event_name, "triggered", day=day)
                self.last_distraction_nudge_ts = epoch_now
                self.distraction_streak_sec = 0
                self.alert_hold_until_ts = epoch_now + 30.0
                if not has_active_tasks:
                    set_mascot_state("alert")
                if notify:
                    speak_background(f"Hey Rohit, {msg} Let's lock back in.")
                    spawn_nudge_hud(event_name, nudge_title, msg)
                events_fired.append(event_name)
        else:
            self.distraction_streak_sec = 0
            if category == "focus":
                self.alert_hold_until_ts = 0.0

        # 2) Smart 45-min Water Reminder
        if self.water_timer_sec >= WATER_REMINDER_SEC:
            self.water_timer_sec = 0
            msg = "45 minutes of continuous focus! Take a quick sip of water to stay sharp."
            _record_habit_event("water_reminder", "triggered", day=day)
            self.hydration_hold_until_ts = epoch_now + HYDRATION_MASCOT_HOLD_SEC
            if not has_active_tasks:
                set_mascot_state("hydration")
            if notify:
                speak_background("Hydration check, Rohit. Take a quick sip of water.")
                spawn_nudge_hud("water_reminder", "Hydration Check", msg)
            events_fired.append("water_reminder")

        # 3) Smart 60-min 20-20-20 Eye-Break Reminder (avoid exact tick collision if water just fired)
        elif self.eye_timer_sec >= EYE_BREAK_SEC:
            self.eye_timer_sec = 0
            msg = "60 minutes on screen! Look 20 feet away for 20 seconds to rest your eyes."
            _record_habit_event("eye_break", "triggered", day=day)
            self.hydration_hold_until_ts = epoch_now + HYDRATION_MASCOT_HOLD_SEC
            if not has_active_tasks:
                set_mascot_state("hydration")
            if notify:
                speak_background("Sixty minutes on screen, Rohit. Rest your eyes for twenty seconds.")
                spawn_nudge_hud("eye_break", "20-20-20 Eye Break", msg)
            events_fired.append("eye_break")

        # 4) Restore normal goal/focus mascot state once temporary holds expire
        if (
            not has_active_tasks
            and not events_fired
            and epoch_now >= self.hydration_hold_until_ts
            and epoch_now >= self.alert_hold_until_ts
        ):
            cur_mascot = load_system_state().get("mascot_state", "focused")
            if cur_mascot != "working":
                refresh_mascot_goal_state(day)

        return {
            "category": category,
            "idle_sec": idle,
            "water_timer_sec": self.water_timer_sec,
            "eye_timer_sec": self.eye_timer_sec,
            "distraction_streak_sec": self.distraction_streak_sec,
            "events_fired": events_fired,
        }


# Global daemon monitor singleton
_MONITOR = ActivityMonitor()


def get_monitor_snapshot() -> Dict[str, Any]:
    """Return live monitor state + today's aggregated focus vs distraction seconds."""
    init_db()
    day = today_iso()
    totals = {"focus": 0, "distraction": 0, "neutral": 0, "idle": 0}
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT category, SUM(duration_sec) AS secs FROM activity_logs WHERE date = ? GROUP BY category",
            (day,),
        ).fetchall()
        for r in rows:
            cat = r["category"]
            if cat in totals:
                totals[cat] = int(r["secs"] or 0)

    return {
        "date": day,
        "paused": is_paused(),
        "current_app": _MONITOR.current_app,
        "current_category": _MONITOR.current_category,
        "idle_sec": round(_MONITOR.last_idle_sec, 1),
        "water_timer_sec": _MONITOR.water_timer_sec,
        "water_next_in_sec": max(0, WATER_REMINDER_SEC - _MONITOR.water_timer_sec),
        "eye_timer_sec": _MONITOR.eye_timer_sec,
        "eye_next_in_sec": max(0, EYE_BREAK_SEC - _MONITOR.eye_timer_sec),
        "distraction_streak_sec": _MONITOR.distraction_streak_sec,
        "today_seconds": totals,
    }


def hold_mascot_state(target_state: str, duration_sec: float = 35.0) -> None:
    """Temporarily hold a mascot state (e.g. 'mail' or 'hydration') without monitor overwrite."""
    if not is_paused():
        _MONITOR.hydration_hold_until_ts = datetime.now().timestamp() + float(duration_sec)
        set_mascot_state(target_state)


def trigger_sample_nudge(kind: str = "water_reminder", notify: bool = True) -> Dict[str, Any]:
    """Trigger a live water/eye/YouTube accountability preview without waiting 45 mins."""
    now_ts = datetime.now().timestamp()
    norm = (kind or "water_reminder").strip().lower()
    if norm in ("water", "water_reminder", "hydration"):
        event_type = "water_reminder"
        title = "Hydration Check"
        msg = "45 minutes of continuous focus! Take a quick sip of water to stay sharp."
        speech = "Hydration check, Rohit. Take a quick sip of water."
        mascot_target = "hydration"
        _MONITOR.hydration_hold_until_ts = now_ts + 16.0
    elif norm in ("eye", "eye_break"):
        event_type = "eye_break"
        title = "20-20-20 Eye Break"
        msg = "60 minutes on screen! Look 20 feet away for 20 seconds to rest your eyes."
        speech = "Sixty minutes on screen, Rohit. Rest your eyes for twenty seconds."
        mascot_target = "hydration"
        _MONITOR.hydration_hold_until_ts = now_ts + 16.0
    else:
        event_type = "youtube_nudge"
        title = "Ditto Focus Check!"
        msg = "Distraction threshold test — time to lock back onto your pending goal!"
        speech = "Hey Rohit, focus check. Let's lock back onto today's goal."
        mascot_target = "alert"
        _MONITOR.alert_hold_until_ts = now_ts + 16.0

    if not is_paused():
        set_mascot_state(mascot_target)
    if notify:
        _record_habit_event(event_type, "test_triggered")
        speak_background(speech)
        spawn_nudge_hud(event_type, title, msg)
    return {"status": "triggered", "event_type": event_type, "mascot_state": mascot_target}


async def monitor_loop(has_active_tasks: Optional[Callable[[], bool]] = None) -> None:
    """Async 15-second background polling loop executed inside FastAPI daemon."""
    print("[Monitor] Active window & habit monitor started (15s cadence, 100% local).")
    while True:
        try:
            busy = bool(has_active_tasks()) if has_active_tasks else False
            _MONITOR.tick(notify=True, has_active_tasks=busy)
        except asyncio.CancelledError:
            print("[Monitor] Background monitor loop stopped.")
            raise
        except Exception as e:
            print(f"[Monitor] Poll tick error (recovered): {e}")
        await asyncio.sleep(POLL_INTERVAL_SEC)


if __name__ == "__main__":
    if len(sys.argv) >= 5 and sys.argv[1] == "--nudge":
        show_nudge_card(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        app, ttl = get_foreground_window_info()
        idle = get_idle_seconds()
        cat = classify_activity(app, ttl, idle_sec=idle)
        print(f"app={app!r} title={ttl!r} idle={idle:.1f}s category={cat}")
