import asyncio
import ctypes
import os
import sys
import uuid
from pathlib import Path
from typing import Dict, Optional

# Ensure pythonw.exe (windowless Python where stdout/stderr are None) never crashes uvicorn logging
if sys.stdout is None or sys.stderr is None:
    _log_dir = Path(__file__).resolve().parent.parent / "data"
    _log_dir.mkdir(parents=True, exist_ok=True)
    _log_file = open(_log_dir / "daemon.log", "a", encoding="utf-8", buffering=1)
    if sys.stdout is None:
        sys.stdout = _log_file
    if sys.stderr is None:
        sys.stderr = _log_file

from pydantic import BaseModel
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
import uvicorn

# Local imports
from db import init_db
from quota import get_daily_quota_summary
from notes import append_daily_note
from router import query_mentor
from hud import spawn_hud_process
from form_filler import auto_fill_active_form
from colab_runner import run_colab_notebook
from voice import is_voice_muted, set_voice_muted, toggle_mute, speak_background
from state import load_system_state, set_paused, toggle_paused, set_mascot_state, is_paused
from planner import get_today_goals, handle_goal_command, mark_goal_done, refresh_mascot_goal_state, today_iso
from monitor import monitor_loop, get_monitor_snapshot, trigger_sample_nudge
from gmail_briefing import run_gmail_briefing, get_mail_briefing_status, mail_scheduler_loop
from researcher import research_topic
from organizer import plan_downloads_organization, undo_last_organize
from office_gen import prepare_office_outline, compile_office_from_draft
from analytics import get_analytics_summary

DASHBOARD_HTML = Path(__file__).resolve().parent / "templates" / "dashboard.html"

# Win32 Execution State Flags (Survives Win + L screen lock)
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_AWAYMODE_REQUIRED = 0x00000040

app = FastAPI(title="Chitran Local Daemon", version="2.0")


@app.middleware("http")
async def _localhost_origin_guard(request: Request, call_next):
    origin = request.headers.get("origin")
    if origin and not (
        origin.startswith("http://127.0.0.1") or origin.startswith("http://localhost")
    ):
        return JSONResponse(status_code=403, content={"detail": "External origins are not allowed"})
    return await call_next(request)

# State tracking
active_tasks: Dict[str, asyncio.Task] = {}
is_lock_active = False
_monitor_task: Optional[asyncio.Task] = None
_mail_task: Optional[asyncio.Task] = None


def acquire_system_lock():
    global is_lock_active
    try:
        if sys.platform == "win32":
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
            )
            is_lock_active = True
    except Exception as e:
        print(f"[Daemon] Failed to acquire Win32 execution lock: {e}")


def release_system_lock():
    global is_lock_active
    try:
        if sys.platform == "win32":
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            is_lock_active = False
    except Exception as e:
        print(f"[Daemon] Failed to release Win32 execution lock: {e}")


async def _managed_colab_task(task_id: str, notebook_url: str):
    """Runs Colab notebook inside a tracked task and releases Win32 lock when done."""
    set_mascot_state("working")
    try:
        result = await run_colab_notebook(notebook_url=notebook_url)
        print(f"[Daemon] Task {task_id} finished: {result}")
    except asyncio.CancelledError:
        print(f"[Daemon] Task {task_id} cancelled by user.")
    except Exception as e:
        print(f"[Daemon] Task {task_id} failed: {e}")
    finally:
        active_tasks.pop(task_id, None)
        if not active_tasks:
            release_system_lock()
            set_mascot_state("paused" if is_paused() else "focused")


class CommandRequest(BaseModel):
    cmd: str
    arg: str = ""


@app.on_event("startup")
async def on_startup():
    global _monitor_task, _mail_task
    init_db()
    if _monitor_task is None or _monitor_task.done():
        _monitor_task = asyncio.create_task(
            monitor_loop(has_active_tasks=lambda: bool(active_tasks))
        )
    if _mail_task is None or _mail_task.done():
        _mail_task = asyncio.create_task(mail_scheduler_loop())


@app.on_event("shutdown")
async def on_shutdown():
    global _monitor_task, _mail_task
    for task in (_monitor_task, _mail_task):
        if task and not task.done():
            task.cancel()
            try:
                await task
            except asyncio.CancelledError:
                pass
    _monitor_task = None
    _mail_task = None


@app.get("/health")
async def health_check():
    st = load_system_state()
    return {
        "status": "ok",
        "service": "Chitran Daemon",
        "version": "2.0",
        "active_tasks": len(active_tasks),
        "keepalive_locked": is_lock_active,
        "voice_muted": is_voice_muted(),
        "paused": st.get("paused", False),
        "mascot_state": st.get("mascot_state", "focused")
    }


@app.get("/quota")
async def quota_status():
    return get_daily_quota_summary()


@app.get("/monitor/status")
async def monitor_status():
    return get_monitor_snapshot()


@app.get("/mail/status")
async def mail_status():
    return get_mail_briefing_status()


@app.post("/monitor/test-nudge")
async def monitor_test_nudge(req: CommandRequest):
    kind = (req.arg or req.cmd or "water_reminder").strip()
    return trigger_sample_nudge(kind=kind, notify=True)


@app.get("/goals/today")
async def goals_today():
    goals = get_today_goals()
    overdue = refresh_mascot_goal_state()
    return {
        "date": today_iso(),
        "count": len(goals),
        "overdue_count": len(overdue),
        "goals": goals,
    }


@app.get("/dashboard", response_class=HTMLResponse)
async def serve_dashboard():
    if not DASHBOARD_HTML.exists():
        raise HTTPException(status_code=404, detail="Dashboard template not found.")
    return HTMLResponse(content=DASHBOARD_HTML.read_text(encoding="utf-8"))


@app.get("/api/analytics")
async def api_analytics(range: str = "day"):
    return get_analytics_summary(range_key=range)


@app.post("/command")
async def execute_command(req: CommandRequest):
    cmd = req.cmd.strip().lower()
    arg = req.arg.strip()

    if cmd == "note":
        return append_daily_note(arg)

    elif cmd == "ask":
        res = await query_mentor(arg)
        return res

    elif cmd == "fill":
        res = await auto_fill_active_form(role_hint=arg)
        return res

    elif cmd == "colab":
        task_id = f"colab-{uuid.uuid4().hex[:8]}"
        acquire_system_lock()
        task = asyncio.create_task(_managed_colab_task(task_id, arg))
        active_tasks[task_id] = task
        return {
            "status": "started",
            "task_id": task_id,
            "keepalive_locked": is_lock_active,
            "message": "Colab execution started in background. Win + L safe."
        }

    elif cmd == "mute":
        if arg.lower() in ("on", "true", "1"):
            muted = set_voice_muted(True)
        elif arg.lower() in ("off", "false", "0"):
            muted = set_voice_muted(False)
        else:
            muted = toggle_mute()
        return {
            "status": "success",
            "muted": muted,
            "message": "Headphone voice muted." if muted else "Headphone voice unmuted."
        }

    elif cmd in ("speak", "say"):
        if is_voice_muted():
            return {"status": "muted", "message": "Voice is currently muted. Run 'cn mute' to unmute."}
        speak_background(arg)
        return {"status": "success", "message": "Speaking in background."}

    elif cmd == "pause":
        if arg.lower() in ("on", "true", "1"):
            st = set_paused(True)
        elif arg.lower() in ("off", "false", "0"):
            st = set_paused(False)
        else:
            st = toggle_paused()
        paused = st["paused"]
        return {
            "status": "success",
            "paused": paused,
            "mascot_state": st["mascot_state"],
            "message": "Chitran monitoring paused." if paused else "Chitran monitoring resumed."
        }

    elif cmd in ("goal", "goals"):
        return handle_goal_command(arg)

    elif cmd == "done":
        return mark_goal_done(arg)

    elif cmd in ("mail", "gmail", "email"):
        spawn_hud = arg.strip().lower() != "plugin"
        return await run_gmail_briefing(trigger="manual", arg=arg, spawn_hud=spawn_hud)

    elif cmd in ("research", "res"):
        set_mascot_state("working")
        try:
            return await research_topic(arg)
        finally:
            if not active_tasks:
                if is_paused():
                    set_mascot_state("paused")
                else:
                    refresh_mascot_goal_state()

    elif cmd in ("organize", "org"):
        return plan_downloads_organization()

    elif cmd == "undo":
        set_mascot_state("working")
        try:
            return undo_last_organize()
        finally:
            if not active_tasks:
                if is_paused():
                    set_mascot_state("paused")
                else:
                    refresh_mascot_goal_state()

    elif cmd in ("doc", "docx", "excel", "xlsx", "ppt", "pptx"):
        set_mascot_state("working")
        try:
            return await prepare_office_outline(cmd, arg)
        finally:
            if not active_tasks:
                if is_paused():
                    set_mascot_state("paused")
                else:
                    refresh_mascot_goal_state()

    elif cmd == "office_build":
        if not arg:
            return {"status": "error", "message": "Missing draft JSON path."}
        set_mascot_state("working")
        try:
            return await compile_office_from_draft(Path(arg))
        finally:
            if not active_tasks:
                if is_paused():
                    set_mascot_state("paused")
                else:
                    refresh_mascot_goal_state()

    elif cmd in ("stats", "dashboard"):
        return {
            "status": "success",
            "url": "http://127.0.0.1:7777/dashboard",
            "message": "Analytics dashboard ready.",
        }

    elif cmd == "kill":
        cancelled_count = 0
        for task_id, task in list(active_tasks.items()):
            if not task.done():
                task.cancel()
                cancelled_count += 1
            active_tasks.pop(task_id, None)

        release_system_lock()
        return {
            "status": "success",
            "message": f"Killed {cancelled_count} active task(s). Keepalive lock released."
        }

    else:
        raise HTTPException(
            status_code=400,
            detail=f"Unknown command: '{cmd}'. Supported commands: note, ask, research, organize, undo, doc, excel, ppt, stats, fill, colab, mute, speak, pause, goal, done, mail, kill"
        )


if __name__ == "__main__":
    port = int(os.environ.get("SIDEKICK_PORT", 7777))
    print(f"[Daemon] Starting Chitran Daemon on 127.0.0.1:{port}...")
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info", use_colors=False)
