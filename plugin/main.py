import sys
import json
import os
import subprocess
import time
import urllib.request
import urllib.error
import webbrowser
from pathlib import Path

DAEMON_URL = "http://127.0.0.1:7777/command"
ICON_PATH = "assets\\icon.png"

# Path references
PLUGIN_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PLUGIN_DIR.parent
DAEMON_SERVER_PY = PROJECT_ROOT / "daemon" / "server.py"
HUD_PY = PROJECT_ROOT / "daemon" / "hud.py"
ORGANIZER_PY = PROJECT_ROOT / "daemon" / "organizer.py"
OFFICE_GEN_PY = PROJECT_ROOT / "daemon" / "office_gen.py"
PLANNER_UI_PY = PROJECT_ROOT / "daemon" / "planner_ui.py"
MASCOT_PY = PROJECT_ROOT / "daemon" / "mascot.py"
MASCOT_PID_FILE = PROJECT_ROOT / "data" / "mascot.pid"
STATE_FILE = PROJECT_ROOT / "data" / "system_state.json"


def ensure_mascot_running(force_show_badge: bool = False):
    """Ensures Byte the Cyber-Slime is running in the user's interactive desktop session."""
    if force_show_badge:
        try:
            st = {}
            if STATE_FILE.exists():
                st = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            st["badge_visible"] = True
            st["respawn_seq"] = int(st.get("respawn_seq", 0)) + 1
            STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
            STATE_FILE.write_text(json.dumps(st, indent=2), encoding="utf-8")
        except Exception:
            pass

    if MASCOT_PID_FILE.exists():
        try:
            import ctypes
            pid = int(MASCOT_PID_FILE.read_text().strip())
            if pid > 0:
                h = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
                if h:
                    ctypes.windll.kernel32.CloseHandle(h)
                    return True
        except Exception:
            pass

    try:
        subprocess.Popen(
            [sys.executable, str(MASCOT_PY)],
            cwd=str(PROJECT_ROOT / "daemon")
        )
        return True
    except Exception:
        return False


def show_toast(title: str, message: str):
    """Shows a quick native Windows notification without external libraries."""
    try:
        import ctypes
        # Play subtle Windows Asterisk chime
        ctypes.windll.user32.MessageBeep(0x00000040)
    except Exception:
        pass

    # Clean strings for powershell
    safe_title = title.replace('"', '`"').replace("'", "''")
    safe_msg = message.replace('"', '`"').replace("'", "''")

    ps_cmd = (
        f"[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); "
        f"$notify = New-Object System.Windows.Forms.NotifyIcon; "
        f"$notify.Icon = [System.Drawing.SystemIcons]::Information; "
        f"$notify.Visible = $True; "
        f"$notify.ShowBalloonTip(3000, '{safe_title}', '{safe_msg}', [System.Windows.Forms.ToolTipIcon]::Info); "
        f"Start-Sleep -Seconds 3; "
        f"$notify.Dispose()"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd],
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
    except Exception:
        pass


def ensure_daemon_running():
    """Checks if the local daemon is reachable; if not, spawns it detached in background."""
    health_url = "http://127.0.0.1:7777/health"
    try:
        req = urllib.request.Request(health_url, method="GET")
        with urllib.request.urlopen(req, timeout=1.0) as resp:
            if resp.status == 200:
                return True
    except Exception:
        pass

    try:
        flags = 0
        if sys.platform == "win32":
            flags = subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS

        subprocess.Popen(
            [sys.executable, str(DAEMON_SERVER_PY)],
            cwd=str(PROJECT_ROOT / "daemon"),
            creationflags=flags,
            close_fds=True
        )
        # Poll /health for up to 8 seconds so cold-start imports finish cleanly
        for _ in range(32):
            time.sleep(0.25)
            try:
                req = urllib.request.Request(health_url, method="GET")
                with urllib.request.urlopen(req, timeout=0.8) as resp:
                    if resp.status == 200:
                        return True
            except Exception:
                pass
        return False
    except Exception:
        return False


def post_command_to_daemon(cmd: str, arg: str = "") -> dict:
    """Posts command to the local Sidekick OS daemon."""
    payload = json.dumps({"cmd": cmd, "arg": arg}).encode("utf-8")
    req = urllib.request.Request(
        DAEMON_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=30.0) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError:
        if ensure_daemon_running():
            try:
                with urllib.request.urlopen(req, timeout=30.0) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except Exception as e:
                return {"status": "error", "message": f"Daemon retry failed: {e}"}
        return {"status": "error", "message": "Failed to connect to Chitran Daemon on 127.0.0.1:7777."}
    except Exception as e:
        return {"status": "error", "message": str(e)}


def handle_query(query_str: str) -> list:
    """Parses keystrokes from Flow Launcher and returns 0ms instant preview rows."""
    q = query_str.strip()
    results = []

    if not q:
        results.append({
            "Title": "cn goal <task> @ <HH:MM>",
            "SubTitle": "Lock a daily goal & deadline (or type 'cn goal' to open Planner HUD)",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["goal", ""], "dontHideAfterAction": False}
        })
        results.append({
            "Title": "cn note <text>",
            "SubTitle": "Append timestamped note to ~/Daily_Notes/YYYY-MM-DD.md",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "show_hint", "parameters": ["Type: cn note <text>"], "dontHideAfterAction": True}
        })
        results.append({
            "Title": "cn ask <query>",
            "SubTitle": "Query candid senior mentor (Groq/Gemini) & open floating HUD",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "show_hint", "parameters": ["Type: cn ask <query>"], "dontHideAfterAction": True}
        })
        results.append({
            "Title": "cn fill [role]",
            "SubTitle": "Auto-fill active Chrome web form using deterministic profile vault",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["fill", ""], "dontHideAfterAction": False}
        })
        results.append({
            "Title": "cn organize",
            "SubTitle": "Preview & safely sort Downloads into Documents (zero deletion, 'cn undo' supported)",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["organize", ""], "dontHideAfterAction": False}
        })
        results.append({
            "Title": "cn doc / excel / ppt <topic>",
            "SubTitle": "Draft Word, Excel, or PowerPoint files with Stage 1 outline approval HUD",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["office_prompt", "doc"], "dontHideAfterAction": False}
        })
        results.append({
            "Title": "cn stats",
            "SubTitle": "Open Dark Analytics Dashboard (Day/Week/Month focus, goals, hydration & API quota)",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["stats", ""], "dontHideAfterAction": False}
        })
        results.append({
            "Title": "cn kill",
            "SubTitle": "Emergency stop: cancel all background tasks & release keepalive lock",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {"method": "dispatch_to_daemon", "parameters": ["kill", ""], "dontHideAfterAction": False}
        })
        return results

    parts = q.split(maxsplit=1)
    subcmd = parts[0].lower()
    arg = parts[1].strip() if len(parts) > 1 else ""

    if subcmd in ("goal", "goals"):
        if arg:
            title_text = f"Lock Daily Goal: {arg}"
            sub_text = "Press Enter to save goal (or open inline deadline prompt if @ HH:MM missing)"
        else:
            title_text = "Open Daily Goal & Deadline Planner"
            sub_text = "View today's goals, mark items done, or add a new target"
        results.append({
            "Title": title_text,
            "SubTitle": sub_text,
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["goal", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "done":
        title_text = f"Complete Goal: {arg}" if arg else "Complete Today's Pending Goal"
        results.append({
            "Title": title_text,
            "SubTitle": "Press Enter to mark goal completed in SQLite (accepts ID or keyword)",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["done", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "note":
        title_text = f"Save Note: {arg}" if arg else "cn note <text to append>"
        results.append({
            "Title": title_text,
            "SubTitle": "Press Enter to append to today's Markdown log file",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["note", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "ask":
        title_text = f"Ask Mentor: {arg}" if arg else "cn ask <architecture or code question>"
        results.append({
            "Title": title_text,
            "SubTitle": "Press Enter to query Mentor & open Floating HUD [Esc to dismiss]",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["ask", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "fill":
        role_label = f" for '{arg}'" if arg else ""
        results.append({
            "Title": f"Auto-Fill Active Form{role_label}",
            "SubTitle": "Press Enter to inspect Chrome (port 9222) and fill known fields",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["fill", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "colab":
        label = arg if arg else "Active Colab Tab"
        results.append({
            "Title": f"Run Background Colab: {label}",
            "SubTitle": "Press Enter to execute notebook with lock-screen keepalive",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["colab", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "mute":
        mode = f" ({arg})" if arg else " (Toggle)"
        results.append({
            "Title": f"Toggle Headphone Voice{mode}",
            "SubTitle": "Press Enter to mute/unmute Chitran spoken alerts",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["mute", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("say", "speak"):
        results.append({
            "Title": f"Speak in Headphones: {arg}" if arg else "cn say <text to speak>",
            "SubTitle": "Press Enter to test Edge Neural TTS voice output",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["speak", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "pause":
        mode = f" ({arg})" if arg else " (Toggle)"
        results.append({
            "Title": f"Pause / Resume Chitran{mode}",
            "SubTitle": "Press Enter to pause or resume Ditto mascot & background reminders",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["pause", arg],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("ditto", "byte", "slime", "cat", "mascot"):
        results.append({
            "Title": "Summon Ditto the Cyber-Slime",
            "SubTitle": "Press Enter to launch or restore Ditto on your Windows taskbar",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["ditto", ""],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("mail", "gmail", "email", "inbox"):
        results.append({
            "Title": "Gmail Briefing (rohit.newchallenger@gmail.com)",
            "SubTitle": "Reads today's emails via silent Chrome CDP, speaks summary & opens dark HUD",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["mail", "plugin"],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("research", "res", "find"):
        if arg:
            results.append({
                "Title": f"Research & Context Pack: {arg}",
                "SubTitle": "Scrapes technical sources in parallel -> saves ~/Daily_Notes/context_packs/ & opens Copy HUD",
                "IcoPath": ICON_PATH,
                "JsonRPCAction": {
                    "method": "dispatch_to_daemon",
                    "parameters": ["research", arg],
                    "dontHideAfterAction": False
                }
            })
        else:
            results.append({
                "Title": "cn research <topic or URL>",
                "SubTitle": "Press Enter to open the Research prompt, or keep typing a topic",
                "IcoPath": ICON_PATH,
                "JsonRPCAction": {
                    "method": "dispatch_to_daemon",
                    "parameters": ["research_prompt", ""],
                    "dontHideAfterAction": False
                }
            })

    elif subcmd in ("organize", "org", "clean", "sort"):
        results.append({
            "Title": "Safe Downloads Auto-Organizer (Preview HUD)",
            "SubTitle": "Scans Downloads -> shows confirmation HUD before moving any file (zero deletion)",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["organize", ""],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "undo":
        results.append({
            "Title": "Undo Last Downloads Organize Batch",
            "SubTitle": "Safely moves files from the last organize batch back to Downloads",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["undo", ""],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("doc", "docx", "excel", "xlsx", "ppt", "pptx"):
        ext_hint = {
            "doc": "Word (.docx)",
            "docx": "Word (.docx)",
            "excel": "Excel (.xlsx)",
            "xlsx": "Excel (.xlsx)",
            "ppt": "PowerPoint (.pptx)",
            "pptx": "PowerPoint (.pptx)",
        }[subcmd]
        if arg:
            results.append({
                "Title": f"Generate {ext_hint}: {arg}",
                "SubTitle": "Stage 1: Research & open editable Outline Approval HUD -> ~/Documents/Chitran_Outputs/",
                "IcoPath": ICON_PATH,
                "JsonRPCAction": {
                    "method": "dispatch_to_daemon",
                    "parameters": [subcmd, arg],
                    "dontHideAfterAction": False
                }
            })
        else:
            results.append({
                "Title": f"cn {subcmd} <topic>",
                "SubTitle": f"Press Enter to open {ext_hint} topic prompt, or keep typing a topic",
                "IcoPath": ICON_PATH,
                "JsonRPCAction": {
                    "method": "dispatch_to_daemon",
                    "parameters": ["office_prompt", subcmd],
                    "dontHideAfterAction": False
                }
            })

    elif subcmd in ("stats", "dashboard", "analytics", "metrics"):
        results.append({
            "Title": "Open Chitran Analytics Dashboard",
            "SubTitle": "Day / Week / Month focus vs distraction, completed goals, hydration streak & ₹0 quota gauges",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["stats", ""],
                "dontHideAfterAction": False
            }
        })

    elif subcmd in ("help", "guide", "docs"):
        results.append({
            "Title": "Open Chitran User Guide & Commands",
            "SubTitle": "Opens USER_GUIDE.md in floating dark HUD [Esc to dismiss]",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["help", ""],
                "dontHideAfterAction": False
            }
        })

    elif subcmd == "kill":
        results.append({
            "Title": "Emergency Stop: Abort Background Tasks",
            "SubTitle": "Press Enter to terminate background jobs and clear Win32 lock",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["kill", ""],
                "dontHideAfterAction": False
            }
        })

    else:
        results.append({
            "Title": f"Quick Note: {q}",
            "SubTitle": "Press Enter to save this as a daily note",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["note", q],
                "dontHideAfterAction": False
            }
        })
        results.append({
            "Title": f"Ask Mentor: {q}",
            "SubTitle": "Press Enter to query senior mentor",
            "IcoPath": ICON_PATH,
            "JsonRPCAction": {
                "method": "dispatch_to_daemon",
                "parameters": ["ask", q],
                "dontHideAfterAction": False
            }
        })

    return results


def main():
    if len(sys.argv) < 2:
        return

    try:
        request = json.loads(sys.argv[1])
    except Exception:
        return

    method = request.get("method")
    parameters = request.get("parameters", [])

    if method == "query":
        query_text = parameters[0] if parameters else ""
        results = handle_query(query_text)
        print(json.dumps({"result": results}))

    elif method == "dispatch_to_daemon":
        cmd = parameters[0] if len(parameters) > 0 else ""
        arg = parameters[1] if len(parameters) > 1 else ""

        if cmd in ("ditto", "byte"):
            ensure_daemon_running()
            ensure_mascot_running(force_show_badge=True)
            show_toast("Chitran — Ditto Online", "Ditto is now active on your taskbar!")
            print(json.dumps({"result": []}))
            return

        if cmd == "help":
            ensure_mascot_running(force_show_badge=False)
            try:
                subprocess.Popen([sys.executable, str(HUD_PY), "--guide"])
            except Exception as e:
                show_toast("Chitran Error", f"Guide HUD error: {e}")
            print(json.dumps({"result": []}))
            return

        if cmd == "research_prompt":
            ensure_daemon_running()
            ensure_mascot_running(force_show_badge=False)
            try:
                subprocess.Popen([sys.executable, str(HUD_PY), "--quick-research"])
            except Exception as e:
                show_toast("Chitran Error", f"Research prompt error: {e}")
            print(json.dumps({"result": []}))
            return

        if cmd == "office_prompt":
            ensure_daemon_running()
            ensure_mascot_running(force_show_badge=False)
            try:
                subprocess.Popen([sys.executable, str(OFFICE_GEN_PY), "--prompt", arg or "doc"])
            except Exception as e:
                show_toast("Chitran Error", f"Office prompt error: {e}")
            print(json.dumps({"result": []}))
            return

        if cmd in ("stats", "dashboard"):
            ensure_daemon_running()
            ensure_mascot_running(force_show_badge=False)
            webbrowser.open("http://127.0.0.1:7777/dashboard")
            show_toast("Chitran — Analytics Dashboard", "Opening http://127.0.0.1:7777/dashboard")
            print(json.dumps({"result": []}))
            return

        ensure_mascot_running(force_show_badge=False)
        res = post_command_to_daemon(cmd, arg)

        if cmd == "note" and res.get("status") == "success":
            entry = res.get("entry", arg)
            show_toast("Chitran — Note Saved", entry)

        elif cmd == "ask":
            # Flow Launcher's process is directly on the user's desktop!
            # Spawn the HUD from here so it ALWAYS appears on the visible screen.
            answer = res.get("answer", "No response returned.")
            try:
                subprocess.Popen(
                    [sys.executable, str(HUD_PY), "Mentor Review", arg, answer]
                )
            except Exception as e:
                show_toast("Chitran Error", f"HUD spawn error: {e}")

        elif cmd == "mail":
            today_cnt = res.get("today_count", 0)
            md = res.get("markdown") or res.get("message") or "No email briefing available."
            title_lbl = f"Gmail Briefing ({today_cnt} today)"
            try:
                subprocess.Popen(
                    [sys.executable, str(HUD_PY), title_lbl, "rohit.newchallenger@gmail.com", md]
                )
            except Exception as e:
                show_toast("Chitran Error", f"Mail HUD spawn error: {e}")

        elif cmd in ("research", "res"):
            if res.get("status") == "success" and res.get("file"):
                pack_file = res["file"]
                cnt = res.get("source_count", 0)
                show_toast("Chitran — Context Pack Ready", f"Saved {cnt} sources -> {Path(pack_file).name}")
                try:
                    subprocess.Popen(
                        [sys.executable, str(HUD_PY), "--context-pack", pack_file]
                    )
                except Exception as e:
                    show_toast("Chitran Error", f"Context HUD spawn error: {e}")
            else:
                show_toast("Chitran Research Error", res.get("message", "Could not build context pack."))

        elif cmd in ("organize", "org"):
            if res.get("status") == "error":
                show_toast("Chitran Organizer Error", res.get("message", "Could not scan Downloads."))
            elif res.get("move_count", 0) == 0:
                skipped = res.get("skipped_count", 0)
                show_toast(
                    "Chitran — Downloads Clean",
                    f"0 matching files ready to move ({skipped} unmatched/protected untouched).",
                )
            else:
                try:
                    subprocess.Popen(
                        [sys.executable, str(ORGANIZER_PY), "--preview"]
                    )
                except Exception as e:
                    show_toast("Chitran Error", f"Organizer HUD spawn error: {e}")

        elif cmd == "undo":
            show_toast("Chitran — Undo Organize", res.get("message", "Undo finished."))

        elif cmd in ("doc", "docx", "excel", "xlsx", "ppt", "pptx"):
            if res.get("status") == "outline_ready" and res.get("draft_path"):
                show_toast(
                    "Chitran — Outline Ready",
                    f"Review '{res.get('title', arg)}' in HUD before building.",
                )
                try:
                    subprocess.Popen(
                        [sys.executable, str(OFFICE_GEN_PY), "--review", res["draft_path"]]
                    )
                except Exception as e:
                    show_toast("Chitran Error", f"Office Review HUD spawn error: {e}")
            else:
                show_toast("Chitran Office Error", res.get("message", "Could not draft outline."))

        elif cmd == "kill":
            msg = res.get("message", "Tasks killed.")
            show_toast("Chitran — Emergency Brake", msg)

        elif cmd == "fill":
            if res.get("status") == "success":
                count = res.get("filled_count", 0)
                show_toast("Chitran — Form Auto-Fill", f"Filled {count} field(s). Review inputs & submit.")
            else:
                msg = res.get("message", "Could not fill active tab.")
                show_toast("Chitran Notice", msg)

        elif cmd == "colab":
            msg = res.get("message", "Colab task dispatched.")
            show_toast("Chitran — Colab Runner", msg)

        elif cmd == "mute":
            msg = res.get("message", "Voice state toggled.")
            show_toast("Chitran — Voice Control", msg)

        elif cmd == "pause":
            msg = res.get("message", "Monitoring state toggled.")
            show_toast("Chitran — Ditto Mascot", msg)

        elif cmd in ("goal", "goals"):
            status = res.get("status")
            if status == "needs_deadline":
                task_title = res.get("title", arg)
                try:
                    subprocess.Popen(
                        [sys.executable, str(PLANNER_UI_PY), "--deadline-for", task_title]
                    )
                except Exception as e:
                    show_toast("Chitran Error", f"Planner HUD error: {e}")
            elif status == "open_planner":
                try:
                    subprocess.Popen([sys.executable, str(PLANNER_UI_PY)])
                except Exception as e:
                    show_toast("Chitran Error", f"Planner HUD error: {e}")
            else:
                show_toast("Chitran — Daily Goal", res.get("message", "Goal updated."))

        elif cmd == "done":
            status = res.get("status")
            msg = res.get("message", "Goal status updated.")
            if status == "ambiguous":
                try:
                    subprocess.Popen([sys.executable, str(PLANNER_UI_PY)])
                except Exception:
                    pass
            show_toast("Chitran — Goal Tracker", msg)

        print(json.dumps({"result": []}))

    elif method == "show_hint":
        print(json.dumps({"result": []}))


if __name__ == "__main__":
    main()
