import ctypes
from ctypes import wintypes
import os
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path
from typing import Optional, List

# Ensure pythonw.exe (stdout/stderr None) never crashes silently
if sys.stdout is None or sys.stderr is None:
    _m_log_dir = Path(__file__).resolve().parent.parent / "data"
    _m_log_dir.mkdir(parents=True, exist_ok=True)
    _m_log_file = open(_m_log_dir / "mascot.log", "a", encoding="utf-8", buffering=1)
    if sys.stdout is None:
        sys.stdout = _m_log_file
    if sys.stderr is None:
        sys.stderr = _m_log_file

import httpx
import pystray
from PIL import Image, ImageTk
import tkinter as tk

from state import load_system_state, save_system_state, toggle_paused, is_paused
from voice import is_voice_muted, toggle_mute, speak_background
from planner import has_goals_today, today_iso
from notify import notify_windows
from win_ui import apply_win11_dwm_style, enable_high_dpi
from byte_art import get_target_hoodie_colors, lerp_rgb, render_byte_frame

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PID_FILE = PROJECT_ROOT / "data" / "mascot.pid"
DAEMON_BASE = "http://127.0.0.1:7777"


def render_byte_image(
    state: str = "focused",
    size: int = 64,
    frame_idx: int = 0,
    transparent_bg: tuple = (0, 0, 0, 0),
    hoodie_rgb: Optional[tuple] = None,
    facing_right: bool = False,
) -> Image.Image:
    """Renders a smooth vector animation frame of 'Byte' the Cyber-Slime."""
    return render_byte_frame(
        state=state,
        size=size,
        frame_idx=frame_idx,
        transparent_bg=transparent_bg,
        hoodie_rgb=hoodie_rgb,
        facing_right=facing_right,
    )


def get_taskbar_inside_coords(screen_width: int, screen_height: int, widget_size: int = 38) -> tuple:
    """
    Computes coordinates to place Byte INSIDE the Windows 11 taskbar strip itself
    (between the centered app icons and the right-side system tray clock).
    """
    if sys.platform == "win32":
        try:
            rect = wintypes.RECT()
            SPI_GETWORKAREA = 0x0030
            ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
            taskbar_top = rect.bottom
            taskbar_height = max(40, screen_height - taskbar_top)
            x = rect.right - widget_size - 210
            y = taskbar_top + max(2, (taskbar_height - widget_size) // 2)
            return x, y, taskbar_top
        except Exception:
            pass
    return screen_width - 250, screen_height - 42, screen_height - 48


def attach_to_taskbar_zorder(hwnd: int) -> None:
    """
    Binds the widget window as an owned tool-window of Windows 11's Shell_TrayWnd taskbar
    so clicking empty space on the taskbar NEVER pushes Byte behind the taskbar.
    """
    if sys.platform != "win32" or not hwnd:
        return
    try:
        user32 = ctypes.windll.user32
        GWL_EXSTYLE = -20
        GWLP_HWNDPARENT = -8
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_TOPMOST = 0x00000008

        ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style | WS_EX_TOOLWINDOW | WS_EX_TOPMOST)

        tray_hwnd = user32.FindWindowW("Shell_TrayWnd", None)
        if tray_hwnd:
            user32.SetWindowLongPtrW(hwnd, GWLP_HWNDPARENT, tray_hwnd)
    except Exception:
        pass


def pin_window_above_taskbar(hwnd: int) -> None:
    """Re-asserts top-of-Z-order above Shell_TrayWnd without stealing focus."""
    if sys.platform == "win32" and hwnd:
        try:
            user32 = ctypes.windll.user32
            HWND_TOPMOST = -1
            HWND_NOTOPMOST = -2
            # NOSIZE(0x1) | NOMOVE(0x2) | NOACTIVATE(0x10) | SHOWWINDOW(0x40) | NOSENDCHANGING(0x400)
            SWP_FLAGS = 0x0001 | 0x0002 | 0x0010 | 0x0040 | 0x0400
            user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
            user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        except Exception:
            pass


def is_process_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    return False


class ByteTrayController:
    """Manages the System Tray icon, animated state polling, and the inside-taskbar Cyber-Slime."""

    def __init__(self):
        st = load_system_state()
        self._last_state = st.get("mascot_state", "focused")
        self._last_respawn_seq = st.get("respawn_seq", 0)
        self._frame_idx = 0
        self._running = True
        self._badge_proc: Optional[subprocess.Popen] = None
        self.icon: Optional[pystray.Icon] = None

    def _post_daemon_cmd(self, cmd: str, arg: str = "") -> None:
        try:
            httpx.post(f"{DAEMON_BASE}/command", json={"cmd": cmd, "arg": arg}, timeout=2.0)
        except Exception:
            pass

    def _start_badge(self) -> None:
        if self._badge_proc is None or self._badge_proc.poll() is not None:
            creation = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            self._badge_proc = subprocess.Popen(
                [sys.executable, str(Path(__file__).resolve()), "--badge"],
                creationflags=creation
            )

    def _stop_badge(self) -> None:
        if self._badge_proc and self._badge_proc.poll() is None:
            try:
                self._badge_proc.terminate()
            except Exception:
                pass
            self._badge_proc = None

    def on_toggle_pause(self, icon, item):
        new_state = toggle_paused()
        self._last_state = new_state["mascot_state"]
        icon.icon = render_byte_image(self._last_state, size=64, frame_idx=0)
        icon.update_menu()

    def on_toggle_mute(self, icon, item):
        toggle_mute()
        icon.update_menu()

    def on_toggle_badge(self, icon, item):
        state = load_system_state()
        visible = not bool(state.get("badge_visible", True))
        state["badge_visible"] = visible
        save_system_state(state)
        if visible:
            self._start_badge()
        else:
            self._stop_badge()
        icon.update_menu()

    def on_open_planner(self, icon=None, item=None):
        creation = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        subprocess.Popen(
            [sys.executable, str(Path(__file__).resolve().parent / "planner_ui.py")],
            creationflags=creation,
        )

    def on_open_dashboard(self, icon, item):
        webbrowser.open(f"{DAEMON_BASE}/dashboard")

    def on_kill_tasks(self, icon, item):
        self._post_daemon_cmd("kill", "")

    def on_exit(self, icon, item):
        self._running = False
        self._stop_badge()
        try:
            if PID_FILE.exists():
                PID_FILE.unlink()
        except Exception:
            pass
        icon.stop()

    def _build_menu(self) -> pystray.Menu:
        return pystray.Menu(
            pystray.MenuItem(
                lambda item: f"Ditto Status: {load_system_state().get('mascot_state', 'focused').upper()}",
                None,
                enabled=False
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                lambda item: "Resume Chitran Monitoring" if is_paused() else "Pause Chitran Monitoring",
                self.on_toggle_pause
            ),
            pystray.MenuItem(
                lambda item: "Unmute Headphone Voice" if is_voice_muted() else "Mute Headphone Voice",
                self.on_toggle_mute
            ),
            pystray.MenuItem(
                lambda item: "Hide Taskbar Ditto" if load_system_state().get("badge_visible", True) else "Show Taskbar Ditto",
                self.on_toggle_badge
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Daily Goals & Planner", self.on_open_planner),
            pystray.MenuItem("Open Analytics Dashboard", self.on_open_dashboard),
            pystray.MenuItem("Emergency Kill Tasks", self.on_kill_tasks),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit Ditto Mascot", self.on_exit),
        )

    def _state_watcher(self):
        while self._running:
            time.sleep(0.8)
            try:
                cur_state = load_system_state()
                mascot_state = cur_state.get("mascot_state", "focused")
                cur_seq = cur_state.get("respawn_seq", 0)
                if cur_seq != self._last_respawn_seq:
                    self._last_respawn_seq = cur_seq
                    self._stop_badge()
                    if cur_state.get("badge_visible", True):
                        self._start_badge()
                elif cur_state.get("badge_visible", True) and (self._badge_proc is None or self._badge_proc.poll() is not None):
                    self._start_badge()
                self._frame_idx = (self._frame_idx + 1) % 12
                if self.icon:
                    self._last_state = mascot_state
                    self.icon.icon = render_byte_image(mascot_state, size=64, frame_idx=self._frame_idx)
                    self.icon.title = f"Chitran — Ditto ({mascot_state})"
            except Exception:
                pass

    def _ensure_daemon_alive(self) -> None:
        """Ensures local FastAPI daemon (127.0.0.1:7777) is running so Byte works standalone without Flow Launcher."""
        try:
            r = httpx.get(f"{DAEMON_BASE}/health", timeout=1.0)
            if r.status_code == 200:
                return
        except Exception:
            pass
        try:
            flags = (subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS) if sys.platform == "win32" else 0
            server_py = Path(__file__).resolve().parent / "server.py"
            subprocess.Popen(
                [sys.executable, str(server_py)],
                cwd=str(server_py.parent),
                creationflags=flags,
                close_fds=True,
            )
        except Exception:
            pass

    def _daily_goal_startup_check(self):
        """Ensures daemon is up and prompts Rohit once per day if no goal has been logged yet today."""
        self._ensure_daemon_alive()
        time.sleep(3.5)
        if not self._running or is_paused():
            return
        try:
            today = today_iso()
            st = load_system_state()
            if st.get("last_goal_prompt_date") == today:
                return
            if not has_goals_today(today):
                st["last_goal_prompt_date"] = today
                save_system_state(st)
                if not is_voice_muted():
                    speak_background("Hey Rohit, you haven't set a goal for today yet. What is our main target?")
                creation = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                subprocess.Popen(
                    [sys.executable, str(Path(__file__).resolve().parent / "planner_ui.py"), "--startup"],
                    creationflags=creation,
                )
        except Exception:
            pass

    def run(self):
        if PID_FILE.exists():
            try:
                old_pid = int(PID_FILE.read_text().strip())
                if old_pid != os.getpid() and is_process_alive(old_pid):
                    return
            except Exception:
                pass

        PID_FILE.parent.mkdir(parents=True, exist_ok=True)
        PID_FILE.write_text(str(os.getpid()))

        state = load_system_state()
        self._last_state = state.get("mascot_state", "focused")
        if state.get("badge_visible", True):
            self._start_badge()

        self.icon = pystray.Icon(
            "Chitran_Ditto",
            icon=render_byte_image(self._last_state, size=64, frame_idx=0),
            title=f"Chitran — Ditto ({self._last_state})",
            menu=self._build_menu()
        )
        threading.Thread(target=self._state_watcher, daemon=True).start()
        threading.Thread(target=self._daily_goal_startup_check, daemon=True).start()
        self.icon.run()


def run_mini_badge():
    """
    Renders 'Byte' the Cyber-Slime INSIDE the Windows 11 taskbar strip with smooth
    vector animations and a custom borderless Dark Cyber-HUD Popover Menu on Left/Right-Click.
    """
    enable_high_dpi()
    TRANSPARENT_KEY = "#010203"
    TRANSPARENT_RGBA = (1, 2, 3, 255)

    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    if sys.platform == "win32":
        root.attributes("-transparentcolor", TRANSPARENT_KEY)
    root.configure(bg=TRANSPARENT_KEY)

    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    _, _, taskbar_top = get_taskbar_inside_coords(sw, sh, 40)
    taskbar_height = max(40, sh - taskbar_top)
    SIZE = max(38, min(50, taskbar_height - 6))
    base_x, base_y, taskbar_top = get_taskbar_inside_coords(sw, sh, SIZE)
    root.geometry(f"{SIZE}x{SIZE}+{base_x}+{base_y}")

    lbl = tk.Label(root, bg=TRANSPARENT_KEY, borderwidth=0, cursor="hand2")
    lbl.pack()

    root.update_idletasks()
    hwnd = 0
    if sys.platform == "win32":
        GA_ROOT = 2
        hwnd = ctypes.windll.user32.GetAncestor(root.winfo_id(), GA_ROOT) or ctypes.windll.user32.GetParent(root.winfo_id()) or root.winfo_id()
        attach_to_taskbar_zorder(hwnd)
        pin_window_above_taskbar(hwnd)

    anim = {
        "tick": 0,
        "photo": None,
        "x": base_x,
        "y": base_y,
        "home_x": base_x,
        "dir": -1,
        "walk_enabled": True,
        "dragging": False,
        "drag_moved": False,
        "drag_start_x": 0,
        "drag_start_y": 0,
        "popup": None,
        "hoodie_rgb": get_target_hoodie_colors(load_system_state().get("mascot_state", "focused"))[0],
    }

    def animate_loop():
        cur = load_system_state()
        if not cur.get("badge_visible", True):
            root.destroy()
            return

        m_state = cur.get("mascot_state", "focused")
        tick = anim["tick"]
        anim["tick"] = tick + 1

        # Smooth color-transition animation across Cyber-Slime modes
        target_hoodie = get_target_hoodie_colors(m_state)[0]
        cur_hoodie = anim["hoodie_rgb"]
        if cur_hoodie != target_hoodie:
            next_rgb = lerp_rgb(cur_hoodie, target_hoodie, 0.30)
            if sum(abs(a - b) for a, b in zip(next_rgb, target_hoodie)) <= 6:
                next_rgb = target_hoodie
            anim["hoodie_rgb"] = next_rgb

        f_idx = tick
        facing_right = bool(anim["walk_enabled"] and m_state in ("focused", "working", "hydration", "mail") and anim["dir"] > 0)
        pil_img = render_byte_image(
            m_state,
            size=SIZE,
            frame_idx=f_idx,
            transparent_bg=TRANSPARENT_RGBA,
            hoodie_rgb=anim["hoodie_rgb"],
            facing_right=facing_right,
        )

        photo = ImageTk.PhotoImage(pil_img)
        anim["photo"] = photo
        lbl.config(image=photo)

        # Gentle horizontal hop-stroll inside the taskbar strip
        if anim["walk_enabled"] and not anim["dragging"] and m_state in ("focused", "working", "hydration", "mail"):
            phase = f_idx % 16
            should_step = (2 <= phase <= 5) if m_state in ("focused", "hydration", "mail") else (tick % 2 == 0)
            if should_step:
                step = 2 if (m_state == "working" or phase in (3, 4)) else 1
                anim["x"] += anim["dir"] * step
                if abs(anim["x"] - anim["home_x"]) >= 24:
                    anim["dir"] *= -1
                root.geometry(f"{SIZE}x{SIZE}+{anim['x']}+{anim['y']}")

        # Keep Byte pinned above Windows 11 taskbar z-order
        pin_window_above_taskbar(hwnd or root.winfo_id())

        if anim["hoodie_rgb"] != target_hoodie or m_state in ("working", "alert"):
            interval_ms = 85
        elif m_state == "paused":
            interval_ms = 125
        else:
            interval_ms = 90
        root.after(interval_ms, animate_loop)

    def on_drag_start(event):
        anim["dragging"] = True
        anim["drag_moved"] = False
        anim["drag_start_x"] = event.x
        anim["drag_start_y"] = event.y

    def on_drag_motion(event):
        dx = event.x - anim["drag_start_x"]
        dy = event.y - anim["drag_start_y"]
        if abs(dx) > 3 or abs(dy) > 3:
            anim["drag_moved"] = True
        new_x = root.winfo_x() + dx
        new_y = root.winfo_y() + dy
        anim["x"] = new_x
        anim["y"] = new_y
        anim["home_x"] = new_x
        root.geometry(f"{SIZE}x{SIZE}+{new_x}+{new_y}")

    def on_drag_end(event):
        was_moved = anim.get("drag_moved", False)
        anim["dragging"] = False
        anim["drag_moved"] = False
        if not was_moved:
            show_custom_hud_menu(event)

    def on_double_click(event):
        toggle_paused()

    def close_popup():
        p = anim.get("popup")
        if p is not None:
            try:
                p.destroy()
            except Exception:
                pass
            anim["popup"] = None

    def show_custom_hud_menu(event):
        """Renders a sleek, borderless Dark Cyber-HUD Popover card right above the taskbar."""
        close_popup()

        cur = load_system_state()
        m_state = cur.get("mascot_state", "focused").lower()
        paused = bool(cur.get("paused", False))
        muted = is_voice_muted()

        accent_map = {
            "focused":   ("#22D3EE", "FOCUSED // ONLINE"),
            "working":   ("#FBBF24", "WORKING // BUSY"),
            "alert":     ("#EF4444", "ALERT // DEADLINE"),
            "paused":    ("#94A3B8", "PAUSED // RESTING"),
            "hydration": ("#38BDF8", "HYDRATION // WATER"),
            "mail":      ("#2DD4BF", "MAIL // BRIEFING"),
        }
        dot_color, status_text = accent_map.get(m_state, ("#22D3EE", "ONLINE"))

        pop = tk.Toplevel(root)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#272732", padx=1, pady=1)
        anim["popup"] = pop

        card = tk.Frame(pop, bg="#111116", padx=10, pady=10)
        card.pack(fill="both", expand=True)

        # Header with glowing status pill
        hdr = tk.Frame(card, bg="#181820", padx=10, pady=8)
        hdr.pack(fill="x", pady=(0, 6))

        tk.Label(
            hdr,
            text="●",
            bg="#181820",
            fg=dot_color,
            font=("Segoe UI", 10, "bold")
        ).pack(side="left")

        tk.Label(
            hdr,
            text=f"  DITTO  ·  {status_text}",
            bg="#181820",
            fg="#F4F4F5",
            font=("Segoe UI", 9, "bold")
        ).pack(side="left")

        def add_menu_row(icon_char: str, label_text: str, callback, danger: bool = False):
            row = tk.Frame(card, bg="#111116", padx=8, pady=6, cursor="hand2")
            row.pack(fill="x", pady=1)

            fg_norm = "#F87171" if danger else "#E4E4E7"
            bg_hover = "#3B181E" if danger else "#1E293B"
            fg_hover = "#FCA5A5" if danger else "#38BDF8"

            lbl_ico = tk.Label(row, text=icon_char, bg="#111116", fg=dot_color if not danger else "#F87171", font=("Segoe UI", 9), width=3, anchor="w")
            lbl_ico.pack(side="left")

            lbl_txt = tk.Label(row, text=label_text, bg="#111116", fg=fg_norm, font=("Segoe UI", 9), anchor="w")
            lbl_txt.pack(side="left", fill="x", expand=True)

            def on_enter(_):
                row.configure(bg=bg_hover)
                lbl_ico.configure(bg=bg_hover, fg=fg_hover)
                lbl_txt.configure(bg=bg_hover, fg="#FFFFFF")

            def on_leave(_):
                row.configure(bg="#111116")
                lbl_ico.configure(bg="#111116", fg=dot_color if not danger else "#F87171")
                lbl_txt.configure(bg="#111116", fg=fg_norm)

            def on_click(_):
                close_popup()
                callback()

            for w in (row, lbl_ico, lbl_txt):
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", on_click)

        def act_pause():
            toggle_paused()

        def act_mute():
            toggle_mute()

        def act_walk():
            anim["walk_enabled"] = not anim["walk_enabled"]

        def _spawn_daemon_tool(script_name: str, *args: str):
            creation = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            subprocess.Popen(
                [sys.executable, str(Path(__file__).resolve().parent / script_name), *args],
                creationflags=creation,
            )

        def act_planner():
            _spawn_daemon_tool("planner_ui.py")

        def act_quick_ask():
            _spawn_daemon_tool("hud.py", "--quick-ask")

        def act_research():
            _spawn_daemon_tool("hud.py", "--quick-research")

        def act_quick_note():
            _spawn_daemon_tool("hud.py", "--quick-note")

        def act_organize():
            _spawn_daemon_tool("organizer.py", "--preview")

        def act_office_gen():
            _spawn_daemon_tool("office_gen.py", "--prompt", "doc")

        def act_undo_organize():
            def _run_undo():
                try:
                    r = httpx.post(f"{DAEMON_BASE}/command", json={"cmd": "undo", "arg": ""}, timeout=15.0)
                    data = r.json()
                    notify_windows("Chitran — Undo Organize", data.get("message", "Undo finished."))
                except Exception as exc:
                    notify_windows("Chitran Error", f"Undo organize failed: {exc}")
            threading.Thread(target=_run_undo, daemon=True).start()

        def act_fill_form():
            def _run_fill():
                try:
                    r = httpx.post(f"{DAEMON_BASE}/command", json={"cmd": "fill", "arg": ""}, timeout=20.0)
                    data = r.json()
                    if data.get("status") == "success":
                        notify_windows("Chitran — Form Auto-Fill", f"Filled {data.get('filled_count', 0)} field(s). Review & submit.")
                    else:
                        notify_windows("Chitran Notice", data.get("message", "Could not fill active form."))
                except Exception as exc:
                    notify_windows("Chitran Error", f"Auto-fill failed: {exc}")
            threading.Thread(target=_run_fill, daemon=True).start()

        def act_guide():
            _spawn_daemon_tool("hud.py", "--guide")

        def act_mail():
            def _run_mail():
                try:
                    r = httpx.post(f"{DAEMON_BASE}/command", json={"cmd": "mail", "arg": "hud"}, timeout=35.0)
                    data = r.json()
                    if data.get("status") != "success":
                        notify_windows("Chitran — Gmail Briefing", data.get("message", "Could not fetch Gmail."))
                except Exception as exc:
                    notify_windows("Chitran Error", f"Gmail briefing failed: {exc}")
            threading.Thread(target=_run_mail, daemon=True).start()

        def act_dashboard():
            webbrowser.open(f"{DAEMON_BASE}/dashboard")

        def act_kill():
            try:
                httpx.post(f"{DAEMON_BASE}/command", json={"cmd": "kill", "arg": ""}, timeout=2.0)
            except Exception:
                pass

        def act_hide():
            c = load_system_state()
            c["badge_visible"] = False
            save_system_state(c)
            root.destroy()

        def act_exit_all():
            try:
                if PID_FILE.exists():
                    parent_pid = int(PID_FILE.read_text().strip())
                    PID_FILE.unlink()
                    if sys.platform == "win32" and parent_pid > 0:
                        subprocess.Popen(
                            ["taskkill", "/PID", str(parent_pid), "/F"],
                            creationflags=subprocess.CREATE_NO_WINDOW,
                        )
            except Exception:
                pass
            root.destroy()

        add_menu_row("🎯", "Daily Goals & Planner", act_planner)
        add_menu_row("📬", "Gmail Briefing", act_mail)
        add_menu_row("🔬", "Quick Web Research", act_research)
        add_menu_row("💬", "Ask Senior Mentor", act_quick_ask)
        add_menu_row("📝", "Quick Daily Note", act_quick_note)
        add_menu_row("🌐", "Auto-Fill Active Form", act_fill_form)
        add_menu_row("📄", "Generate Office File (Doc/PPT/Excel)", act_office_gen)
        add_menu_row("📁", "Organize Downloads (Preview)", act_organize)
        add_menu_row("↩", "Undo Last Organize", act_undo_organize)
        add_menu_row("📖", "User Guide & Commands", act_guide)

        sep1 = tk.Frame(card, bg="#272732", height=1)
        sep1.pack(fill="x", pady=4)

        add_menu_row("⏯", "Resume Monitoring" if paused else "Pause Monitoring", act_pause)
        add_menu_row("🔊", "Unmute Voice" if muted else "Mute Voice", act_mute)
        add_menu_row("🫧", "Lock Ditto Position" if anim["walk_enabled"] else "Enable Taskbar Stroll", act_walk)
        add_menu_row("📊", "Analytics Dashboard", act_dashboard)
        add_menu_row("⚡", "Emergency Kill Tasks", act_kill, danger=True)

        sep2 = tk.Frame(card, bg="#272732", height=1)
        sep2.pack(fill="x", pady=4)

        add_menu_row("⬇", "Minimize to Tray (restore: cn ditto)", act_hide)
        add_menu_row("✕", "Exit Ditto Completely", act_exit_all, danger=True)

        pop.update_idletasks()
        pw = max(260, pop.winfo_reqwidth())
        ph = pop.winfo_reqheight()
        px = min(sw - pw - 14, max(14, root.winfo_x() - pw // 2 + SIZE // 2))
        py = max(14, taskbar_top - ph - 10)
        pop.geometry(f"{pw}x{ph}+{px}+{py}")
        apply_win11_dwm_style(pop, dark_titlebar=False, rounded=True)

        pop.bind("<Escape>", lambda _: close_popup())
        pop.bind("<FocusOut>", lambda _: close_popup())
        pop.focus_force()

    lbl.bind("<ButtonPress-1>", lambda e: (close_popup(), on_drag_start(e)))
    lbl.bind("<B1-Motion>", on_drag_motion)
    lbl.bind("<ButtonRelease-1>", on_drag_end)
    lbl.bind("<Double-Button-1>", on_double_click)
    lbl.bind("<Button-3>", show_custom_hud_menu)

    animate_loop()
    root.mainloop()


if __name__ == "__main__":
    if "--badge" in sys.argv:
        run_mini_badge()
    else:
        ByteTrayController().run()
