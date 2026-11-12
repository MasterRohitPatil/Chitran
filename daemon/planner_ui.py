import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from typing import Optional

from notify import notify_windows
from planner import (
    add_goal,
    get_today_goals,
    mark_goal_done,
    normalize_time_str,
    refresh_mascot_goal_state,
    today_iso,
)
from voice import is_voice_muted, speak_background
from win_ui import apply_win11_dwm_style, enable_high_dpi

# Dark Cyber-HUD palette matching hud.py & mascot.py
BG_ROOT = "#121214"
BG_CARD = "#18181B"
BG_INPUT = "#202026"
BORDER_COLOR = "#272732"
FG_PRIMARY = "#F4F4F5"
FG_MUTED = "#A1A1AA"
FG_SUBTLE = "#71717A"
ACCENT_CYAN = "#22D3EE"
ACCENT_GREEN = "#10B981"
ACCENT_AMBER = "#FBBF24"
ACCENT_RED = "#F87171"


def _quick_time_offsets() -> list[tuple[str, str]]:
    now = datetime.now()
    plus1 = (now + timedelta(hours=1)).strftime("%H:%M")
    plus2 = (now + timedelta(hours=2)).strftime("%H:%M")
    return [
        (f"+1 hr ({plus1})", plus1),
        (f"+2 hrs ({plus2})", plus2),
        ("18:00", "18:00"),
        ("21:00", "21:00"),
        ("23:00", "23:00"),
    ]


def _dock_bottom_left(root: tk.Tk, width: int, height: int) -> None:
    screen_height = root.winfo_screenheight()
    x = 24
    y = max(40, screen_height - height - 68)
    root.geometry(f"{width}x{height}+{x}+{y}")


def show_deadline_prompt(task_title: str) -> Optional[dict]:
    """Inline Cyber-HUD popup asking only for the missing deadline of a task."""
    enable_high_dpi()
    saved_goal = {"value": None}
    if not is_voice_muted():
        speak_background("Deadline missing. Please pick a target time.")

    root = tk.Tk()
    root.title("Chitran — Set Goal Deadline")
    root.attributes("-topmost", True)
    root.configure(bg=BG_ROOT)
    _dock_bottom_left(root, 520, 285)
    root.resizable(False, False)
    apply_win11_dwm_style(root)

    header = tk.Frame(root, bg=BG_CARD, padx=16, pady=12)
    header.pack(fill="x")
    tk.Label(
        header,
        text="●  DEADLINE REQUIRED",
        bg=BG_CARD,
        fg=ACCENT_CYAN,
        font=("Segoe UI", 9, "bold"),
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        header,
        text=f"Goal: {task_title}",
        bg=BG_CARD,
        fg=FG_PRIMARY,
        font=("Segoe UI", 11, "bold"),
        anchor="w",
        wraplength=450,
        justify="left",
    ).pack(fill="x", pady=(4, 0))

    body = tk.Frame(root, bg=BG_ROOT, padx=16, pady=12)
    body.pack(fill="both", expand=True)

    tk.Label(
        body,
        text="Enter deadline (e.g. 21:30 or 9:30 pm) or click a quick preset:",
        bg=BG_ROOT,
        fg=FG_MUTED,
        font=("Segoe UI", 9),
        anchor="w",
    ).pack(fill="x")

    default_time = (datetime.now() + timedelta(hours=2)).strftime("%H:%M")
    time_var = tk.StringVar(value=default_time)
    entry = tk.Entry(
        body,
        textvariable=time_var,
        bg=BG_INPUT,
        fg=FG_PRIMARY,
        insertbackground=FG_PRIMARY,
        relief="flat",
        font=("Consolas", 12, "bold"),
        highlightthickness=1,
        highlightbackground=BORDER_COLOR,
        highlightcolor=ACCENT_CYAN,
    )
    entry.pack(fill="x", ipady=6, pady=(6, 10))
    entry.select_range(0, tk.END)

    chips_frame = tk.Frame(body, bg=BG_ROOT)
    chips_frame.pack(fill="x", pady=(0, 8))

    err_lbl = tk.Label(body, text="", bg=BG_ROOT, fg=ACCENT_RED, font=("Segoe UI", 9), anchor="w")
    err_lbl.pack(fill="x")

    def submit(chosen_time: Optional[str] = None):
        raw_t = (chosen_time if chosen_time is not None else time_var.get()).strip()
        norm = normalize_time_str(raw_t)
        if not norm:
            err_lbl.config(text="Invalid time format. Try 21:30 or 9:30 pm.")
            return
        goal = add_goal(task_title, norm)
        refresh_mascot_goal_state()
        saved_goal["value"] = goal
        notify_windows("Chitran — Goal Locked", f"#{goal['id']} {goal['title']} @ {goal['deadline_time']}")
        if not is_voice_muted():
            speak_background(f"Goal locked: {goal['title']} by {goal['deadline_time']}.")
        root.destroy()

    for label, hhmm in _quick_time_offsets():
        btn = tk.Button(
            chips_frame,
            text=label,
            command=lambda t=hhmm: submit(t),
            bg=BG_CARD,
            fg=FG_PRIMARY,
            activebackground="#27273A",
            activeforeground=ACCENT_CYAN,
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=8,
            pady=4,
            cursor="hand2",
        )
        btn.pack(side="left", padx=(0, 6))

    footer = tk.Frame(root, bg=BG_CARD, padx=16, pady=10)
    footer.pack(fill="x", side="bottom")
    tk.Label(footer, text="Enter to Lock  ·  Esc to Cancel", bg=BG_CARD, fg=FG_SUBTLE, font=("Segoe UI", 9)).pack(side="left")

    tk.Button(
        footer,
        text="Lock Goal",
        command=submit,
        bg=ACCENT_CYAN,
        fg="#09090B",
        activebackground="#06B6D4",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=14,
        pady=4,
        cursor="hand2",
    ).pack(side="right")

    root.bind("<Return>", lambda _: submit())
    root.bind("<Escape>", lambda _: root.destroy())
    root.focus_force()
    entry.focus_set()
    root.mainloop()
    return saved_goal["value"]


def show_planner_hud(startup_mode: bool = False) -> None:
    """Interactive Cyber-HUD showing today's goals and an inline New Goal + Deadline bar."""
    enable_high_dpi()
    root = tk.Tk()
    root.title("Chitran — Daily Goal Planner")
    root.attributes("-topmost", True)
    root.configure(bg=BG_ROOT)
    _dock_bottom_left(root, 660, 490)
    root.minsize(540, 390)
    apply_win11_dwm_style(root)

    header = tk.Frame(root, bg=BG_CARD, padx=16, pady=12)
    header.pack(fill="x")

    banner_txt = (
        f"☀  MORNING CHECK-IN  ·  NO GOAL LOGGED YET ({today_iso()})"
        if startup_mode
        else f"●  DAILY GOAL PLANNER  ·  {today_iso()}"
    )
    tk.Label(
        header,
        text=banner_txt,
        bg=BG_CARD,
        fg=ACCENT_AMBER if startup_mode else ACCENT_CYAN,
        font=("Segoe UI", 9, "bold"),
        anchor="w",
    ).pack(fill="x")

    if startup_mode:
        tk.Label(
            header,
            text="What is your primary target today? Lock a goal and deadline below.",
            bg=BG_CARD,
            fg=FG_MUTED,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(3, 0))

    # Add Goal Form
    form_card = tk.Frame(root, bg=BG_CARD, padx=14, pady=10, highlightthickness=1, highlightbackground=BORDER_COLOR)
    form_card.pack(fill="x", padx=14, pady=(12, 8))

    row1 = tk.Frame(form_card, bg=BG_CARD)
    row1.pack(fill="x")

    tk.Label(row1, text="Goal:", bg=BG_CARD, fg=FG_MUTED, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))
    goal_var = tk.StringVar()
    goal_entry = tk.Entry(
        row1,
        textvariable=goal_var,
        bg=BG_INPUT,
        fg=FG_PRIMARY,
        insertbackground=FG_PRIMARY,
        relief="flat",
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=BORDER_COLOR,
        highlightcolor=ACCENT_CYAN,
    )
    goal_entry.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))

    tk.Label(row1, text="@", bg=BG_CARD, fg=ACCENT_CYAN, font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 4))
    time_var = tk.StringVar(value=(datetime.now() + timedelta(hours=2)).strftime("%H:%M"))
    time_entry = tk.Entry(
        row1,
        textvariable=time_var,
        width=9,
        bg=BG_INPUT,
        fg=FG_PRIMARY,
        insertbackground=FG_PRIMARY,
        relief="flat",
        font=("Consolas", 10, "bold"),
        highlightthickness=1,
        highlightbackground=BORDER_COLOR,
        highlightcolor=ACCENT_CYAN,
    )
    time_entry.pack(side="left", ipady=4, padx=(0, 8))

    add_btn = tk.Button(
        row1,
        text="+ Lock Goal",
        bg=ACCENT_CYAN,
        fg="#09090B",
        activebackground="#06B6D4",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=10,
        pady=3,
        cursor="hand2",
    )
    add_btn.pack(side="right")

    chips = tk.Frame(form_card, bg=BG_CARD)
    chips.pack(fill="x", pady=(8, 0))
    tk.Label(chips, text="Presets:", bg=BG_CARD, fg=FG_SUBTLE, font=("Segoe UI", 8)).pack(side="left", padx=(0, 6))
    for label, hhmm in _quick_time_offsets():
        tk.Button(
            chips,
            text=label,
            command=lambda t=hhmm: time_var.set(t),
            bg=BG_INPUT,
            fg=FG_MUTED,
            activebackground="#2D2D3A",
            activeforeground=FG_PRIMARY,
            font=("Segoe UI", 8),
            relief="flat",
            padx=6,
            pady=2,
            cursor="hand2",
        ).pack(side="left", padx=(0, 5))

    status_lbl = tk.Label(root, text="", bg=BG_ROOT, fg=ACCENT_GREEN, font=("Segoe UI", 9), anchor="w", padx=16)
    status_lbl.pack(fill="x")

    # Scrollable Goal List Frame
    list_container = tk.Frame(root, bg=BG_ROOT, padx=14, pady=6)
    list_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(list_container, bg=BG_ROOT, highlightthickness=0)
    scrollbar = tk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
    goals_frame = tk.Frame(canvas, bg=BG_ROOT)

    goals_frame.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
    window_id = canvas.create_window((0, 0), window=goals_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def render_goals():
        for child in goals_frame.winfo_children():
            child.destroy()

        goals = get_today_goals()
        now_hhmm = datetime.now().strftime("%H:%M")

        if not goals:
            tk.Label(
                goals_frame,
                text="No goals logged for today yet. Add one above!",
                bg=BG_ROOT,
                fg=FG_SUBTLE,
                font=("Segoe UI", 10),
                pady=24,
            ).pack(fill="x")
            return

        for g in goals:
            is_done = g["status"] == "completed"
            is_overdue = (not is_done) and bool(g.get("deadline_time")) and (g["deadline_time"] < now_hhmm)

            row = tk.Frame(
                goals_frame,
                bg=BG_CARD,
                padx=12,
                pady=8,
                highlightthickness=1,
                highlightbackground=ACCENT_RED if is_overdue else BORDER_COLOR,
            )
            row.pack(fill="x", pady=3)

            badge_color = ACCENT_GREEN if is_done else (ACCENT_RED if is_overdue else ACCENT_CYAN)
            badge_text = "DONE" if is_done else ("OVERDUE" if is_overdue else "PENDING")

            tk.Label(
                row,
                text=f"#{g['id']}  [{badge_text}]",
                bg=BG_CARD,
                fg=badge_color,
                font=("Consolas", 9, "bold"),
                width=14,
                anchor="w",
            ).pack(side="left")

            tk.Label(
                row,
                text=g["title"],
                bg=BG_CARD,
                fg=FG_SUBTLE if is_done else FG_PRIMARY,
                font=("Segoe UI", 10, "overstrike" if is_done else "normal"),
                anchor="w",
                wraplength=310,
                justify="left",
            ).pack(side="left", fill="x", expand=True)

            tk.Label(
                row,
                text=f"@ {g.get('deadline_time') or '--:--'}",
                bg=BG_CARD,
                fg=ACCENT_RED if is_overdue else FG_MUTED,
                font=("Consolas", 9, "bold"),
                padx=8,
            ).pack(side="left")

            if not is_done:
                tk.Button(
                    row,
                    text="✓ Done",
                    command=lambda gid=g["id"]: complete_goal(gid),
                    bg="#14332B",
                    fg=ACCENT_GREEN,
                    activebackground=ACCENT_GREEN,
                    activeforeground="#09090B",
                    font=("Segoe UI", 8, "bold"),
                    relief="flat",
                    padx=8,
                    pady=2,
                    cursor="hand2",
                ).pack(side="right")

    def complete_goal(goal_id: int):
        res = mark_goal_done(str(goal_id))
        status_lbl.config(text=res.get("message", "Completed."), fg=ACCENT_GREEN)
        render_goals()

    def submit_new_goal():
        title = goal_var.get().strip()
        raw_t = time_var.get().strip()
        if not title:
            status_lbl.config(text="Please enter a goal description first.", fg=ACCENT_RED)
            return
        norm = normalize_time_str(raw_t)
        if not norm:
            status_lbl.config(text="Invalid time. Use HH:MM (e.g. 21:30 or 9:30 pm).", fg=ACCENT_RED)
            return
        g = add_goal(title, norm)
        refresh_mascot_goal_state()
        goal_var.set("")
        status_lbl.config(text=f"Locked #{g['id']}: {g['title']} @ {g['deadline_time']}", fg=ACCENT_GREEN)
        if not is_voice_muted():
            speak_background(f"Goal locked: {g['title']} by {g['deadline_time']}.")
        render_goals()

    add_btn.config(command=submit_new_goal)
    goal_entry.bind("<Return>", lambda _: submit_new_goal())
    time_entry.bind("<Return>", lambda _: submit_new_goal())

    footer = tk.Frame(root, bg=BG_CARD, padx=16, pady=10)
    footer.pack(fill="x", side="bottom")
    tk.Label(footer, text="Press [Esc] to close", bg=BG_CARD, fg=FG_SUBTLE, font=("Segoe UI", 9)).pack(side="left")
    tk.Button(
        footer,
        text="Close",
        command=root.destroy,
        bg="#27272A",
        fg=FG_PRIMARY,
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=3,
        cursor="hand2",
    ).pack(side="right")

    render_goals()
    root.bind("<Escape>", lambda _: root.destroy())
    root.focus_force()
    goal_entry.focus_set()
    root.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sidekick OS Daily Goal & Deadline HUD")
    parser.add_argument("--deadline-for", type=str, default="", help="Prompt only for deadline of a given task")
    parser.add_argument("--startup", action="store_true", help="Show morning/startup check-in banner")
    args = parser.parse_args()

    if args.deadline_for.strip():
        show_deadline_prompt(args.deadline_for.strip())
    else:
        show_planner_hud(startup_mode=args.startup)

