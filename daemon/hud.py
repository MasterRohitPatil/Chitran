import json
import os
import sys
import threading
import tkinter as tk
from tkinter import font as tkfont
import subprocess
import urllib.request
from pathlib import Path
from typing import Optional

from win_ui import apply_win11_dwm_style, enable_high_dpi, render_markdown_into_text

DAEMON_COMMAND_URL = "http://127.0.0.1:7777/command"
USER_GUIDE_PATH = Path(__file__).resolve().parent.parent / "USER_GUIDE.md"


def _post_to_daemon_or_fallback(cmd: str, arg: str) -> dict:
    payload = json.dumps({"cmd": cmd, "arg": arg}).encode("utf-8")
    req = urllib.request.Request(
        DAEMON_COMMAND_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=25.0) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"status": "error", "message": f"Daemon unreachable: {e}"}


def show_quick_input_hud(mode: str):
    """Compact bottom-left dark Cyber-HUD prompt for Quick Note, Ask Mentor, or Web Research."""
    enable_high_dpi()
    mode_cfg = {
        "note": (
            "Chitran — Quick Note",
            "●  QUICK DAILY NOTE",
            "Saves to ~/Daily_Notes/YYYY-MM-DD.md",
            "Save Note",
            "Enter to Submit  ·  Esc to Cancel",
        ),
        "research": (
            "Chitran — Web Research",
            "●  PARALLEL WEB RESEARCHER",
            "Scrapes top technical sources & creates Claude/ChatGPT Context Pack",
            "Build Context Pack",
            "Enter to Research  ·  Esc to Cancel",
        ),
        "ask": (
            "Chitran — Ask Mentor",
            "●  ASK SENIOR MENTOR",
            "Routes to Groq / Gemini (₹0 free tier) & opens answer HUD",
            "Ask Mentor",
            "Enter to Submit  ·  Esc to Cancel",
        ),
    }
    win_title, badge_txt, sub_txt, btn_txt, foot_txt = mode_cfg.get(mode, mode_cfg["ask"])
    root = tk.Tk()
    root.title(win_title)
    root.attributes("-topmost", True)
    root.configure(bg="#121214")
    screen_height = root.winfo_screenheight()
    width, height = 540, 225
    root.geometry(f"{width}x{height}+24+{max(40, screen_height - height - 68)}")
    root.resizable(False, False)
    apply_win11_dwm_style(root)

    hdr = tk.Frame(root, bg="#18181B", padx=16, pady=10)
    hdr.pack(fill="x")
    tk.Label(
        hdr,
        text=badge_txt,
        bg="#18181B",
        fg="#22D3EE",
        font=("Segoe UI", 9, "bold"),
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        hdr,
        text=sub_txt,
        bg="#18181B",
        fg="#A1A1AA",
        font=("Segoe UI", 9),
        anchor="w",
    ).pack(fill="x", pady=(2, 0))

    body = tk.Frame(root, bg="#121214", padx=16, pady=12)
    body.pack(fill="both", expand=True)

    text_var = tk.StringVar()
    entry = tk.Entry(
        body,
        textvariable=text_var,
        bg="#202026",
        fg="#F4F4F5",
        insertbackground="#F4F4F5",
        relief="flat",
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground="#272732",
        highlightcolor="#22D3EE",
    )
    entry.pack(fill="x", ipady=6)

    status_lbl = tk.Label(body, text="", bg="#121214", fg="#A1A1AA", font=("Segoe UI", 9), anchor="w")
    status_lbl.pack(fill="x", pady=(8, 0))

    def submit():
        val = text_var.get().strip()
        if not val:
            status_lbl.config(text="Please type something first.", fg="#F87171")
            return

        if mode == "note":
            res = _post_to_daemon_or_fallback("note", val)
            if res.get("status") != "success":
                try:
                    from notes import append_daily_note
                    res = append_daily_note(val)
                except Exception as exc:
                    status_lbl.config(text=f"Error: {exc}", fg="#F87171")
                    return
            try:
                from notify import notify_windows
                notify_windows("Chitran — Note Saved", res.get("entry", val))
            except Exception:
                pass
            root.destroy()
            return

        if mode == "research":
            status_lbl.config(text="Searching & scraping sources in parallel...", fg="#FBBF24")
            submit_btn.config(state="disabled")

            def res_worker():
                res = _post_to_daemon_or_fallback("research", val)
                if res.get("status") == "error":
                    try:
                        import asyncio
                        from researcher import research_topic
                        res = asyncio.run(research_topic(val))
                    except Exception as exc:
                        res = {"status": "error", "message": f"Research error: {exc}"}

                def _open_pack():
                    root.destroy()
                    if res.get("file"):
                        subprocess.Popen(
                            [sys.executable, str(Path(__file__).resolve()), "--context-pack", res["file"]],
                            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                        )
                    else:
                        spawn_hud_process("Research Error", val, res.get("message", "Unable to complete research."))

                root.after(0, _open_pack)

            threading.Thread(target=res_worker, daemon=True).start()
            return

        status_lbl.config(text="Thinking... waiting for mentor reply.", fg="#FBBF24")
        submit_btn.config(state="disabled")

        def worker():
            res = _post_to_daemon_or_fallback("ask", val)
            if res.get("status") == "error":
                try:
                    import asyncio
                    from router import query_mentor
                    res = asyncio.run(query_mentor(val))
                except Exception as exc:
                    res = {"answer": f"Mentor error: {exc}"}
            answer = res.get("answer") or res.get("message") or "No response returned."
            def _open_result():
                root.destroy()
                spawn_hud_process("Mentor Review", val, answer)
            root.after(0, _open_result)

        threading.Thread(target=worker, daemon=True).start()

    footer = tk.Frame(root, bg="#18181B", padx=16, pady=10)
    footer.pack(fill="x", side="bottom")
    tk.Label(footer, text=foot_txt, bg="#18181B", fg="#71717A", font=("Segoe UI", 9)).pack(side="left")

    submit_btn = tk.Button(
        footer,
        text=btn_txt,
        command=submit,
        bg="#22D3EE",
        fg="#09090B",
        activebackground="#06B6D4",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=14,
        pady=4,
        cursor="hand2",
    )
    submit_btn.pack(side="right")

    root.bind("<Return>", lambda _: submit())
    root.bind("<Escape>", lambda _: root.destroy())
    root.focus_force()
    entry.focus_set()
    root.mainloop()


def create_hud_window(
    title: str,
    query: str,
    content: str,
    copy_text: Optional[str] = None,
    copy_label: str = "Copy to Clipboard",
    open_file_path: Optional[str] = None,
):
    """Spawns an always-on-top, dark-mode floating HUD window with crisp High-DPI & Markdown styling."""
    enable_high_dpi()
    root = tk.Tk()
    root.title(f"Chitran — {title}")
    root.attributes("-topmost", True)
    root.configure(bg="#121214")

    # Geometry & positioning (docked bottom-left so center of screen stays unobstructed)
    width = 680
    height = 500
    screen_height = root.winfo_screenheight()
    x = 24
    y = max(40, screen_height - height - 68)
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.minsize(500, 340)
    apply_win11_dwm_style(root)

    # Header Frame
    header_frame = tk.Frame(root, bg="#18181B", padx=18, pady=12)
    header_frame.pack(fill="x", side="top")

    tk.Label(
        header_frame,
        text=f"●  {title.upper()}",
        bg="#18181B",
        fg="#22D3EE",
        font=("Segoe UI", 9, "bold"),
        anchor="w",
    ).pack(fill="x")

    lbl_title = tk.Label(
        header_frame,
        text=query,
        bg="#18181B",
        fg="#F4F4F5",
        font=("Segoe UI", 11, "bold"),
        anchor="w",
        wraplength=630,
        justify="left",
    )
    lbl_title.pack(fill="x", pady=(2, 0))

    # Content Frame
    content_frame = tk.Frame(root, bg="#121214", padx=16, pady=12)
    content_frame.pack(fill="both", expand=True)

    text_area = tk.Text(
        content_frame,
        wrap="word",
        bg="#16161D",
        fg="#F4F4F5",
        insertbackground="#FFFFFF",
        font=("Segoe UI", 10),
        relief="flat",
        padx=16,
        pady=14,
        borderwidth=0,
        highlightthickness=1,
        highlightbackground="#272732",
    )
    render_markdown_into_text(text_area, content)

    scrollbar = tk.Scrollbar(content_frame, command=text_area.yview, bg="#27272A")
    text_area.config(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    text_area.pack(side="left", fill="both", expand=True)

    # Footer Frame
    footer_frame = tk.Frame(root, bg="#18181B", padx=16, pady=10)
    footer_frame.pack(fill="x", side="bottom")

    status_lbl = tk.Label(
        footer_frame,
        text="Press [Esc] to dismiss",
        bg="#18181B",
        fg="#71717A",
        font=("Segoe UI", 9)
    )
    status_lbl.pack(side="left")

    copy_payload = copy_text if copy_text is not None else content
    use_primary = copy_text is not None
    default_bg = "#22D3EE" if use_primary else "#27272A"
    default_fg = "#09090B" if use_primary else "#E4E4E7"

    def copy_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(copy_payload)
        root.update()
        copy_btn.config(text="✓ Copied Prompt!" if use_primary else "✓ Copied!", bg="#10B981", fg="#09090B")
        root.after(2000, lambda: copy_btn.config(text=copy_label, bg=default_bg, fg=default_fg))

    copy_btn = tk.Button(
        footer_frame,
        text=copy_label,
        command=copy_to_clipboard,
        bg=default_bg,
        fg=default_fg,
        activebackground="#06B6D4" if use_primary else "#3F3F46",
        activeforeground="#09090B" if use_primary else "#FFFFFF",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=4,
        cursor="hand2"
    )
    copy_btn.pack(side="right")

    if open_file_path and Path(open_file_path).exists():
        def open_pack_file():
            try:
                if sys.platform == "win32":
                    os.startfile(open_file_path)
            except Exception:
                pass

        tk.Button(
            footer_frame,
            text="Open .md",
            command=open_pack_file,
            bg="#27272A",
            fg="#E4E4E7",
            activebackground="#3F3F46",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9),
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side="right", padx=(0, 8))

    # Dismiss keybinds
    root.bind("<Escape>", lambda event: root.destroy())

    # Focus window
    root.focus_force()
    root.mainloop()


def spawn_hud_process(title: str, query: str, content: str):
    """Spawns the HUD window in an isolated process to avoid blocking the daemon."""
    hud_py = Path(__file__).resolve()
    subprocess.Popen(
        [sys.executable, str(hud_py), title, query, content],
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    )


if __name__ == "__main__":
    arg1 = sys.argv[1] if len(sys.argv) > 1 else ""
    if arg1 == "--quick-note":
        show_quick_input_hud("note")
    elif arg1 == "--quick-ask":
        show_quick_input_hud("ask")
    elif arg1 == "--quick-research":
        show_quick_input_hud("research")
    elif arg1 == "--context-pack" and len(sys.argv) > 2:
        pack_path = Path(sys.argv[2])
        if pack_path.exists():
            from researcher import extract_prompt_from_pack
            raw_md = pack_path.read_text(encoding="utf-8")
            prompt_only = extract_prompt_from_pack(raw_md)
            first_line = raw_md.splitlines()[0].lstrip("# ").strip() if raw_md.splitlines() else "Context Pack"
            create_hud_window(
                "Research Context Pack",
                f"{first_line}  ·  Saved to {pack_path.name}",
                raw_md,
                copy_text=prompt_only,
                copy_label="Copy Prompt for Claude/ChatGPT",
                open_file_path=str(pack_path),
            )
        else:
            create_hud_window("Research Error", str(pack_path), "Context pack file was not found.")
    elif arg1 == "--guide":
        guide_text = USER_GUIDE_PATH.read_text(encoding="utf-8") if USER_GUIDE_PATH.exists() else "USER_GUIDE.md not found."
        create_hud_window("User Guide & Commands", "Chitran & Ditto Cheat Sheet", guide_text)
    else:
        t = arg1 or "Mentor Response"
        q = sys.argv[2] if len(sys.argv) > 2 else ""
        c = sys.argv[3] if len(sys.argv) > 3 else "No content returned."
        create_hud_window(t, q, c)

