import ctypes
import re
import sys
import tkinter as tk


def enable_high_dpi() -> None:
    """
    Enables Per-Monitor v2 / System High-DPI awareness on Windows 11 so Tkinter
    renders fonts and borders at native physical pixels instead of blurry bitmap scaling.
    Must be called BEFORE creating `tk.Tk()`.
    """
    if sys.platform != "win32":
        return
    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def get_hwnd(widget: tk.Misc) -> int:
    """Returns the top-level Win32 HWND for a Tk or Toplevel window."""
    if sys.platform != "win32":
        return 0
    try:
        widget.update_idletasks()
        GA_ROOT = 2
        wid = widget.winfo_id()
        return (
            ctypes.windll.user32.GetAncestor(wid, GA_ROOT)
            or ctypes.windll.user32.GetParent(wid)
            or wid
        )
    except Exception:
        return 0


def apply_win11_dwm_style(
    window: tk.Misc,
    dark_titlebar: bool = True,
    rounded: bool = True,
    caption_color_bgr: int = 0x00181414,  # #141418 in COLORREF (0x00BBGGRR)
    border_color_bgr: int = 0x00382D2A,   # #2A2D38 subtle border
) -> None:
    """
    Applies native Windows 11 DWM attributes:
    - Dark mode title bar (DWMWA_USE_IMMERSIVE_DARK_MODE = 20)
    - Hardware-accelerated rounded corners & shadow (DWMWA_WINDOW_CORNER_PREFERENCE = 33)
    - Custom dark caption & border colors (34, 35)
    """
    if sys.platform != "win32":
        return
    hwnd = get_hwnd(window)
    if not hwnd:
        return
    try:
        dwmapi = ctypes.windll.dwmapi
        if dark_titlebar:
            val_dark = ctypes.c_int(1)
            dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(val_dark), ctypes.sizeof(val_dark))

        if rounded:
            # DWMWCP_ROUND = 2
            val_round = ctypes.c_int(2)
            dwmapi.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(val_round), ctypes.sizeof(val_round))

        val_border = ctypes.c_int(border_color_bgr)
        dwmapi.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(val_border), ctypes.sizeof(val_border))

        val_caption = ctypes.c_int(caption_color_bgr)
        dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(val_caption), ctypes.sizeof(val_caption))
    except Exception:
        pass


def render_markdown_into_text(text_widget: tk.Text, markdown_str: str) -> None:
    """
    Formats Markdown into a Tkinter Text widget with clean visual hierarchy:
    - H1 / H2 / H3 headings with accent colors
    - Markdown tables converted into clean aligned command/description rows
    - Bullet points with cyan bullets (•)
    - Inline `code` badges and **bold** spans
    - Fenced ```code``` blocks in monospace cards
    """
    text_widget.config(state="normal")
    text_widget.delete("1.0", tk.END)

    # Configure typography tags
    text_widget.tag_configure(
        "h1",
        font=("Segoe UI", 14, "bold"),
        foreground="#22D3EE",
        spacing1=6,
        spacing3=8,
    )
    text_widget.tag_configure(
        "h2",
        font=("Segoe UI", 11, "bold"),
        foreground="#38BDF8",
        spacing1=12,
        spacing3=4,
    )
    text_widget.tag_configure(
        "h3",
        font=("Segoe UI", 10, "bold"),
        foreground="#FBBF24",
        spacing1=10,
        spacing3=3,
    )
    text_widget.tag_configure(
        "body",
        font=("Segoe UI", 10),
        foreground="#E4E4E7",
        spacing1=2,
        spacing2=2,
        spacing3=3,
    )
    text_widget.tag_configure(
        "bold",
        font=("Segoe UI", 10, "bold"),
        foreground="#FFFFFF",
    )
    text_widget.tag_configure(
        "inline_code",
        font=("Consolas", 10, "bold"),
        foreground="#22D3EE",
        background="#222430",
    )
    text_widget.tag_configure(
        "code_block",
        font=("Consolas", 9),
        foreground="#A7F3D0",
        background="#181922",
        lmargin1=14,
        lmargin2=14,
        rmargin=14,
        spacing1=2,
        spacing3=2,
    )
    text_widget.tag_configure(
        "bullet_dot",
        font=("Segoe UI", 10, "bold"),
        foreground="#22D3EE",
        lmargin1=8,
        lmargin2=22,
    )
    text_widget.tag_configure(
        "bullet_body",
        font=("Segoe UI", 10),
        foreground="#E4E4E7",
        lmargin1=8,
        lmargin2=22,
        spacing1=2,
        spacing3=3,
    )
    text_widget.tag_configure(
        "table_key",
        font=("Consolas", 10, "bold"),
        foreground="#22D3EE",
        background="#1C1E28",
        lmargin1=10,
        lmargin2=10,
        spacing1=3,
    )
    text_widget.tag_configure(
        "table_val",
        font=("Segoe UI", 10),
        foreground="#D4D4D8",
        lmargin1=22,
        lmargin2=22,
        spacing3=5,
    )
    text_widget.tag_configure(
        "divider",
        font=("Consolas", 8),
        foreground="#2E303E",
        spacing1=6,
        spacing3=6,
    )

    def _insert_rich_spans(line_text: str, base_tag: str = "body") -> None:
        # Tokenize `inline code` and **bold**
        token_re = re.compile(r"(`[^`]+`|\*\*[^*]+\*\*)")
        pos = 0
        for m in token_re.finditer(line_text):
            if m.start() > pos:
                text_widget.insert(tk.END, line_text[pos : m.start()], (base_tag,))
            tok = m.group(0)
            if tok.startswith("`") and tok.endswith("`"):
                clean = tok[1:-1]
                text_widget.insert(tk.END, f" {clean} ", ("inline_code",))
            elif tok.startswith("**") and tok.endswith("**"):
                clean = tok[2:-2]
                text_widget.insert(tk.END, clean, ("bold", base_tag))
            pos = m.end()
        if pos < len(line_text):
            text_widget.insert(tk.END, line_text[pos:], (base_tag,))
        text_widget.insert(tk.END, "\n", (base_tag,))

    in_code_block = False
    lines = (markdown_str or "").splitlines()

    for raw_line in lines:
        stripped = raw_line.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue

        if in_code_block:
            text_widget.insert(tk.END, f"  {raw_line}\n", ("code_block",))
            continue

        if not stripped:
            text_widget.insert(tk.END, "\n", ("body",))
            continue

        if stripped in ("---", "***", "___"):
            text_widget.insert(tk.END, "─" * 58 + "\n", ("divider",))
            continue

        if stripped.startswith("# "):
            text_widget.insert(tk.END, stripped[2:].strip() + "\n", ("h1",))
            continue
        if stripped.startswith("## "):
            text_widget.insert(tk.END, stripped[3:].strip() + "\n", ("h2",))
            continue
        if stripped.startswith("### "):
            text_widget.insert(tk.END, stripped[4:].strip() + "\n", ("h3",))
            continue

        # Markdown table row handling
        if stripped.startswith("|") and stripped.endswith("|"):
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            # Skip alignment separator rows like | :--- | :--- |
            if all(re.match(r"^:?-{2,}:?$", c) for c in cells if c):
                continue
            # Skip generic table header row if it's just Column titles
            if cells and cells[0].lower() in ("command", "ditto's look", "byte's look", "risk scenario"):
                continue
            if len(cells) >= 2:
                col1 = re.sub(r"(\*\*|`)", "", cells[0]).strip()
                col2 = " — ".join(cells[1:])
                text_widget.insert(tk.END, f" ▸ {col1} \n", ("table_key",))
                _insert_rich_spans(col2, base_tag="table_val")
                continue

        if stripped.startswith(("* ", "- ")):
            text_widget.insert(tk.END, "  •  ", ("bullet_dot",))
            _insert_rich_spans(stripped[2:].strip(), base_tag="bullet_body")
            continue

        _insert_rich_spans(stripped, base_tag="body")

    text_widget.config(state="disabled")

