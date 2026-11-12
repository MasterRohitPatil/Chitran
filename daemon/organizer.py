"""
Sidekick OS — Safe Downloads Auto-Organizer (Step 5.2)

Maps top-level files in Downloads into organized Documents subfolders using
config/organizer_rules.json, gated behind a Dark Cyber-HUD preview window.

Strict safety guarantee: Sidekick OS enforces zero file deletion here.
Every move resolves destination name conflicts safely and logs reversible
batches into data/undo_log.json for `sk undo`.
"""

import argparse
from datetime import datetime
import json
from pathlib import Path
import shutil
import sys
from typing import Any, Dict, List, Optional

if sys.stdout is None:
    sys.stdout = open("nul", "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open("nul", "w", encoding="utf-8")

DAEMON_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = DAEMON_DIR.parent
if str(DAEMON_DIR) not in sys.path:
    sys.path.insert(0, str(DAEMON_DIR))

from notify import notify_windows
from win_ui import apply_win11_dwm_style, enable_high_dpi

RULES_FILE = PROJECT_ROOT / "config" / "organizer_rules.json"
UNDO_LOG_FILE = PROJECT_ROOT / "data" / "undo_log.json"

DEFAULT_IGNORE_EXTENSIONS = {
    ".tmp",
    ".crdownload",
    ".part",
    ".partial",
    ".download",
    ".ini",
    ".lnk",
    ".exe",
    ".msi",
    ".bat",
}
SKIP_EXACT_NAMES = {"desktop.ini", "thumbs.db"}

# Dark Cyber-HUD palette matching daemon/hud.py
BG_ROOT = "#121214"
BG_CARD = "#18181B"
BG_ROW = "#1E1E24"
BORDER_SUBTLE = "#27272A"
TEXT_PRIMARY = "#E4E4E7"
TEXT_MUTED = "#A1A1AA"
ACCENT_CYAN = "#22D3EE"
ACCENT_EMERALD = "#10B981"
ACCENT_AMBER = "#F59E0B"
BTN_CANCEL_BG = "#27272A"


def load_organizer_rules(rules_path: Optional[Path] = None) -> Dict[str, Any]:
    """Load rules configuration from JSON."""
    path = Path(rules_path) if rules_path else RULES_FILE
    if not path.exists():
        raise FileNotFoundError(f"Organizer rules file not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_unique_destination(
    dest: Path,
    reserved: Optional[set] = None,
) -> Path:
    """Return a non-colliding destination path without overwriting existing files."""
    reserved_norm = reserved if reserved is not None else set()
    candidate = dest
    counter = 1
    while candidate.exists() or str(candidate).lower() in reserved_norm:
        candidate = dest.with_name(f"{dest.stem} ({counter}){dest.suffix}")
        counter += 1
    return candidate


def _match_rule(file_path: Path, rules: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Return the first matching rule for a file, or None if unmatched."""
    ext = file_path.suffix.lower()
    lower_name = file_path.name.lower()

    for rule in rules:
        rule_exts = {e.lower() for e in rule.get("extensions", [])}
        if ext and ext in rule_exts:
            return rule

        allowed_kw_exts = {
            e.lower() for e in rule.get("allowed_extensions_for_keywords", [])
        }
        if allowed_kw_exts and ext not in allowed_kw_exts:
            continue

        keywords = [kw.lower() for kw in rule.get("keywords", []) if kw]
        if any(kw in lower_name for kw in keywords):
            return rule

    return None


def plan_downloads_organization(
    rules_path: Optional[Path] = None,
    source_dir: Optional[Path] = None,
    docs_root: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Scan source_dir (non-recursive) and build a non-destructive move plan.
    Returns preview metadata without touching any files on disk.
    """
    config = load_organizer_rules(rules_path)
    src = Path(source_dir) if source_dir else Path(config["source_dir"])
    docs = Path(docs_root) if docs_root else Path(config["documents_root"])

    if not src.exists():
        return {
            "status": "error",
            "message": f"Source folder does not exist: {src}",
            "source_dir": str(src),
            "documents_root": str(docs),
            "move_count": 0,
            "skipped_count": 0,
            "moves": [],
        }

    ignore_exts = {
        e.lower() for e in config.get("ignore_extensions", DEFAULT_IGNORE_EXTENSIONS)
    }
    rules = config.get("rules", [])

    moves: List[Dict[str, Any]] = []
    skipped_count = 0
    reserved: set = set()

    entries = sorted(src.iterdir(), key=lambda p: p.name.lower())
    for item in entries:
        if not item.is_file():
            continue

        lower_name = item.name.lower()
        ext = item.suffix.lower()

        if item.name.startswith(".") or lower_name in SKIP_EXACT_NAMES or ext in ignore_exts:
            skipped_count += 1
            continue

        matched = _match_rule(item, rules)
        if not matched:
            skipped_count += 1
            continue

        folder_name = matched["folder"]
        target_dir = docs / folder_name
        raw_dest = target_dir / item.name

        if item.resolve() == raw_dest.resolve():
            skipped_count += 1
            continue

        unique_dest = resolve_unique_destination(raw_dest, reserved=reserved)
        reserved.add(str(unique_dest).lower())

        moves.append({
            "filename": item.name,
            "src": str(item),
            "dst_dir": str(target_dir),
            "dst": str(unique_dest),
            "folder": folder_name,
            "rule": matched.get("name", folder_name),
            "renamed": unique_dest.name != item.name,
        })

    return {
        "status": "preview",
        "source_dir": str(src),
        "documents_root": str(docs),
        "move_count": len(moves),
        "skipped_count": skipped_count,
        "moves": moves,
    }


def _read_undo_log(log_path: Path) -> List[Dict[str, Any]]:
    if not log_path.exists():
        return []
    try:
        data = json.loads(log_path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _write_undo_log(log_path: Path, batches: List[Dict[str, Any]]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(batches, indent=2), encoding="utf-8")


def apply_organization_plan(
    moves: List[Dict[str, Any]],
    undo_log_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Execute confirmed file moves safely and record a reversible batch in undo_log.json.
    Never overwrites existing files and performs zero deletions.
    """
    if not moves:
        return {
            "status": "cancelled",
            "moved_count": 0,
            "errors": [],
            "message": "No files selected to move.",
        }

    log_path = Path(undo_log_path) if undo_log_path else UNDO_LOG_FILE
    applied: List[Dict[str, str]] = []
    errors: List[str] = []

    for item in moves:
        src_path = Path(item["src"])
        dst_dir = Path(item["dst_dir"])
        desired_dst = Path(item["dst"])

        if not src_path.exists():
            errors.append(f"Missing source file: {src_path.name}")
            continue

        try:
            dst_dir.mkdir(parents=True, exist_ok=True)
            final_dst = resolve_unique_destination(desired_dst)
            shutil.move(str(src_path), str(final_dst))
            applied.append({
                "filename": src_path.name,
                "folder": item.get("folder", dst_dir.name),
                "src": str(src_path),
                "dst": str(final_dst),
            })
        except Exception as exc:
            errors.append(f"{src_path.name}: {exc}")

    if applied:
        batches = _read_undo_log(log_path)
        batches.append({
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "count": len(applied),
            "moves": applied,
        })
        _write_undo_log(log_path, batches)

    status = "success" if applied else "error"
    msg = (
        f"Safely moved {len(applied)} file(s) into Documents."
        if applied
        else "No files could be moved."
    )
    if errors:
        msg += f" ({len(errors)} warning/error)"

    return {
        "status": status,
        "moved_count": len(applied),
        "applied": applied,
        "errors": errors,
        "message": msg,
    }


def undo_last_organize(undo_log_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Reverse the most recent organize batch from undo_log.json without deleting files.
    """
    log_path = Path(undo_log_path) if undo_log_path else UNDO_LOG_FILE
    batches = _read_undo_log(log_path)
    if not batches:
        return {
            "status": "empty",
            "restored_count": 0,
            "missing_count": 0,
            "errors": [],
            "message": "No organize batches found to undo.",
        }

    last_batch = batches.pop()
    restored: List[Dict[str, str]] = []
    missing_count = 0
    errors: List[str] = []

    for item in reversed(last_batch.get("moves", [])):
        dst_path = Path(item["dst"])
        orig_src = Path(item["src"])

        if not dst_path.exists():
            missing_count += 1
            errors.append(f"Not found at destination: {dst_path.name}")
            continue

        try:
            orig_src.parent.mkdir(parents=True, exist_ok=True)
            final_src = resolve_unique_destination(orig_src)
            shutil.move(str(dst_path), str(final_src))
            restored.append({
                "from": str(dst_path),
                "to": str(final_src),
            })
        except Exception as exc:
            errors.append(f"{dst_path.name}: {exc}")

    _write_undo_log(log_path, batches)

    msg = f"Restored {len(restored)} file(s) back to Downloads."
    if missing_count:
        msg += f" ({missing_count} missing)"

    return {
        "status": "success" if restored or missing_count == 0 else "partial",
        "restored_count": len(restored),
        "missing_count": missing_count,
        "restored": restored,
        "errors": errors,
        "message": msg,
    }


def show_organizer_preview_hud(
    rules_path: Optional[Path] = None,
    undo_log_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Render the Dark Cyber-HUD confirmation gate before moving any Downloads file.
    Returns the final action summary dict once the HUD closes.
    """
    import tkinter as tk

    enable_high_dpi()
    plan = plan_downloads_organization(rules_path=rules_path)
    outcome: Dict[str, Any] = {"status": "cancelled", "moved_count": 0}

    root = tk.Tk()
    root.title("Chitran — Safe Downloads Organizer")
    root.configure(bg=BG_ROOT)
    root.attributes("-topmost", True)
    root.resizable(True, True)

    width = 680
    height = 540 if plan["move_count"] > 0 else 250
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    pos_x = max(20, (screen_w - width) // 2)
    pos_y = max(20, int((screen_h - height) * 0.34))
    root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
    root.minsize(560, 230)

    apply_win11_dwm_style(root)

    outer = tk.Frame(root, bg=BG_ROOT, padx=20, pady=18)
    outer.pack(fill="both", expand=True)

    # Header row
    header = tk.Frame(outer, bg=BG_ROOT)
    header.pack(fill="x", pady=(0, 10))

    tk.Label(
        header,
        text="SAFE DOWNLOADS ORGANIZER",
        font=("Segoe UI", 10, "bold"),
        fg=ACCENT_CYAN,
        bg=BG_ROOT,
    ).pack(side="left")

    badge_text = (
        f"{plan['move_count']} files ready  •  Zero deletion"
        if plan["move_count"] > 0
        else "Downloads clean  •  Zero deletion"
    )
    tk.Label(
        header,
        text=badge_text,
        font=("Segoe UI", 9, "bold"),
        fg=ACCENT_EMERALD if plan["move_count"] > 0 else TEXT_MUTED,
        bg=BG_CARD,
        padx=10,
        pady=3,
    ).pack(side="right")

    if plan.get("status") == "error":
        tk.Label(
            outer,
            text=plan.get("message", "Unable to scan Downloads folder."),
            font=("Segoe UI", 10),
            fg=ACCENT_AMBER,
            bg=BG_ROOT,
            wraplength=width - 60,
            justify="left",
        ).pack(anchor="w", pady=(10, 16))
        tk.Button(
            outer,
            text="Close [Esc]",
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_PRIMARY,
            bg=BTN_CANCEL_BG,
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            command=root.destroy,
        ).pack(anchor="e")
        root.bind("<Escape>", lambda _e: root.destroy())
        root.mainloop()
        return plan

    if plan["move_count"] == 0:
        card = tk.Frame(
            outer,
            bg=BG_CARD,
            highlightbackground=BORDER_SUBTLE,
            highlightthickness=1,
            padx=18,
            pady=18,
        )
        card.pack(fill="both", expand=True, pady=(4, 14))
        tk.Label(
            card,
            text="Nothing new to sort right now",
            font=("Segoe UI", 12, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD,
        ).pack(anchor="w")
        tk.Label(
            card,
            text=(
                f"Checked {plan['source_dir']} and left {plan['skipped_count']} unmatched "
                "or protected file(s) untouched (.exe, active downloads, personal docs)."
            ),
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CARD,
            wraplength=width - 90,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        btn_row = tk.Frame(outer, bg=BG_ROOT)
        btn_row.pack(fill="x")
        tk.Button(
            btn_row,
            text="OK [Esc]",
            font=("Segoe UI", 9, "bold"),
            fg=BG_ROOT,
            bg=ACCENT_CYAN,
            relief="flat",
            padx=16,
            pady=7,
            cursor="hand2",
            command=root.destroy,
        ).pack(side="right")

        root.bind("<Escape>", lambda _e: root.destroy())
        root.bind("<Return>", lambda _e: root.destroy())
        root.mainloop()
        return outcome

    # Toolbar row: subtitle + quick selection toggles
    toolbar = tk.Frame(outer, bg=BG_ROOT)
    toolbar.pack(fill="x", pady=(0, 8))

    tk.Label(
        toolbar,
        text="Review matched moves before anything changes on disk:",
        font=("Segoe UI", 9),
        fg=TEXT_MUTED,
        bg=BG_ROOT,
    ).pack(side="left")

    row_vars: List[tk.BooleanVar] = []
    confirm_btn_ref: Dict[str, Any] = {}

    def _update_confirm_label(*_args: Any) -> None:
        selected = sum(1 for v in row_vars if v.get())
        btn = confirm_btn_ref.get("btn")
        if btn:
            btn.config(
                text=f"Confirm & Move ({selected})",
                state="normal" if selected > 0 else "disabled",
                bg=ACCENT_CYAN if selected > 0 else BTN_CANCEL_BG,
                fg=BG_ROOT if selected > 0 else TEXT_MUTED,
            )

    def _set_all(val: bool) -> None:
        for var in row_vars:
            var.set(val)
        _update_confirm_label()

    tk.Button(
        toolbar,
        text="Clear",
        font=("Segoe UI", 8, "bold"),
        fg=TEXT_MUTED,
        bg=BG_CARD,
        activebackground=BORDER_SUBTLE,
        activeforeground=TEXT_PRIMARY,
        relief="flat",
        padx=8,
        pady=2,
        cursor="hand2",
        command=lambda: _set_all(False),
    ).pack(side="right", padx=(6, 0))

    tk.Button(
        toolbar,
        text="Select All",
        font=("Segoe UI", 8, "bold"),
        fg=ACCENT_CYAN,
        bg=BG_CARD,
        activebackground=BORDER_SUBTLE,
        activeforeground=ACCENT_CYAN,
        relief="flat",
        padx=8,
        pady=2,
        cursor="hand2",
        command=lambda: _set_all(True),
    ).pack(side="right")

    # Scrollable preview list container
    list_card = tk.Frame(
        outer,
        bg=BG_CARD,
        highlightbackground=BORDER_SUBTLE,
        highlightthickness=1,
    )
    list_card.pack(fill="both", expand=True)

    canvas = tk.Canvas(list_card, bg=BG_CARD, highlightthickness=0, bd=0)
    scrollbar = tk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
    scroll_inner = tk.Frame(canvas, bg=BG_CARD, padx=10, pady=8)

    scroll_inner.bind(
        "<Configure>",
        lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas_win = canvas.create_window((0, 0), window=scroll_inner, anchor="nw")
    canvas.bind(
        "<Configure>",
        lambda e: canvas.itemconfigure(canvas_win, width=e.width),
    )
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event: Any) -> None:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    for idx, move in enumerate(plan["moves"]):
        var = tk.BooleanVar(value=True)
        var.trace_add("write", _update_confirm_label)
        row_vars.append(var)

        row_bg = BG_ROW if idx % 2 == 0 else BG_CARD
        row = tk.Frame(scroll_inner, bg=row_bg, padx=8, pady=6)
        row.pack(fill="x", pady=2)

        cb = tk.Checkbutton(
            row,
            variable=var,
            bg=row_bg,
            activebackground=row_bg,
            selectcolor=BG_ROOT,
            fg=ACCENT_CYAN,
            relief="flat",
            highlightthickness=0,
            cursor="hand2",
        )
        cb.pack(side="left")

        disp_name = move["filename"]
        if len(disp_name) > 40:
            disp_name = disp_name[:37] + "..."

        tk.Label(
            row,
            text=disp_name,
            font=("Segoe UI", 9),
            fg=TEXT_PRIMARY,
            bg=row_bg,
            anchor="w",
        ).pack(side="left", padx=(4, 8))

        dest_label = f"-> {move['folder']}/"
        if move.get("renamed"):
            dest_label += f" ({Path(move['dst']).name})"

        tk.Label(
            row,
            text=dest_label,
            font=("Consolas", 9, "bold"),
            fg=ACCENT_AMBER if move.get("renamed") else ACCENT_CYAN,
            bg=row_bg,
            anchor="e",
        ).pack(side="right")

    # Safety footer + action bar
    footer = tk.Frame(outer, bg=BG_ROOT)
    footer.pack(fill="x", pady=(12, 0))

    tk.Label(
        footer,
        text=(
            f"Skipping {plan['skipped_count']} unmatched/protected item(s). "
            "Reversible anytime via 'cn undo'."
        ),
        font=("Segoe UI", 8),
        fg=TEXT_MUTED,
        bg=BG_ROOT,
    ).pack(side="left")

    def _show_completion_view(result: Dict[str, Any]) -> None:
        nonlocal outcome
        outcome = result
        canvas.unbind_all("<MouseWheel>")
        for child in outer.winfo_children():
            child.destroy()

        hdr = tk.Frame(outer, bg=BG_ROOT)
        hdr.pack(fill="x", pady=(0, 10))
        tk.Label(
            hdr,
            text="SAFE DOWNLOADS ORGANIZER — COMPLETE",
            font=("Segoe UI", 10, "bold"),
            fg=ACCENT_EMERALD,
            bg=BG_ROOT,
        ).pack(side="left")

        done_card = tk.Frame(
            outer,
            bg=BG_CARD,
            highlightbackground=BORDER_SUBTLE,
            highlightthickness=1,
            padx=18,
            pady=18,
        )
        done_card.pack(fill="both", expand=True, pady=(4, 14))

        tk.Label(
            done_card,
            text=result.get("message", "Organization complete."),
            font=("Segoe UI", 11, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD,
        ).pack(anchor="w")

        tk.Label(
            done_card,
            text=(
                "Batch saved to data/undo_log.json. Run 'cn undo' in Flow Launcher "
                "or click Undo Now to move everything right back."
            ),
            font=("Segoe UI", 9),
            fg=TEXT_MUTED,
            bg=BG_CARD,
            wraplength=width - 90,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

        if result.get("errors"):
            tk.Label(
                done_card,
                text="\n".join(result["errors"][:5]),
                font=("Consolas", 8),
                fg=ACCENT_AMBER,
                bg=BG_CARD,
                justify="left",
            ).pack(anchor="w", pady=(8, 0))

        actions = tk.Frame(outer, bg=BG_ROOT)
        actions.pack(fill="x")

        def _undo_now() -> None:
            nonlocal outcome
            undo_res = undo_last_organize(undo_log_path=undo_log_path)
            notify_windows("Chitran — Undo Organize", undo_res["message"])
            outcome = undo_res
            root.destroy()

        tk.Button(
            actions,
            text="Undo Now",
            font=("Segoe UI", 9, "bold"),
            fg=ACCENT_AMBER,
            bg=BTN_CANCEL_BG,
            relief="flat",
            padx=12,
            pady=7,
            cursor="hand2",
            command=_undo_now,
        ).pack(side="left")

        tk.Button(
            actions,
            text="Done [Esc]",
            font=("Segoe UI", 9, "bold"),
            fg=BG_ROOT,
            bg=ACCENT_EMERALD,
            relief="flat",
            padx=16,
            pady=7,
            cursor="hand2",
            command=root.destroy,
        ).pack(side="right")

        root.bind("<Return>", lambda _e: root.destroy())

    def _on_confirm() -> None:
        selected_moves = [
            move for move, var in zip(plan["moves"], row_vars) if var.get()
        ]
        result = apply_organization_plan(selected_moves, undo_log_path=undo_log_path)
        notify_windows("Chitran — Downloads Organized", result["message"])
        _show_completion_view(result)

    def _on_cancel() -> None:
        nonlocal outcome
        outcome = {"status": "cancelled", "moved_count": 0}
        root.destroy()

    confirm_btn = tk.Button(
        footer,
        text=f"Confirm & Move ({len(plan['moves'])})",
        font=("Segoe UI", 9, "bold"),
        fg=BG_ROOT,
        bg=ACCENT_CYAN,
        activebackground="#06B6D4",
        relief="flat",
        padx=14,
        pady=7,
        cursor="hand2",
        command=_on_confirm,
    )
    confirm_btn.pack(side="right")
    confirm_btn_ref["btn"] = confirm_btn

    tk.Button(
        footer,
        text="Cancel [Esc]",
        font=("Segoe UI", 9),
        fg=TEXT_PRIMARY,
        bg=BTN_CANCEL_BG,
        activebackground=BORDER_SUBTLE,
        activeforeground=TEXT_PRIMARY,
        relief="flat",
        padx=12,
        pady=7,
        cursor="hand2",
        command=_on_cancel,
    ).pack(side="right", padx=(0, 8))

    root.bind("<Escape>", lambda _e: _on_cancel())
    root.mainloop()
    return outcome


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sidekick OS Safe Downloads Organizer")
    parser.add_argument("--preview", action="store_true", help="Open Cyber-HUD confirmation preview")
    parser.add_argument("--undo", action="store_true", help="Undo the most recent organize batch")
    parser.add_argument("--plan-json", action="store_true", help="Print dry-run move plan as JSON")
    args = parser.parse_args()

    if args.undo:
        print(json.dumps(undo_last_organize(), indent=2))
    elif args.plan_json:
        print(json.dumps(plan_downloads_organization(), indent=2))
    else:
        show_organizer_preview_hud()
