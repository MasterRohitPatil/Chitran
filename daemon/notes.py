import os
from datetime import datetime
from pathlib import Path


def append_daily_note(note_text: str) -> dict:
    """Appends a timestamped Markdown note to ~/Daily_Notes/YYYY-MM-DD.md"""
    note_text = note_text.strip()
    if not note_text:
        return {"status": "error", "message": "Note text cannot be empty"}

    home_dir = Path.home()
    notes_dir = home_dir / "Daily_Notes"
    notes_dir.mkdir(parents=True, exist_ok=True)

    today_str = datetime.now().strftime("%Y-%m-%d")
    note_file = notes_dir / f"{today_str}.md"

    current_time_str = datetime.now().strftime("%H:%M")
    note_line = f"- [{current_time_str}] {note_text}\n"

    # If new file for the day, add a top-level date header
    is_new = not note_file.exists() or note_file.stat().st_size == 0
    with open(note_file, "a", encoding="utf-8") as f:
        if is_new:
            f.write(f"# Daily Log — {today_str}\n\n")
        f.write(note_line)

    return {
        "status": "success",
        "file": str(note_file),
        "entry": note_line.strip()
    }

