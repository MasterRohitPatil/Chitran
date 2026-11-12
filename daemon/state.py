import json
from pathlib import Path
from typing import Dict, Any

PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATE_FILE = PROJECT_ROOT / "data" / "system_state.json"

VALID_STATES = ("focused", "working", "alert", "paused", "hydration", "mail")


def load_system_state() -> Dict[str, Any]:
    """Loads shared Sidekick OS runtime state (paused, mascot_state, badge_visible)."""
    default = {
        "paused": False,
        "mascot_state": "focused",
        "badge_visible": True,
        "last_goal_prompt_date": "",
    }
    if not STATE_FILE.exists():
        return default
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            default.update(data)
            return default
    except Exception:
        return default


def save_system_state(state: Dict[str, Any]) -> Dict[str, Any]:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
    return state


def is_paused() -> bool:
    return bool(load_system_state().get("paused", False))


def set_paused(paused: bool) -> Dict[str, Any]:
    state = load_system_state()
    state["paused"] = bool(paused)
    state["mascot_state"] = "paused" if paused else "focused"
    return save_system_state(state)


def toggle_paused() -> Dict[str, Any]:
    return set_paused(not is_paused())


def set_mascot_state(new_state: str) -> Dict[str, Any]:
    state = load_system_state()
    if new_state in VALID_STATES:
        state["mascot_state"] = new_state
        state["paused"] = (new_state == "paused")
        save_system_state(state)
    return state

