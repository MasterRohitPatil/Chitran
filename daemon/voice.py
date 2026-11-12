import asyncio
import ctypes
import json
import os
import re
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Dict, Any

import edge_tts
import pyttsx3

PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATE_FILE = PROJECT_ROOT / "data" / "voice_state.json"

DEFAULT_VOICE = "en-US-AndrewNeural"
_speech_locks: Dict[int, asyncio.Lock] = {}


def _get_speech_lock() -> asyncio.Lock:
    """Return an asyncio.Lock bound cleanly to the currently running event loop."""
    loop = asyncio.get_running_loop()
    key = id(loop)
    lock = _speech_locks.get(key)
    if lock is None:
        lock = asyncio.Lock()
        _speech_locks[key] = lock
    return lock


def _load_state() -> dict:
    if not STATE_FILE.exists():
        return {"muted": False, "voice": DEFAULT_VOICE}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"muted": False, "voice": DEFAULT_VOICE}


def _save_state(state: dict) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def is_voice_muted() -> bool:
    return bool(_load_state().get("muted", False))


def set_voice_muted(muted: bool) -> bool:
    state = _load_state()
    state["muted"] = bool(muted)
    _save_state(state)
    return state["muted"]


def toggle_mute() -> bool:
    return set_voice_muted(not is_voice_muted())


def _clean_for_speech(text: str) -> str:
    """Strips Markdown syntax and long URLs for natural headphone speech."""
    cleaned = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    cleaned = re.sub(r"https?://\S+", "", cleaned)
    cleaned = re.sub(r"[`*_#>]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _play_mp3_winmm(mp3_path: str) -> None:
    """Plays an MP3 silently via native Windows winmm.dll (blocking on worker thread)."""
    if sys.platform != "win32":
        return

    alias = f"sk_{uuid.uuid4().hex[:8]}"
    mci = ctypes.windll.winmm.mciSendStringW
    err_open = mci(f'open "{mp3_path}" type mpegvideo alias {alias}', None, 0, None)
    if err_open != 0:
        raise RuntimeError(f"winmm open failed with code {err_open}")
    try:
        err_play = mci(f"play {alias} wait", None, 0, None)
        if err_play != 0:
            raise RuntimeError(f"winmm play failed with code {err_play}")
    finally:
        mci(f"close {alias}", None, 0, None)


def _speak_offline_pyttsx3(text: str) -> None:
    """Fallback offline SAPI5 speech synthesis via pyttsx3 (COM-safe on worker threads)."""
    coinit = False
    try:
        if sys.platform == "win32":
            ctypes.windll.ole32.CoInitialize(None)
            coinit = True
        engine = pyttsx3.init()
        engine.setProperty("rate", 175)
        engine.say(text)
        engine.runAndWait()
        engine.stop()
    finally:
        if coinit:
            try:
                ctypes.windll.ole32.CoUninitialize()
            except Exception:
                pass


async def speak(text: str, force: bool = False) -> Dict[str, Any]:
    """
    Speaks text through user headphones using Edge Neural TTS with offline SAPI5 fallback.
    Respects mute state unless force=True.
    """
    if not force and is_voice_muted():
        return {"status": "muted", "spoken": False}

    clean_text = _clean_for_speech(text)
    if not clean_text:
        return {"status": "skipped", "reason": "Empty text", "spoken": False}

    voice_name = _load_state().get("voice", DEFAULT_VOICE)

    async with _get_speech_lock():
        temp_mp3 = Path(tempfile.gettempdir()) / f"sidekick_tts_{uuid.uuid4().hex[:8]}.mp3"
        try:
            communicate = edge_tts.Communicate(clean_text, voice_name)
            await asyncio.wait_for(communicate.save(str(temp_mp3)), timeout=12.0)
            await asyncio.to_thread(_play_mp3_winmm, str(temp_mp3))
            return {"status": "success", "engine": "edge-tts", "voice": voice_name, "spoken": True}
        except Exception as edge_err:
            print(f"[Voice] edge-tts failed ({type(edge_err).__name__}: {edge_err!r}), falling back to offline SAPI5...")
            try:
                await asyncio.to_thread(_speak_offline_pyttsx3, clean_text)
                return {"status": "success", "engine": "pyttsx3-sapi5", "spoken": True}
            except Exception as sapi_err:
                return {"status": "error", "message": f"All TTS engines failed: {sapi_err}", "spoken": False}
        finally:
            try:
                if temp_mp3.exists():
                    os.remove(temp_mp3)
            except Exception:
                pass


def speak_background(text: str) -> None:
    """Fire-and-forget helper for daemon background nudges."""
    def _log_result(task: asyncio.Task) -> None:
        try:
            res = task.result()
            if isinstance(res, dict) and res.get("status") == "error":
                print(f"[Voice] Background speech error: {res.get('message')}")
        except Exception as e:
            print(f"[Voice] Background speech exception: {e}")

    try:
        loop = asyncio.get_running_loop()
        task = loop.create_task(speak(text))
        task.add_done_callback(_log_result)
    except RuntimeError:
        asyncio.run(speak(text))


if __name__ == "__main__":
    print("[Voice] Running self-test...")
    res = asyncio.run(speak("Chitran voice engine online. Ready for headphones."))
    print("[Voice] Result:", res)

