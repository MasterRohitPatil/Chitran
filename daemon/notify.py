import ctypes
import subprocess
import sys


def notify_windows(title: str, message: str, chime: bool = True) -> None:
    """Plays an optional Windows chime and triggers a non-blocking balloon toast."""
    if chime and sys.platform == "win32":
        try:
            ctypes.windll.user32.MessageBeep(0x00000040)
        except Exception:
            pass

    if sys.platform != "win32":
        print(f"[Notify] {title}: {message}")
        return

    safe_title = title.replace('"', '`"').replace("'", "''")
    safe_msg = message.replace('"', '`"').replace("'", "''")

    ps_cmd = (
        f"[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); "
        f"$notify = New-Object System.Windows.Forms.NotifyIcon; "
        f"$notify.Icon = [System.Drawing.SystemIcons]::Information; "
        f"$notify.Visible = $True; "
        f"$notify.ShowBalloonTip(4000, '{safe_title}', '{safe_msg}', [System.Windows.Forms.ToolTipIcon]::Info); "
        f"Start-Sleep -Seconds 4; "
        f"$notify.Dispose()"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except Exception as e:
        print(f"[Notify] Toast error: {e}")

