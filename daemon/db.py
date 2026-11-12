import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
DB_PATH = DATA_DIR / "sidekick.db"


@contextmanager
def get_connection(db_path: Optional[Path] = None) -> Iterator[sqlite3.Connection]:
    """Yields a SQLite connection with WAL mode, auto-commit/rollback, and guaranteed close."""
    target = db_path or DB_PATH
    target.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(target), timeout=10.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(db_path: Optional[Path] = None) -> Path:
    """Initializes all Sidekick OS tables and indexes if they do not exist."""
    target = db_path or DB_PATH
    with get_connection(target) as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS goals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                title TEXT NOT NULL,
                deadline_time TEXT,
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS activity_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                app_name TEXT NOT NULL,
                window_title TEXT NOT NULL,
                category TEXT NOT NULL,
                duration_sec INTEGER NOT NULL DEFAULT 15
            );

            CREATE TABLE IF NOT EXISTS quota_ledger (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                provider TEXT NOT NULL,
                model TEXT NOT NULL,
                prompt_tokens INTEGER NOT NULL DEFAULT 0,
                completion_tokens INTEGER NOT NULL DEFAULT 0,
                total_tokens INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS habit_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                event_type TEXT NOT NULL,
                action_taken TEXT NOT NULL DEFAULT 'triggered'
            );

            CREATE INDEX IF NOT EXISTS idx_goals_date ON goals(date);
            CREATE INDEX IF NOT EXISTS idx_activity_date ON activity_logs(date);
            CREATE INDEX IF NOT EXISTS idx_quota_date_provider ON quota_ledger(date, provider);
            CREATE INDEX IF NOT EXISTS idx_habit_date ON habit_events(date);
        """)
    return target


if __name__ == "__main__":
    created = init_db()
    print(f"[DB] Initialized SQLite database at: {created}")
