"""MailBahn — SQLite registry for settings and paths."""

import os
import sqlite3
import sys
import threading
from pathlib import Path

import keyring

_KEYRING_SERVICE = "mailbahn"

# ── Secure credential storage (OS keyring) ────────────────────────
def save_password(email: str, password: str) -> None:
    def worker():
        try:
            keyring.set_password(_KEYRING_SERVICE, email, password)
        except Exception:
            pass
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout=3.0)

def load_password(email: str) -> str:
    result = [""]
    def worker():
        try:
            result[0] = keyring.get_password(_KEYRING_SERVICE, email) or ""
        except Exception:
            pass
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout=3.0)
    return result[0]

def delete_password(email: str) -> None:
    def worker():
        try:
            keyring.delete_password(_KEYRING_SERVICE, email)
        except Exception:
            pass
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout=3.0)

def _app_data_dir() -> Path:
    """Return the platform-appropriate app data directory, creating it if needed.

    Windows:  %APPDATA%/MailBahn
    macOS:    ~/Library/Application Support/MailBahn
    Linux:    $XDG_DATA_HOME/MailBahn  (defaults to ~/.local/share/MailBahn)
    """
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    d = base / "MailBahn"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _db_path() -> Path:
    return _app_data_dir() / "mailbahn.db"


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_db_path()))
    conn.execute(
        "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS template_meta ("
        "  name TEXT PRIMARY KEY,"
        "  subject TEXT DEFAULT '',"
        "  attachment_dir TEXT DEFAULT '',"
        "  cc_file TEXT DEFAULT '',"
        "  last_sent_at TEXT DEFAULT '',"
        "  last_sent_count INTEGER DEFAULT 0,"
        "  last_sent_recipients TEXT DEFAULT ''"
        ")"
    )
    # Migrate: add cc_file column if upgrading from older schema
    try:
        conn.execute("ALTER TABLE template_meta ADD COLUMN cc_file TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass  # column already exists
    conn.commit()
    return conn


def get_all_settings() -> dict[str, str]:
    conn = _connect()
    rows = conn.execute("SELECT key, value FROM settings").fetchall()
    conn.close()
    return {k: v for k, v in rows}


def get_setting(key: str, default: str = "") -> str:
    conn = _connect()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key: str, value: str) -> None:
    conn = _connect()
    conn.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value)
    )
    conn.commit()
    conn.close()


def set_many(settings: dict[str, str]) -> None:
    conn = _connect()
    conn.executemany(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        settings.items(),
    )
    conn.commit()
    conn.close()


def delete_setting(key: str) -> None:
    conn = _connect()
    conn.execute("DELETE FROM settings WHERE key = ?", (key,))
    conn.commit()
    conn.close()


def default_templates_dir() -> Path:
    d = _app_data_dir() / "templates"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_templates_dir() -> Path:
    custom = get_setting("TEMPLATES_DIR")
    if custom:
        p = Path(custom)
        if p.is_dir():
            return p
    return default_templates_dir()


def get_db_location() -> str:
    return str(_db_path())


# ── Template metadata ─────────────────────────────────────────────

def get_template_meta(name: str) -> dict:
    conn = _connect()
    row = conn.execute(
        "SELECT subject, attachment_dir, cc_file, last_sent_at, last_sent_count, last_sent_recipients "
        "FROM template_meta WHERE name = ?", (name,)
    ).fetchone()
    conn.close()
    if not row:
        return {"subject": "", "attachment_dir": "", "cc_file": "",
                "last_sent_at": "", "last_sent_count": 0, "last_sent_recipients": ""}
    return {"subject": row[0], "attachment_dir": row[1], "cc_file": row[2],
            "last_sent_at": row[3], "last_sent_count": row[4], "last_sent_recipients": row[5]}


def save_template_meta(name: str, **fields) -> None:
    conn = _connect()
    existing = conn.execute(
        "SELECT name FROM template_meta WHERE name = ?", (name,)
    ).fetchone()
    if existing:
        parts = []
        vals = []
        for k, v in fields.items():
            parts.append(f"{k} = ?")
            vals.append(v)
        if parts:
            vals.append(name)
            conn.execute(f"UPDATE template_meta SET {', '.join(parts)} WHERE name = ?", vals)
    else:
        cols = {"name": name, "subject": "", "attachment_dir": "", "cc_file": "",
                "last_sent_at": "", "last_sent_count": 0, "last_sent_recipients": ""}
        cols.update(fields)
        conn.execute(
            "INSERT INTO template_meta (name, subject, attachment_dir, cc_file, "
            "last_sent_at, last_sent_count, last_sent_recipients) "
            "VALUES (:name, :subject, :attachment_dir, :cc_file, "
            ":last_sent_at, :last_sent_count, :last_sent_recipients)", cols
        )
    conn.commit()
    conn.close()


def delete_template_meta(name: str) -> None:
    conn = _connect()
    conn.execute("DELETE FROM template_meta WHERE name = ?", (name,))
    conn.commit()
    conn.close()
