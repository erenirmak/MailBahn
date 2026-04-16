"""Auto-Mail — SQLite registry for settings and paths."""

import os
import sqlite3
import sys
from pathlib import Path

import keyring

_KEYRING_SERVICE = "auto-mail"


def _app_data_dir() -> Path:
    """Return the platform-appropriate app data directory, creating it if needed.

    Windows:  %APPDATA%/auto-mail
    macOS:    ~/Library/Application Support/auto-mail
    Linux:    $XDG_DATA_HOME/auto-mail  (defaults to ~/.local/share/auto-mail)
    """
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    d = base / "auto-mail"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _db_path() -> Path:
    return _app_data_dir() / "automail.db"


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_db_path()))
    conn.execute(
        "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)"
    )
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


# ── Secure credential storage (OS keyring) ────────────────────────

def save_password(email: str, password: str) -> None:
    keyring.set_password(_KEYRING_SERVICE, email, password)


def load_password(email: str) -> str:
    return keyring.get_password(_KEYRING_SERVICE, email) or ""


def delete_password(email: str) -> None:
    try:
        keyring.delete_password(_KEYRING_SERVICE, email)
    except keyring.errors.PasswordDeleteError:
        pass
