"""CSV helpers for recipient and CC file workflows."""

import csv
from pathlib import Path

from mailer import decode_bytes


def load_recipients(filepath: Path, name_col: str = "name", email_col: str = "email") -> list[dict[str, str]]:
    """Load recipients from a CSV with configurable column names."""
    if not filepath.is_file():
        return []
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    recipients = []
    reader = csv.DictReader(text.splitlines(), skipinitialspace=True)
    for i, row in enumerate(reader, start=2):
        name_val = row.get(name_col)
        email_val = row.get(email_col)
        if name_val is None or email_val is None:
            raise ValueError(f"Malformed data at row {i}. A column is missing. Please check your CSV.")
        name = name_val.strip()
        email = email_val.strip()
        if email:
            recipients.append({"name": name, "email": email})
    return recipients


def preview_csv(filepath: Path, max_rows: int = 10) -> dict:
    """Read CSV headers and first N rows for preview."""
    if not filepath.is_file():
        return {"ok": False, "error": "File not found."}
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    lines = text.splitlines()
    if not lines:
        return {"ok": False, "error": "File is empty."}
    reader = csv.DictReader(lines, skipinitialspace=True)
    headers = reader.fieldnames or []
    if not headers:
        return {"ok": False, "error": "No columns found."}
    rows = []
    for i, row in enumerate(reader):
        if i >= max_rows:
            break
        rows.append({h: (row.get(h) or "") for h in headers})
    total = sum(1 for _ in csv.DictReader(lines, skipinitialspace=True))
    return {"ok": True, "headers": headers, "rows": rows, "total": total}


def detect_duplicates(filepath: Path, email_col: str = "email") -> list[str]:
    """Return list of duplicate email addresses in the CSV."""
    if not filepath.is_file():
        return []
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    reader = csv.DictReader(text.splitlines(), skipinitialspace=True)
    seen = {}
    for i, row in enumerate(reader, start=2):
        email_val = row.get(email_col)
        if email_val is None:
            raise ValueError(f"Malformed data at row {i}. A column is missing. Please check your CSV.")
        addr = email_val.strip().lower()
        if addr:
            seen[addr] = seen.get(addr, 0) + 1
    return [addr for addr, count in seen.items() if count > 1]


def load_cc(filepath: Path, email_col: str = "email") -> list[str]:
    """Load CC addresses from a CSV with a configurable email column."""
    if not filepath.is_file():
        return []
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    cc = []
    reader = csv.DictReader(text.splitlines(), skipinitialspace=True)
    for i, row in enumerate(reader, start=2):
        email_val = row.get(email_col)
        if email_val is None:
            raise ValueError(f"Malformed data at row {i}. A column is missing. Please check your CSV.")
        email = email_val.strip()
        if email:
            cc.append(email)
    return cc


def format_preview(items: list[str], max_items: int = 5) -> str:
    """Format a short preview string for loaded CSV content."""
    preview = ", ".join(items[:max_items])
    if len(items) > max_items:
        preview += f" +{len(items) - max_items} more"
    return preview


def confirm_recipient_csv(path: str, name_col: str, email_col: str) -> tuple[list[dict[str, str]], dict]:
    """Load recipients with selected columns and return API-ready metadata."""
    csv_path = Path(path)
    try:
        duplicates = detect_duplicates(csv_path, email_col)
        recipients = load_recipients(csv_path, name_col, email_col)
    except ValueError as e:
        return [], {"ok": False, "error": str(e)}
    except Exception as e:
        return [], {"ok": False, "error": f"Error parsing CSV: {e}"}
    if not recipients:
        return [], {"ok": False, "error": "No recipients found with those columns."}

    names = [row["name"] or row["email"] for row in recipients]
    result = {"ok": True, "count": len(recipients), "preview": format_preview(names)}
    if duplicates:
        result["duplicates"] = duplicates
    return recipients, result


def load_recipient_csv(path: str) -> tuple[list[dict[str, str]], dict]:
    """Load recipients with default columns and return API-ready metadata."""
    try:
        recipients = load_recipients(Path(path))
    except ValueError as e:
        return [], {"ok": False, "error": str(e)}
    except Exception as e:
        return [], {"ok": False, "error": f"Error parsing CSV: {e}"}
    if not recipients:
        return [], {"ok": False, "error": "No recipients found."}

    names = [row["name"] or row["email"] for row in recipients]
    return recipients, {"ok": True, "count": len(recipients), "preview": format_preview(names)}


def confirm_cc_csv(path: str, email_col: str) -> tuple[list[str], dict]:
    """Load CC addresses with the selected column and return API-ready metadata."""
    csv_path = Path(path)
    try:
        duplicates = detect_duplicates(csv_path, email_col)
        cc_list = load_cc(csv_path, email_col)
    except ValueError as e:
        return [], {"ok": False, "error": str(e)}
    except Exception as e:
        return [], {"ok": False, "error": f"Error parsing CSV: {e}"}
    if not cc_list:
        return [], {"ok": False, "error": "No CC addresses found with that column."}

    result = {"ok": True, "count": len(cc_list), "preview": format_preview(cc_list)}
    if duplicates:
        result["duplicates"] = duplicates
    return cc_list, result


def load_cc_csv(path: str) -> tuple[list[str], dict]:
    """Load CC addresses with the default email column and return API-ready metadata."""
    try:
        cc_list = load_cc(Path(path))
    except ValueError as e:
        return [], {"ok": False, "error": str(e)}
    except Exception as e:
        return [], {"ok": False, "error": f"Error parsing CSV: {e}"}
    if not cc_list:
        return [], {"ok": False, "error": "No CC addresses found."}
    return cc_list, {"ok": True, "count": len(cc_list), "preview": format_preview(cc_list)}
