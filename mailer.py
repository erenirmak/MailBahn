"""Auto-Mail — email sending engine."""

import csv
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from pathlib import Path


def decode_bytes(raw: bytes) -> str:
    """Decode bytes to str with universal encoding support.

    Handles UTF-8 (with/without BOM), UTF-16, and falls back through
    common Windows codepages. Works for Turkish, Cyrillic, Greek, etc.
    Latin-1 is the final fallback (never fails).
    """
    # UTF BOM detection
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8")

    # Try UTF-8 (covers all scripts: CJK, Arabic, Cyrillic, etc.)
    try:
        return raw.decode("utf-8")
    except (UnicodeDecodeError, ValueError):
        pass

    # Fallback: common Windows codepages
    for codec in ("cp1254", "cp1252", "cp1251", "cp1253", "cp1256"):
        try:
            return raw.decode(codec)
        except (UnicodeDecodeError, ValueError):
            continue

    # Final fallback — latin-1 accepts every byte value
    return raw.decode("latin-1")


def load_env(filepath: Path) -> dict[str, str]:
    """Parse a .env file into a dict."""
    env = {}
    if not filepath.is_file():
        return env
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        env[key.strip()] = value.strip()
    return env


def load_recipients(filepath: Path, name_col: str = "name", email_col: str = "email") -> list[dict[str, str]]:
    """Load recipients from a CSV with configurable column names."""
    if not filepath.is_file():
        return []
    raw = filepath.read_bytes()
    text = decode_bytes(raw)
    recipients = []
    reader = csv.DictReader(text.splitlines(), skipinitialspace=True)
    for row in reader:
        name = row.get(name_col, "").strip()
        email = row.get(email_col, "").strip()
        if email:
            recipients.append({"name": name, "email": email})
    return recipients


def preview_csv(filepath: Path, max_rows: int = 10) -> dict:
    """Read CSV headers and first N rows for preview. Also detect duplicate emails."""
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
        rows.append({h: row.get(h, "") for h in headers})
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
    for row in reader:
        addr = row.get(email_col, "").strip().lower()
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
    for row in reader:
        email = row.get(email_col, "").strip()
        if email:
            cc.append(email)
    return cc


def get_attachments(directory: Path | None) -> list[Path]:
    """Return all files in a directory."""
    if directory is None or not directory.is_dir():
        return []
    return [f for f in directory.iterdir() if f.is_file()]


def build_message(
    sender: str,
    recipient: str,
    subject: str,
    html_body: str,
    inline_images: dict[str, Path],
    attachments: list[Path],
    cc: list[str] | None = None,
) -> MIMEMultipart:
    """Build a MIME message with HTML body, inline images, and file attachments."""
    msg = MIMEMultipart("mixed")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = ", ".join(cc)

    related = MIMEMultipart("related")
    related.attach(MIMEText(html_body, "html", "utf-8"))

    for cid, img_path in inline_images.items():
        with open(img_path, "rb") as f:
            img_data = f.read()
        ext = img_path.suffix.lower().lstrip(".")
        if ext == "jpg":
            ext = "jpeg"
        if ext not in ("jpeg", "png", "gif", "bmp", "webp"):
            ext = "png"
        img_mime = MIMEImage(img_data, _subtype=ext)
        img_mime.add_header("Content-ID", f"<{cid}>")
        img_mime.add_header("Content-Disposition", "inline", filename=img_path.name)
        related.attach(img_mime)

    msg.attach(related)

    for filepath in attachments:
        part = MIMEBase("application", "octet-stream")
        with open(filepath, "rb") as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=filepath.name)
        msg.attach(part)

    return msg


def send_emails(
    smtp_server: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    recipients: list[dict[str, str]],
    subject: str,
    html_body_template: str,
    inline_images: dict[str, Path],
    attachments: list[Path],
    delay: float = 2.0,
    on_progress=None,
    cc: list[str] | None = None,
):
    """Send personalized HTML emails to all recipients."""
    total = len(recipients)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)

        if on_progress:
            on_progress(0, total, "Connected and logged in.")

        for i, recipient in enumerate(recipients):
            email = recipient["email"]
            name = recipient["name"]
            html_body = html_body_template.replace("{name}", name)

            try:
                msg = build_message(
                    sender_email, email, subject,
                    html_body, inline_images, attachments,
                    cc=cc,
                )
                all_recipients = [email] + (cc or [])
                server.sendmail(sender_email, all_recipients, msg.as_string())
                status = f"[OK] Sent to {name} <{email}>"
            except smtplib.SMTPException as e:
                status = f"[FAIL] {name} <{email}>: {e}"

            if on_progress:
                on_progress(i + 1, total, status)

            if i < total - 1:
                time.sleep(delay)

    if on_progress:
        on_progress(total, total, "All emails processed.")
