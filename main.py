import os
os.environ.setdefault("PYTHONUTF8", "1")

import base64
import html as html_mod
import json
import logging
import sys
import threading
import uuid
from pathlib import Path

import webview

from csvutils import confirm_cc_csv, confirm_recipient_csv, load_cc_csv, load_recipient_csv, preview_csv
from mailer import decode_bytes, get_attachments, send_emails
import db


class Api:
    """Python backend exposed to the browser UI via pywebview."""

    def __init__(self):
        self.env_config: dict[str, str] = {}
        self.recipients: list[dict[str, str]] = []
        self.cc_list: list[str] = []
        self.image_paths: dict[str, Path] = {}
        self._last_csv_path: str = ""
        self._last_cc_path: str = ""
        self.window = None

    # ── File Dialogs ──────────────────────────────────────────────

    def browse_csv(self):
        r = self.window.create_file_dialog(
            webview.FileDialog.OPEN, file_types=("CSV Files (*.csv)", "All Files (*.*)")
        )
        return r[0] if r else None

    def browse_folder(self):
        r = self.window.create_file_dialog(webview.FileDialog.FOLDER)
        return r[0] if r else None

    def browse_cc_csv(self):
        r = self.window.create_file_dialog(
            webview.FileDialog.OPEN, file_types=("CSV Files (*.csv)", "All Files (*.*)")
        )
        return r[0] if r else None

    # ── Settings (SQLite registry + OS keyring) ────────────────────

    def load_settings(self):
        s = db.get_all_settings()
        if not s.get("SMTP_SERVER"):
            return {"ok": False}
        email = s.get("SENDER_EMAIL", "")
        if email:
            s["SENDER_PASSWORD"] = db.load_password(email)
        self.env_config = s
        return {"ok": True, **s}

    def save_settings(self, settings):
        try:
            password = settings.pop("SENDER_PASSWORD", "")
            email = settings.get("SENDER_EMAIL", "")
            if email and password:
                db.save_password(email, password)
            db.set_many(settings)
            self.env_config = {**db.get_all_settings()}
            if email:
                self.env_config["SENDER_PASSWORD"] = db.load_password(email)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def get_db_location(self):
        return db.get_db_location()

    def test_connection(self):
        """Try SMTP login without sending anything."""
        import smtplib
        srv = self.env_config.get("SMTP_SERVER", "")
        port_s = self.env_config.get("SMTP_PORT", "587")
        email = self.env_config.get("SENDER_EMAIL", "")
        pw = self.env_config.get("SENDER_PASSWORD", "")
        if not srv or not email or not pw:
            return {"ok": False, "error": "Save your SMTP settings first."}
        try:
            port = int(port_s)
        except ValueError:
            return {"ok": False, "error": f"Invalid port: {port_s}"}
        try:
            with smtplib.SMTP(srv, port, timeout=10) as server:
                server.starttls()
                server.login(email, pw)
            return {"ok": True, "message": f"Connected to {srv}:{port} as {email}"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── Loaders ───────────────────────────────────────────────────

    def preview_csv_file(self, path):
        return preview_csv(Path(path))

    def confirm_csv(self, path, name_col, email_col):
        """Load recipients with the chosen column mapping, detect duplicates."""
        recipients, result = confirm_recipient_csv(path, name_col, email_col)
        if not result.get("ok"):
            return result
        self.recipients = recipients
        self._last_csv_path = path
        return result

    def load_csv_file(self, path):
        recipients, result = load_recipient_csv(path)
        if not result.get("ok"):
            return result
        self.recipients = recipients
        self._last_csv_path = path
        return result

    def load_cc_file(self, path):
        cc_list, result = load_cc_csv(path)
        if not result.get("ok"):
            return result
        self.cc_list = cc_list
        self._last_cc_path = path
        return result

    def preview_cc_file(self, path):
        return preview_csv(Path(path))

    def confirm_cc(self, path, email_col):
        """Load CC with the chosen email column, detect duplicates."""
        cc_list, result = confirm_cc_csv(path, email_col)
        if not result.get("ok"):
            return result
        self.cc_list = cc_list
        self._last_cc_path = path
        return result

    # ── Message Editing ───────────────────────────────────────────

    def pick_image(self):
        r = self.window.create_file_dialog(
            webview.FileDialog.OPEN,
            file_types=(
                "PNG images (*.png)",
                "JPEG images (*.jpg;*.jpeg)",
                "GIF images (*.gif)",
                "BMP images (*.bmp)",
                "WebP images (*.webp)",
            ),
        )
        if not r:
            return None
        img_path = Path(r[0])
        cid = f"img_{uuid.uuid4().hex[:8]}"
        self.image_paths[cid] = img_path
        raw = img_path.read_bytes()
        b64 = base64.b64encode(raw).decode("ascii")
        ext = img_path.suffix.lower().lstrip(".")
        if ext == "jpg":
            ext = "jpeg"
        if ext not in ("jpeg", "png", "gif", "bmp", "webp"):
            ext = "png"
        return {"cid": cid, "dataUri": f"data:image/{ext};base64,{b64}"}

    def pick_message_file(self):
        r = self.window.create_file_dialog(
            webview.FileDialog.OPEN,
            file_types=(
                "Text files (*.txt)",
                "HTML files (*.html;*.htm)",
                "All Files (*.*)",
            ),
        )
        if not r:
            return None
        fp = Path(r[0])
        text = decode_bytes(fp.read_bytes())
        if fp.suffix.lower() in (".html", ".htm"):
            return text
        return html_mod.escape(text).replace("\n", "<br>\n")

    # ── Sending ───────────────────────────────────────────────────

    def do_send(self, subject, html_content, used_cids, attach_path, template_name):
        if not self.env_config:
            return {"ok": False, "error": "Open Settings (\u2699) and configure SMTP first."}
        srv = self.env_config.get("SMTP_SERVER", "")
        port_s = self.env_config.get("SMTP_PORT", "587")
        email = self.env_config.get("SENDER_EMAIL", "")
        pw = self.env_config.get("SENDER_PASSWORD", "")
        if not srv:
            return {"ok": False, "error": "SMTP server not configured. Open Settings (\u2699)."}
        if not email or not pw:
            return {"ok": False, "error": "Email and password required. Open Settings (\u2699)."}
        if not self.recipients:
            return {"ok": False, "error": "Load recipients first."}
        if not subject or not subject.strip():
            return {"ok": False, "error": "Enter a subject."}
        try:
            port = int(port_s)
        except ValueError:
            return {"ok": False, "error": f"Invalid SMTP_PORT: {port_s}"}

        cc = self.cc_list or None

        inline = {c: self.image_paths[c] for c in (used_cids or []) if c in self.image_paths}
        html_tpl = (
            '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;font-size:14px;">'
            f"{html_content}</body></html>"
        )
        att = get_attachments(Path(attach_path)) if attach_path and attach_path.strip() else []
        total = len(self.recipients)
        try:
            delay = float(self.env_config.get("EMAIL_DELAY", "2"))
        except ValueError:
            delay = 2.0

        csv_path = self._last_csv_path or ""
        tpl_name = template_name or ""

        def progress(cur, tot, msg):
            self.window.evaluate_js(f"onProgress({cur},{tot},{json.dumps(msg)})")

        def worker():
            try:
                send_emails(
                    srv, port, email, pw, self.recipients, subject,
                    html_tpl, inline, att, delay, progress,
                    cc=cc,
                )
                if tpl_name:
                    from datetime import datetime, timezone
                    db.save_template_meta(
                        tpl_name,
                        last_sent_at=datetime.now(timezone.utc).isoformat(),
                        last_sent_count=total,
                        last_sent_recipients=csv_path,
                    )
            except Exception as e:
                self.window.evaluate_js(f"onProgress(0,0,{json.dumps(f'[ERROR] {e}')})")
            finally:
                self.window.evaluate_js("onSendComplete()")

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "total": total}

    # ── Templates ─────────────────────────────────────────────────

    def list_templates(self):
        d = db.get_templates_dir()
        return sorted(p.stem for p in d.glob("*.html"))

    def save_template(self, name, subject, html_content, attachment_dir, cc_file):
        if not name or not name.strip():
            return {"ok": False, "error": "Enter a template name."}
        safe = "".join(c for c in name.strip() if c.isalnum() or c in " _-").strip()
        if not safe:
            return {"ok": False, "error": "Invalid template name."}
        d = db.get_templates_dir()
        (d / f"{safe}.html").write_text(html_content, encoding="utf-8")
        db.save_template_meta(safe, subject=subject or "",
                              attachment_dir=attachment_dir or "",
                              cc_file=cc_file or "")
        return {"ok": True, "name": safe}

    def load_template(self, name):
        p = db.get_templates_dir() / f"{name}.html"
        if not p.is_file():
            return {"ok": False, "error": "Template not found."}
        html = p.read_text(encoding="utf-8")
        meta = db.get_template_meta(name)
        return {"ok": True, "html": html, **meta}

    def delete_template(self, name):
        p = db.get_templates_dir() / f"{name}.html"
        if p.is_file():
            p.unlink()
        db.delete_template_meta(name)
        return {"ok": True}

    def get_templates_dir(self):
        return str(db.get_templates_dir())

    # ── Auto-detection ────────────────────────────────────────────

    def auto_detect(self):
        res = {}
        sr = self.load_settings()
        if sr.get("ok"):
            res["settings"] = sr
        return res


# ── HTML / CSS / JS (embedded) ────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
:root {
    --bg: #1a1a2e; --panel: #16213e; --input: #0f3460;
    --border: #1a508b; --focus: #4a9eff;
    --text: #e0e0e0; --dim: #a0a0b8; --muted: #6a6a8a;
    --accent: #1a508b; --accent-h: #2d6ab5;
    --green: #1e7e34; --green-h: #28a745;
    --ok: #4ec24e; --warn: #e8a838; --err: #e84848;
}
* { margin:0; padding:0; box-sizing:border-box; }
html, body { height:100%; overflow:hidden; }
body {
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: var(--bg); color: var(--text);
    display: flex; flex-direction: column;
    padding: 14px; gap: 10px;
}
.sec {
    background: var(--panel); border-radius: 8px;
    padding: 12px 14px; flex-shrink: 0;
}
.sec.compose {
    flex: 1; display: flex; flex-direction: column;
    min-height: 0; overflow: hidden;
}
.sec-title {
    font-size: 13px; font-weight: 600; margin-bottom: 8px;
    text-transform: uppercase; letter-spacing: .5px; color: var(--dim);
}
.row { display: flex; align-items: center; gap: 8px; }
.row + .row { margin-top: 6px; }
.row label { font-size: 13px; white-space: nowrap; min-width: 80px; color: var(--dim); }
input[type="text"] {
    flex: 1; background: var(--input); border: 1px solid var(--border);
    border-radius: 4px; color: var(--text); padding: 7px 10px;
    font-size: 13px; font-family: inherit; outline: none;
    transition: border-color .2s;
}
input:focus { border-color: var(--focus); }
input::placeholder { color: var(--muted); }
.btn {
    background: var(--accent); border: none; color: var(--text);
    padding: 7px 14px; border-radius: 4px; cursor: pointer;
    font-size: 12px; font-family: inherit; white-space: nowrap;
    transition: background .15s;
}
.btn:hover { background: var(--accent-h); }
.btn-sm { padding: 5px 10px; font-size: 11px; }
.btn-green { background: var(--green); }
.btn-green:hover { background: var(--green-h); }
.btn-send {
    width: 100%; padding: 10px; font-size: 14px; font-weight: 600;
    border-radius: 6px; margin-top: 8px;
}
.btn-send:disabled { background: #333; cursor: not-allowed; opacity: .6; }
.status { font-size: 11px; margin-top: 4px; min-height: 16px; }
.status-ok { color: var(--ok); }
.status-warn { color: var(--warn); }
.toolbar {
    display: flex; align-items: center; gap: 6px;
    margin-top: 6px; margin-bottom: 4px;
}
.toolbar .hint { font-size: 11px; color: var(--muted); margin-left: auto; }
#editor {
    flex: 1; background: var(--input); border: 1px solid var(--border);
    border-radius: 4px; padding: 10px 12px; font-size: 13px;
    line-height: 1.6; color: var(--text); overflow-y: auto;
    outline: none; min-height: 80px; transition: border-color .2s;
}
#editor:focus { border-color: var(--focus); }
#editor img { max-width: 400px; display: block; margin: 4px 0; border-radius: 2px; }
#editor:empty::before {
    content: 'Type your message here... Use {name} for personalization';
    color: var(--muted); pointer-events: none;
}
/* ── Settings gear button ─────────────────────────────────── */
.top-bar { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
.top-bar .app-title { font-size: 15px; font-weight: 700; color: var(--text); }
.top-bar .spacer { flex: 1; }
.settings-status { font-size: 11px; color: var(--ok); }
.settings-status.not-configured { color: var(--warn); }
.gear-btn {
    background: none; border: 1px solid var(--border); border-radius: 6px;
    color: var(--dim); font-size: 20px; cursor: pointer; width: 36px; height: 36px;
    display: flex; align-items: center; justify-content: center;
    transition: background .15s, color .15s, border-color .15s;
}
.gear-btn:hover { background: var(--input); color: var(--text); border-color: var(--focus); }
/* ── Modal ────────────────────────────────────────────────── */
.modal-bg {
    display: none; position: fixed; inset: 0; background: rgba(0,0,0,.55);
    z-index: 100; align-items: center; justify-content: center;
}
.modal-bg.open { display: flex; }
.modal {
    background: var(--panel); border: 1px solid var(--border); border-radius: 12px;
    width: 480px; max-width: 92vw; box-shadow: 0 12px 48px rgba(0,0,0,.5);
    overflow: hidden; display: flex; flex-direction: column;
}
.modal-header {
    display: flex; align-items: center; gap: 8px;
    padding: 16px 20px 0; font-size: 16px; font-weight: 700; color: var(--text);
}
.modal-header span { font-size: 20px; }
/* ── Tabs ─────────────────────────────────────────────────── */
.tab-bar {
    display: flex; gap: 0; padding: 12px 20px 0; border-bottom: 1px solid var(--border);
}
.tab-btn {
    background: none; border: none; border-bottom: 2px solid transparent;
    color: var(--muted); font-size: 12px; font-family: inherit; font-weight: 600;
    padding: 8px 16px; cursor: pointer; text-transform: uppercase; letter-spacing: .4px;
    transition: color .15s, border-color .15s;
}
.tab-btn:hover { color: var(--dim); }
.tab-btn.active { color: var(--focus); border-bottom-color: var(--focus); }
.tab-panel { display: none; padding: 16px 20px 8px; }
.tab-panel.active { display: block; }
.modal .field { margin-bottom: 12px; }
.modal .field label { display: block; font-size: 12px; color: var(--dim); margin-bottom: 4px; font-weight: 500; }
.modal .field input {
    width: 100%; background: var(--input); border: 1px solid var(--border);
    border-radius: 6px; color: var(--text); padding: 8px 12px;
    font-size: 13px; font-family: inherit; outline: none;
    transition: border-color .2s;
}
.modal .field input:focus { border-color: var(--focus); }
.modal .field input::placeholder { color: var(--muted); }
.modal .field .hint { font-size: 10px; color: var(--muted); margin-top: 3px; }
.field-row { display: flex; gap: 8px; align-items: center; }
.field-row input { flex: 1; }
.modal-footer {
    display: flex; gap: 8px; padding: 12px 20px 16px; justify-content: flex-end;
    border-top: 1px solid var(--border); margin-top: 4px;
}
/* ── CSV Preview Modal ────────────────────────────────────── */
.modal.csv-modal { width: 620px; }
.csv-table-wrap {
    max-height: 220px; overflow: auto; border-radius: 6px;
    border: 1px solid var(--border); margin: 10px 0;
}
.csv-table {
    width: 100%; border-collapse: collapse; font-size: 12px;
}
.csv-table th {
    background: var(--input); color: var(--dim); padding: 6px 10px;
    text-align: left; font-weight: 600; position: sticky; top: 0;
    border-bottom: 1px solid var(--border);
}
.csv-table td {
    padding: 5px 10px; border-bottom: 1px solid rgba(255,255,255,.04);
    color: var(--text); white-space: nowrap; max-width: 200px;
    overflow: hidden; text-overflow: ellipsis;
}
.csv-table tr:hover td { background: rgba(74,158,255,.06); }
.csv-mapping { display: flex; gap: 16px; align-items: center; margin: 10px 0; flex-wrap: wrap; }
.csv-mapping label { font-size: 12px; color: var(--dim); }
.csv-mapping select {
    background: var(--input); color: var(--text); border: 1px solid var(--border);
    border-radius: 4px; padding: 5px 8px; font-size: 12px; font-family: inherit; outline: none;
}
.csv-mapping select:focus { border-color: var(--focus); }
.csv-warn {
    background: rgba(232,168,56,.1); border: 1px solid var(--warn); border-radius: 6px;
    padding: 8px 12px; font-size: 11px; color: var(--warn); margin: 8px 0; line-height: 1.5;
}
.csv-total { font-size: 11px; color: var(--dim); }
.info-row { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
.info-row .info-label { font-size: 11px; color: var(--muted); min-width: 80px; }
.info-row .info-value { font-size: 11px; color: var(--dim); word-break: break-all; }
.prog-row { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
.prog-bar {
    flex: 1; height: 5px; background: var(--input);
    border-radius: 3px; overflow: hidden;
}
.prog-fill { height: 100%; background: var(--focus); width: 0%; transition: width .3s; }
.prog-lbl { font-size: 11px; color: var(--dim); min-width: 44px; text-align: right; }
#log {
    background: #0a0a16; border-radius: 4px; padding: 8px 10px;
    font-family: 'Cascadia Code', Consolas, monospace; font-size: 11px;
    color: #999; max-height: 100px; overflow-y: auto; margin-top: 6px;
    line-height: 1.5; white-space: pre-wrap;
}
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #444; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #666; }
</style>
</head>
<body>

<!-- Top bar with settings gear -->
<div class="top-bar">
    <span class="app-title">MailBahn</span>
    <span class="spacer"></span>
    <span class="settings-status not-configured" id="settingsStatus">Not configured</span>
    <button class="gear-btn" onclick="openSettings()" title="Settings">&#9881;</button>
</div>

<!-- Settings Modal -->
<div class="modal-bg" id="settingsModal">
    <div class="modal">
        <div class="modal-header"><span>&#9881;</span> Settings</div>
        <div class="tab-bar">
            <button class="tab-btn active" onclick="switchTab('smtp')">Mail Server</button>
            <button class="tab-btn" onclick="switchTab('prefs')">Preferences</button>
            <button class="tab-btn" onclick="switchTab('about')">About</button>
        </div>
        <!-- Tab: SMTP -->
        <div class="tab-panel active" id="tab-smtp">
            <div class="field">
                <label>SMTP Server</label>
                <input type="text" id="cfgServer" placeholder="e.g. smtp.gmail.com">
            </div>
            <div class="field">
                <label>SMTP Port</label>
                <input type="text" id="cfgPort" placeholder="587">
                <div class="hint">Common ports: 587 (STARTTLS), 465 (SSL), 25 (unencrypted)</div>
            </div>
            <div class="field">
                <label>Sender Email</label>
                <input type="text" id="cfgEmail" placeholder="you@example.com">
            </div>
            <div class="field">
                <label>Sender Password</label>
                <input type="password" id="cfgPassword" placeholder="App password or email password">
                <div class="hint">For Gmail, use an App Password from your Google Account security settings</div>
                <div class="hint" style="color:var(--ok);font-style:italic;">&#128274; Password is stored securely in your OS credential manager</div>
            </div>
            <div style="margin-top:4px;">
                <button class="btn" onclick="testConnection()" id="testConnBtn" style="width:100%;">
                    &#128268; Test Connection
                </button>
                <div class="hint" id="testConnStatus" style="margin-top:6px;min-height:16px;"></div>
            </div>
        </div>
        <!-- Tab: Preferences -->
        <div class="tab-panel" id="tab-prefs">
            <div class="field">
                <label>Templates Directory</label>
                <div class="field-row">
                    <input type="text" id="cfgTemplatesDir" placeholder="Default: %APPDATA%/MailBahn/templates">
                    <button class="btn btn-sm" onclick="browseTemplatesDir()">Browse</button>
                </div>
                <div class="hint">Where your .html email templates are stored</div>
            </div>
            <div class="field">
                <label>Min Email Delay</label>
                <input type="text" id="cfgDelay" placeholder="2" style="width:80px;">
                <div class="hint">Minimum seconds to wait between sending each email (helps avoid rate limiting)</div>
            </div>
        </div>
        <!-- Tab: About -->
        <div class="tab-panel" id="tab-about">
            <div class="info-row">
                <span class="info-label">App</span>
                <span class="info-value">MailBahn v1.0</span>
            </div>
            <div class="info-row">
                <span class="info-label">Data Location</span>
                <span class="info-value" id="cfgDbPath"></span>
            </div>
            <div class="info-row">
                <span class="info-label">Templates</span>
                <span class="info-value" id="cfgTplPath"></span>
            </div>
            <div style="margin-top:16px;font-size:11px;color:var(--muted);line-height:1.6;">
                Settings and paths are stored in a local SQLite database.<br>
                Templates are saved as .html files in the templates directory.
            </div>
        </div>
        <div class="modal-footer">
            <button class="btn" onclick="closeSettings()">Cancel</button>
            <button class="btn btn-green" onclick="saveSettings()">Save</button>
        </div>
    </div>
</div>

<!-- CSV Preview Modal -->
<div class="modal-bg" id="csvModal">
    <div class="modal csv-modal">
        <div class="modal-header"><span>&#128196;</span> CSV Preview</div>
        <div style="padding:12px 20px 0;">
            <div class="csv-total" id="csvTotal"></div>
            <div class="csv-table-wrap">
                <table class="csv-table" id="csvTable"><tbody></tbody></table>
            </div>
            <div class="csv-mapping">
                <div>
                    <label>Name column:</label>
                    <select id="csvNameCol"></select>
                </div>
                <div>
                    <label>Email column:</label>
                    <select id="csvEmailCol"></select>
                </div>
            </div>
            <div id="csvDupeWarn"></div>
        </div>
        <div class="modal-footer">
            <button class="btn" onclick="closeCsvModal()">Cancel</button>
            <button class="btn btn-green" onclick="confirmCsv()">Confirm</button>
        </div>
    </div>
</div>

<!-- Alert Modal -->
<div class="modal-bg" id="alertModal" style="z-index: 999;">
    <div class="modal" style="width: 380px;">
        <div class="modal-header"><span style="color: var(--err);">&#9888;</span>&nbsp;<span id="alertTitle" style="font-size: 16px;">Alert</span></div>
        <div style="padding: 16px 20px; font-size: 13px; line-height: 1.5; white-space: pre-wrap; color: var(--text);" id="alertMsg"></div>
        <div class="modal-footer">
            <button class="btn btn-green" onclick="closeAlert()">OK</button>
        </div>
    </div>
</div>

<!-- Recipients -->
<div class="sec">
    <div class="sec-title">Recipients</div>
    <div class="row">
        <label>CSV File:</label>
        <input type="text" id="csvPath" placeholder="Path to recipients.csv">
        <button class="btn" onclick="browseCsv()">Browse</button>
        <button class="btn btn-green" onclick="loadCsv()">Load</button>
    </div>
    <div class="status" id="csvStatus"></div>
</div>

<!-- Compose -->
<div class="sec compose">
    <div class="sec-title">Compose</div>
    <div class="row">
        <label>Subject:</label>
        <input type="text" id="subject" placeholder="Email subject line">
    </div>
    <div class="row">
        <label>CC (CSV):</label>
        <input type="text" id="ccPath" placeholder="Path to cc.csv (optional)">
        <button class="btn" onclick="browseCc()">Browse</button>
        <button class="btn btn-green" onclick="loadCc()">Load</button>
    </div>
    <div class="status" id="ccStatus"></div>

<!-- CC Preview Modal -->
<div class="modal-bg" id="ccModal">
    <div class="modal csv-modal">
        <div class="modal-header"><span>&#128196;</span> CC CSV Preview</div>
        <div style="padding:12px 20px 0;">
            <div class="csv-total" id="ccTotal"></div>
            <div class="csv-table-wrap">
                <table class="csv-table" id="ccTable"><tbody></tbody></table>
            </div>
            <div class="csv-mapping">
                <div>
                    <label>Email column:</label>
                    <select id="ccEmailCol"></select>
                </div>
            </div>
            <div id="ccDupeWarn"></div>
        </div>
        <div class="modal-footer">
            <button class="btn" onclick="closeCcModal()">Cancel</button>
            <button class="btn btn-green" onclick="confirmCc()">Confirm</button>
        </div>
    </div>
</div>
    <div class="toolbar">
        <span style="font-size:12px;color:var(--dim)">Message:</span>
        <button class="btn btn-sm" onclick="loadFile()">Load File</button>
        <button class="btn btn-sm" onclick="insertImage()">Insert Image</button>
        <span class="tpl-group" style="margin-left:auto;display:flex;gap:4px;align-items:center;">
            <select id="tplSelect" style="background:var(--input);color:var(--text);border:1px solid var(--border);
                border-radius:4px;padding:4px 6px;font-size:11px;font-family:inherit;outline:none;">
                <option value="">— Templates —</option>
            </select>
            <button class="btn btn-sm" onclick="loadTemplate()">Load</button>
            <button class="btn btn-sm btn-green" onclick="saveTemplate()">Save</button>
            <button class="btn btn-sm" onclick="deleteTemplate()" style="background:#6e2020;">Del</button>
        </span>
    </div>
    <div id="editor" contenteditable="true"></div>
</div>

<!-- Bottom -->
<div class="sec">
    <div class="row">
        <label>Attachments:</label>
        <input type="text" id="attachPath" placeholder="Folder path (all files in folder will be attached)">
        <button class="btn" onclick="browseAttach()">Browse</button>
    </div>
    <button class="btn btn-send" id="sendBtn" onclick="doSend()">Send Emails</button>
    <div class="prog-row">
        <div class="prog-bar"><div class="prog-fill" id="progFill"></div></div>
        <span class="prog-lbl" id="progLbl">0 / 0</span>
    </div>
    <div id="log"></div>
</div>

<script>
const $ = id => document.getElementById(id);

/* ── keep placeholder working after content is cleared ────── */
$('editor').addEventListener('input', function() {
    if (this.innerHTML === '<br>' || this.innerHTML === '<div><br></div>') {
        this.innerHTML = '';
    }
});

/* ── auto-detect on startup ───────────────────────────────── */
window.addEventListener('pywebviewready', async () => {
    const r = await pywebview.api.auto_detect();
    if (r && r.settings && r.settings.ok) { applySettings(r.settings); }
    await refreshTemplates();
});

/* ── Settings Modal ───────────────────────────────────────── */
function switchTab(name) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelector('[onclick*="' + name + '"]').classList.add('active');
    document.getElementById('tab-' + name).classList.add('active');
}
async function openSettings() {
    $('settingsModal').classList.add('open');
    switchTab('smtp');
    $('cfgDbPath').textContent = await pywebview.api.get_db_location();
    $('cfgTplPath').textContent = await pywebview.api.get_templates_dir();
    $('cfgServer').focus();
}
function closeSettings() {
    $('settingsModal').classList.remove('open');
}
$('settingsModal').addEventListener('click', function(e) {
    if (e.target === this) closeSettings();
});
async function browseTemplatesDir() {
    const p = await pywebview.api.browse_folder();
    if (p) $('cfgTemplatesDir').value = p;
}
async function testConnection() {
    const btn = $('testConnBtn');
    const status = $('testConnStatus');
    // Save current values first so backend has them
    const settings = {
        SMTP_SERVER: $('cfgServer').value.trim(),
        SMTP_PORT: $('cfgPort').value.trim() || '587',
        SENDER_EMAIL: $('cfgEmail').value.trim(),
        SENDER_PASSWORD: $('cfgPassword').value,
        TEMPLATES_DIR: $('cfgTemplatesDir').value.trim(),
        EMAIL_DELAY: $('cfgDelay').value.trim() || '2',
    };
    if (!settings.SMTP_SERVER || !settings.SENDER_EMAIL || !settings.SENDER_PASSWORD) {
        status.style.color = 'var(--warn)';
        status.textContent = 'Fill in server, email, and password first.';
        return;
    }
    btn.disabled = true;
    btn.textContent = 'Testing...';
    status.style.color = 'var(--dim)';
    status.textContent = 'Connecting...';
    await pywebview.api.save_settings(settings);
    applySettings(settings);
    const r = await pywebview.api.test_connection();
    if (r.ok) {
        status.style.color = 'var(--ok)';
        status.textContent = '\\u2714 ' + r.message;
    } else {
        status.style.color = 'var(--err)';
        status.textContent = '\\u2718 ' + r.error;
    }
    btn.disabled = false;
    btn.textContent = '\\ud83d\\udd0c Test Connection';
}
async function saveSettings() {
    const settings = {
        SMTP_SERVER: $('cfgServer').value.trim(),
        SMTP_PORT: $('cfgPort').value.trim() || '587',
        SENDER_EMAIL: $('cfgEmail').value.trim(),
        SENDER_PASSWORD: $('cfgPassword').value,
        TEMPLATES_DIR: $('cfgTemplatesDir').value.trim(),
        EMAIL_DELAY: $('cfgDelay').value.trim() || '2',
    };
    if (!settings.SMTP_SERVER || !settings.SENDER_EMAIL || !settings.SENDER_PASSWORD) {
        showAlert('Server, email, and password are required.', 'Settings Error');
        return;
    }
    const r = await pywebview.api.save_settings(settings);
    if (r.ok) {
        applySettings(settings);
        closeSettings();
        await refreshTemplates();
        addLog('[OK] Settings saved.');
    } else {
        showAlert('Failed to save: ' + (r.error || 'Unknown error'), 'Settings Error');
    }
}
function applySettings(s) {
    const server = s.SMTP_SERVER || '';
    const port = s.SMTP_PORT || '587';
    const email = s.SENDER_EMAIL || '';
    if (server && email) {
        $('settingsStatus').textContent = email + ' \\u2014 ' + server + ':' + port;
        $('settingsStatus').className = 'settings-status';
    }
    $('cfgServer').value = server;
    $('cfgPort').value = port;
    $('cfgEmail').value = email;
    if (s.SENDER_PASSWORD) $('cfgPassword').value = s.SENDER_PASSWORD;
    if (s.TEMPLATES_DIR) $('cfgTemplatesDir').value = s.TEMPLATES_DIR;
    if (s.EMAIL_DELAY) $('cfgDelay').value = s.EMAIL_DELAY;
}

/* ── Recipients & CSV Preview ─────────────────────────────── */
let _csvPath = '';
async function browseCsv() {
    const p = await pywebview.api.browse_csv();
    if (p) $('csvPath').value = p;
}
async function loadCsv() {
    const p = $('csvPath').value.trim();
    if (!p) return setStatus('csvStatus', 'Enter a path first.', 'warn');
    _csvPath = p;
    const r = await pywebview.api.preview_csv_file(p);
    if (!r.ok) return setStatus('csvStatus', r.error, 'warn');
    showCsvModal(r);
}
function showCsvModal(data) {
    const headers = data.headers;
    const rows = data.rows;
    // Build table
    let html = '<thead><tr>' + headers.map(h => '<th>' + h + '</th>').join('') + '</tr></thead><tbody>';
    rows.forEach(row => {
        html += '<tr>' + headers.map(h => '<td>' + (row[h] || '') + '</td>').join('') + '</tr>';
    });
    html += '</tbody>';
    $('csvTable').innerHTML = html;
    $('csvTotal').textContent = data.total + ' row' + (data.total !== 1 ? 's' : '') + ' total' +
        (data.rows.length < data.total ? ' (showing first ' + data.rows.length + ')' : '');
    // Populate dropdowns
    function fillSelect(id, guess) {
        const sel = $(id);
        sel.innerHTML = '<option value="">(none)</option>';
        headers.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h; opt.textContent = h;
            if (h.toLowerCase() === guess) opt.selected = true;
            sel.appendChild(opt);
        });
    }
    fillSelect('csvNameCol', 'name');
    fillSelect('csvEmailCol', 'email');
    $('csvDupeWarn').innerHTML = '';
    $('csvModal').classList.add('open');
}
function closeCsvModal() {
    $('csvModal').classList.remove('open');
}
$('csvModal').addEventListener('click', function(e) {
    if (e.target === this) closeCsvModal();
});
async function confirmCsv() {
    const nameCol = $('csvNameCol').value;
    const emailCol = $('csvEmailCol').value;
    if (!emailCol) {
        $('csvDupeWarn').innerHTML = '<div class="csv-warn">Select an email column.</div>';
        return;
    }
    const r = await pywebview.api.confirm_csv(_csvPath, nameCol, emailCol);
    if (!r.ok) {
        showAlert(r.error, "CSV Data Error");
        $('csvDupeWarn').innerHTML = '<div class="csv-warn">' + r.error + '</div>';
        return;
    }
    let warn = '';
    if (r.duplicates && r.duplicates.length > 0) {
        warn = '\\u26a0 Duplicates: ' + r.duplicates.join(', ');
        $('csvDupeWarn').innerHTML = '<div class="csv-warn">' + warn + '</div>';
        // Still load, but warn — user can cancel or proceed
    }
    closeCsvModal();
    setStatus('csvStatus', r.count + ' loaded \\u2014 ' + r.preview +
        (r.duplicates && r.duplicates.length ? ' (\\u26a0 ' + r.duplicates.length + ' duplicate' +
        (r.duplicates.length > 1 ? 's' : '') + ')' : ''), r.duplicates && r.duplicates.length ? 'warn' : 'ok');
}
function applyCsv(r) {
    if (!r.ok) return setStatus('csvStatus', r.error, 'warn');
    setStatus('csvStatus', r.count + ' loaded \\u2014 ' + r.preview, 'ok');
}

/* ── CC & CC Preview ──────────────────────────────────────── */
let _ccPath = '';
async function browseCc() {
    const p = await pywebview.api.browse_cc_csv();
    if (p) $('ccPath').value = p;
}
async function loadCc() {
    const p = $('ccPath').value.trim();
    if (!p) return setStatus('ccStatus', 'Enter a path first.', 'warn');
    _ccPath = p;
    const r = await pywebview.api.preview_cc_file(p);
    if (!r.ok) return setStatus('ccStatus', r.error, 'warn');
    showCcModal(r);
}
function showCcModal(data) {
    const headers = data.headers;
    const rows = data.rows;
    let html = '<thead><tr>' + headers.map(h => '<th>' + h + '</th>').join('') + '</tr></thead><tbody>';
    rows.forEach(row => {
        html += '<tr>' + headers.map(h => '<td>' + (row[h] || '') + '</td>').join('') + '</tr>';
    });
    html += '</tbody>';
    $('ccTable').innerHTML = html;
    $('ccTotal').textContent = data.total + ' row' + (data.total !== 1 ? 's' : '') + ' total' +
        (data.rows.length < data.total ? ' (showing first ' + data.rows.length + ')' : '');
    const sel = $('ccEmailCol');
    sel.innerHTML = '<option value="">(none)</option>';
    headers.forEach(h => {
        const opt = document.createElement('option');
        opt.value = h; opt.textContent = h;
        if (h.toLowerCase() === 'email') opt.selected = true;
        sel.appendChild(opt);
    });
    $('ccDupeWarn').innerHTML = '';
    $('ccModal').classList.add('open');
}
function closeCcModal() {
    $('ccModal').classList.remove('open');
}
$('ccModal').addEventListener('click', function(e) {
    if (e.target === this) closeCcModal();
});
async function confirmCc() {
    const emailCol = $('ccEmailCol').value;
    if (!emailCol) {
        $('ccDupeWarn').innerHTML = '<div class="csv-warn">Select an email column.</div>';
        return;
    }
    const r = await pywebview.api.confirm_cc(_ccPath, emailCol);
    if (!r.ok) {
        showAlert(r.error, "CSV Data Error");
        $('ccDupeWarn').innerHTML = '<div class="csv-warn">' + r.error + '</div>';
        return;
    }
    if (r.duplicates && r.duplicates.length > 0) {
        $('ccDupeWarn').innerHTML = '<div class="csv-warn">\\u26a0 Duplicates: ' + r.duplicates.join(', ') + '</div>';
    }
    closeCcModal();
    setStatus('ccStatus', r.count + ' loaded \\u2014 ' + r.preview +
        (r.duplicates && r.duplicates.length ? ' (\\u26a0 ' + r.duplicates.length + ' duplicate' +
        (r.duplicates.length > 1 ? 's' : '') + ')' : ''), r.duplicates && r.duplicates.length ? 'warn' : 'ok');
}
function applyCc(r) {
    if (!r.ok) return setStatus('ccStatus', r.error, 'warn');
    setStatus('ccStatus', r.count + ' loaded \\u2014 ' + r.preview, 'ok');
}

/* ── Attachments ──────────────────────────────────────────── */
async function browseAttach() {
    const p = await pywebview.api.browse_folder();
    if (p) $('attachPath').value = p;
}

/* ── Editor: insert image ─────────────────────────────────── */
async function insertImage() {
    const r = await pywebview.api.pick_image();
    if (!r) return;
    const editor = $('editor');
    editor.focus();
    const img = document.createElement('img');
    img.src = r.dataUri;
    img.setAttribute('data-cid', r.cid);
    const sel = window.getSelection();
    if (sel.rangeCount) {
        const range = sel.getRangeAt(0);
        range.deleteContents();
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
    } else {
        editor.appendChild(img);
    }
}

/* ── Editor: load from file ───────────────────────────────── */
async function loadFile() {
    const content = await pywebview.api.pick_message_file();
    if (content != null) $('editor').innerHTML = content;
}

/* ── Send ─────────────────────────────────────────────────── */
function getContent() {
    const clone = $('editor').cloneNode(true);
    const cids = [];
    clone.querySelectorAll('img[data-cid]').forEach(img => {
        const cid = img.getAttribute('data-cid');
        cids.push(cid);
        img.src = 'cid:' + cid;
        img.removeAttribute('data-cid');
        img.removeAttribute('style');
    });
    return { html: clone.innerHTML, cids };
}

async function doSend() {
    const btn = $('sendBtn');
    if (btn.disabled) return;
    const subject = $('subject').value;
    const { html, cids } = getContent();
    const attachPath = $('attachPath').value;
    clearLog();
    btn.disabled = true;
    btn.textContent = 'Sending...';
    const r = await pywebview.api.do_send(subject, html, cids, attachPath, $('tplSelect').value);
    if (!r.ok) {
        addLog('[ERROR] ' + r.error);
        btn.disabled = false;
        btn.textContent = 'Send Emails';
    }
}

/* ── Progress (called from Python via evaluate_js) ────────── */
function onProgress(cur, tot, msg) {
    const pct = tot > 0 ? (cur / tot * 100) : 0;
    $('progFill').style.width = pct + '%';
    $('progLbl').textContent = cur + ' / ' + tot;
    addLog(msg);
}
function onSendComplete() {
    $('sendBtn').disabled = false;
    $('sendBtn').textContent = 'Send Emails';
}

/* ── Templates ────────────────────────────────────────────── */
async function refreshTemplates() {
    const list = await pywebview.api.list_templates();
    const sel = $('tplSelect');
    sel.innerHTML = '<option value="">\\u2014 Templates \\u2014</option>';
    list.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name; opt.textContent = name;
        sel.appendChild(opt);
    });
}
async function saveTemplate() {
    const name = prompt('Template name:');
    if (!name) return;
    const subject = $('subject').value;
    const html = $('editor').innerHTML;
    const attachDir = $('attachPath').value;
    const ccFile = $('ccPath').value;
    const r = await pywebview.api.save_template(name, subject, html, attachDir, ccFile);
    if (r.ok) {
        await refreshTemplates();
        $('tplSelect').value = r.name;
        addLog('[OK] Template saved: ' + r.name);
    } else { addLog('[ERROR] ' + r.error); }
}
async function loadTemplate() {
    const name = $('tplSelect').value;
    if (!name) return;
    const r = await pywebview.api.load_template(name);
    if (r.ok) {
        $('editor').innerHTML = r.html;
        if (r.subject) $('subject').value = r.subject;
        if (r.attachment_dir) $('attachPath').value = r.attachment_dir;
        if (r.cc_file) $('ccPath').value = r.cc_file;
    } else addLog('[ERROR] ' + r.error);
}
async function deleteTemplate() {
    const name = $('tplSelect').value;
    if (!name) return;
    if (!confirm('Delete template "' + name + '"?')) return;
    await pywebview.api.delete_template(name);
    await refreshTemplates();
    addLog('[OK] Template deleted: ' + name);
}

/* ── Helpers ──────────────────────────────────────────────── */
function showAlert(msg, title) {
    $('alertTitle').textContent = title || 'Alert';
    $('alertMsg').textContent = msg;
    $('alertModal').classList.add('open');
}

function closeAlert() {
    $('alertModal').classList.remove('open');
}
function setStatus(id, text, type) {
    const el = $(id);
    el.textContent = text;
    el.className = 'status status-' + type;
}
function addLog(msg) {
    const log = $('log');
    if (log.textContent) log.textContent += '\\n';
    log.textContent += msg;
    log.scrollTop = log.scrollHeight;
}
function clearLog() {
    $('log').textContent = '';
    $('progFill').style.width = '0%';
    $('progLbl').textContent = '0 / 0';
}
</script>
</body>
</html>"""


if __name__ == "__main__":
    logging.getLogger("pywebview").setLevel(logging.CRITICAL)
    api = Api()
    window = webview.create_window(
        "MailBahn", html=HTML, js_api=api,
        width=860, height=750, min_size=(720, 600),
    )
    api.window = window
    webview.start()
