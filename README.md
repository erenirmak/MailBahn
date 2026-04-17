# Auto-Mail

A desktop app for sending personalized HTML emails with attachments to multiple recipients.

## Features

- Rich HTML editor with inline images and `{name}` personalization
- SMTP settings stored in a local SQLite database (`%APPDATA%/auto-mail/`)
- Reusable email templates saved as `.html` files
- CC support, file attachments (entire folder), configurable send delay
- No `.env` files — everything is configured through the Settings UI (⚙️)

## Setup

```
uv sync
```

## Run

```
uv run main.py
```

On first launch, click the ⚙️ gear icon to configure your SMTP server, then load a recipients CSV and compose your email.

### Recipients CSV format

```csv
name,email
Alice,alice@example.com
Bob,bob@example.com
```

## Build standalone .exe

```
uv run pyinstaller --onefile --noconsole --name auto-mail main.py
```

The script connects via SMTP (TLS on port 587), sends one email per recipient with a 2-second delay between sends, and prints the status of each delivery.
