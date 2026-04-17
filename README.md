# MailBahn

Auto-Mail is a desktop application designed for sending personalized HTML emails with attachments to multiple recipients. It provides a user-friendly interface for managing email templates, recipients, and SMTP settings.

## Features

- **Rich HTML Email Editor**: Compose emails with inline images and `{name}` placeholders for personalization.
- **SMTP Configuration**: Store SMTP settings securely in a local SQLite database.
- **Reusable Templates**: Save and reuse email templates as `.html` files.
- **CSV Support**: Load recipient and CC lists from CSV files.
  - **Flexible Column Mapping**: Select which columns in your CSV file correspond to recipient names and email addresses.
- **Attachments**: Attach files from a folder to your emails.
- **Send Delay**: Configure a delay between emails to avoid rate-limiting.
- **Secure Credentials**: Passwords are stored securely using the OS keyring.

## Installation

### Prerequisites

- Python 3.13 or higher
- Dependencies:
  - `keyring`
  - `pywebview`

### Install Dependencies

Run the following command to install the required dependencies:

```bash
pip install -r requirements.txt
```

Alternatively, use the dependencies listed in `pyproject.toml`:

```bash
pip install keyring>=25.7.0 pywebview>=6.2
```

## Usage

### Running the Application

To start the application, run:

```bash
python main.py
```

On the first launch:
1. Click the ⚙️ gear icon to configure your SMTP server.
2. Load a recipients CSV file.
3. Compose your email and send.

### Recipients CSV Format

The recipients CSV file can have flexible column names. By default, it expects the following format:

```csv
name,email
Alice,alice@example.com
Bob,bob@example.com
```

However, you can specify which columns correspond to recipient names and email addresses when loading the file.

### CC CSV Format

The CC CSV file can also have flexible column names. By default, it expects the following format:

```csv
email
cc1@example.com
cc2@example.com
```

You can select the column to use for email addresses when loading the file.

### Email Personalization

Use `{name}` in your email template to personalize messages for each recipient.

### Attachments

Specify a folder path to attach all files within that folder to your emails.

## Building a Standalone Executable

To build a standalone `.exe` file for distribution:

1. Install `pyinstaller`:

   ```bash
   pip install pyinstaller
   ```

2. Run the following command:

   ```bash
   pyinstaller --onefile --noconsole --name MailBahn main.py
   ```

The executable will be created in the `dist` folder.

## Project Structure

- **`main.py`**: Entry point for the application.
- **`mailer.py`**: Handles email composition and sending.
- **`db.py`**: Manages SQLite database for settings and templates.
- **`csvutils.py`**: Utilities for loading and validating CSV files.

## Development

### Dependencies for Development

Install development dependencies:

```bash
pip install pyinstaller>=6.19.0
```

### Running Tests

Use the provided test files in `test_files/` to validate the application's functionality.

## License

This project is licensed under the MIT License.
