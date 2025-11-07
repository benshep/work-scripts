# Work Scripts

A grab‑bag of small, pragmatic scripts I use at work to automate repetitive tasks across Outlook/Exchange, Zoom, web portals, spreadsheets, and file management. Most tools are Python; a few Windows conveniences are written in AutoHotkey.

> Heads‑up: Many scripts assume a Windows workstation with Outlook installed and access to internal STFC systems; some also rely on personal environment variables or small, unversioned helper files (e.g., credentials/API keys). See the per‑script notes below and inline docstrings for specifics.


## Contents

Highlighted scripts (not exhaustive):

- **join_zoom_meeting.pyw** – Instantly joins the Zoom or Teams meeting closest to now found in your Outlook calendar (link in subject or body). If there are several possible meetings to join, show a menu. Also starts a notes file. 
- **start_meeting_notes.pyw** – Looks at today’s calendar, lets you pick the relevant meeting, and creates a Markdown notes file (title/date/attendees pre‑filled). If it finds an Indico agenda, it pulls items into Markdown automatically. Files are placed in a sensible `Documents` subfolder inferred from the meeting text.
- **get_risk_assessments.py** – Automates login to the SHE Assure site and bulk‑downloads project risk assessment attachments, filing them into `Documents\Safety\RAs`.
- **AutoHotkey.ahk** – Handy Windows hotkeys to launch/activate apps and other small conveniences.

The repository root shows additional helpers for events, Oracle/finance data pulls, spreadsheets, image/OCR and more: browse the file list for ideas.

## Quick Start

### 1) Clone

```shell
git clone https://github.com/benshep/work-scripts.git
cd work-scripts
```

### 2) Environment & dependencies

Python 3.10+ is required, and newer versions are recommended. Install Python dependencies with:

```shell
pip install -r requirements.txt
```

Most Python scripts assume Windows (primarily for Outlook COM and `os.startfile`). See the script headers for details.

### 3) Optional: AutoHotkey

If you want the Windows hotkeys, install [AutoHotkey](https://www.autohotkey.com/) v1 and load `AutoHotkey.ahk`.

------

### Configuration & Secrets

Some scripts expect small local modules or environment variables (e.g., `stfc_credentials.py` for usernames/passwords or `pushbullet_api_key.py` for an API key). These aren’t committed; create them locally following the import names used in each script.

------

## Tips & Conventions

- **Platform**: Many scripts are Windows‑centric (Outlook COM, AutoHotkey). Test on your machine and adapt if you’re on Linux/macOS.
- **Calendars**: Outlook helpers live in `outlook.py` and are reused across several tools (parsing events, checking OoO/WFH, etc.). If you have multiple mailboxes, adjust the calls accordingly.
- **Filing**: Scripts try to create sensible subfolders (e.g., under `Documents`) and sanitise filenames to avoid illegal characters.

------

## Contributing

This is primarily a personal toolkit, but issues and PRs that improve robustness, portability, or documentation are welcome. If you’re adapting for a different org:

- Extract org‑specific constants (mail domains, group names, SharePoint/portal URLs) into a small config module.
- Avoid committing credentials or tokens; import from a local, ignored file as done here.