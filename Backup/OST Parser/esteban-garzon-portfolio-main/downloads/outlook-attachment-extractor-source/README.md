# Outlook Attachment Extractor

Sanitized public source package for a Windows-based Outlook OST attachment extraction utility.

## What is included

- GUI entry point
- CLI entry point
- extraction logic
- OST backend implementation
- Python package structure
- requirements file

## What is intentionally excluded

- extracted mailbox samples
- temporary output folders
- build artifacts
- installers and binaries
- logs, reports, and generated PDFs
- cached files and local machine traces

## Typical use case

Use this tool to inspect an `.ost` file, browse folders, filter messages, and export attachments in a more structured workflow.

## Run locally

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```
