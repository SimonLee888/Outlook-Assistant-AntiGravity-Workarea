from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Callable

from .backends.base import MailBackend
from .models import ExtractionFilters, ExtractionResult, MessageInfo

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]+')


def sanitize_segment(value: str, default: str = "_", max_length: int = 80) -> str:
    candidate = INVALID_FILENAME_CHARS.sub("_", (value or "").strip())
    candidate = candidate.strip(" .")
    if not candidate:
        candidate = default
    if len(candidate) > max_length:
        candidate = candidate[:max_length].rstrip(" .")
    return candidate or default


def build_message_output_dir(base_output: Path, message: MessageInfo) -> Path:
    folder_parts = [sanitize_segment(part, default="_", max_length=60) for part in message.folder_path.replace("\\", "/").split("/") if part]
    folder_path = Path(*folder_parts) if folder_parts else Path("root")
    date_part = message.received.strftime("%Y%m%d_%H%M%S") if message.received else "unknown_date"
    subject_part = sanitize_segment(message.subject, default="sin_asunto", max_length=70)
    message_part = sanitize_segment(f"{date_part}_{subject_part}", max_length=100)
    return base_output / folder_path / message_part


def extract_attachments(
    backend: MailBackend,
    folder_id: str,
    output_dir: Path,
    filters: ExtractionFilters,
    recursive: bool = True,
    progress_callback: Callable[[str], None] | None = None,
) -> ExtractionResult:
    result = ExtractionResult()
    output_dir = output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest_rows: list[dict[str, str]] = []

    for message in backend.iter_messages(folder_id, recursive=recursive):
        result.scanned_messages += 1
        if progress_callback and result.scanned_messages % 100 == 0:
            progress_callback(f"Escaneados {result.scanned_messages} mensajes...")

        if not filters.matches(message):
            continue

        result.matched_messages += 1
        destination_dir = build_message_output_dir(output_dir, message)
        saved_files = backend.save_attachments(message, destination_dir)
        if not saved_files:
            continue

        for file_path in saved_files:
            result.exported_files.append(file_path)
            result.exported_attachments += 1
            manifest_rows.append(
                {
                    "message_id": message.id,
                    "folder_path": message.folder_path,
                    "sender": message.sender,
                    "subject": message.subject,
                    "received": message.received.isoformat() if message.received else "",
                    "saved_path": str(file_path),
                }
            )

    manifest_path = output_dir / "attachment_manifest.csv"
    with manifest_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["message_id", "folder_path", "sender", "subject", "received", "saved_path"],
        )
        writer.writeheader()
        writer.writerows(manifest_rows)
    result.manifest_path = manifest_path
    if progress_callback:
        progress_callback(
            f"Finalizado. Mensajes: {result.scanned_messages}, coinciden: {result.matched_messages}, adjuntos: {result.exported_attachments}"
        )
    return result

