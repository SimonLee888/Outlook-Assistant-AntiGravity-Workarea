from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import re
from typing import Any


@dataclass(slots=True)
class FolderNode:
    id: str
    name: str
    path: str
    parent_id: str | None
    children: list["FolderNode"] = field(default_factory=list)


@dataclass(slots=True)
class AttachmentInfo:
    id: str
    name: str
    size: int | None
    extension: str


@dataclass(slots=True)
class MessageInfo:
    id: str
    folder_id: str
    folder_path: str
    sender: str
    subject: str
    received: datetime | None
    has_attachments: bool
    attachment_count: int
    attachments: list[AttachmentInfo] = field(default_factory=list)
    attachments_probed: bool = False
    thread_has_attachments: bool = False
    body_text: str = ""
    body_html: str = ""
    headers_text: str = ""
    preview_loaded: bool = False
    backend_ref: Any = None


@dataclass(slots=True)
class ExtractionFilters:
    sender_contains: str = ""
    subject_contains: str = ""
    attachment_name_contains: str = ""
    extension: str = ""
    has_attachments_only: bool = True
    date_from: datetime | None = None
    date_to: datetime | None = None

    def matches(self, message: MessageInfo) -> bool:
        if self.has_attachments_only and not message.has_attachments:
            return False

        if self.sender_contains:
            sender = message.sender or ""
            if self.sender_contains.lower() not in sender.lower():
                if not self._matches_sender_alias(sender, self.sender_contains):
                    chain_blob = f"{message.headers_text}\n{message.body_text}"
                    if self.sender_contains.lower() not in chain_blob.lower():
                        if not self._matches_sender_alias(chain_blob, self.sender_contains):
                            return False

        if self.subject_contains and self.subject_contains.lower() not in (message.subject or "").lower():
            return False

        if self.date_from and (message.received is None or message.received < self.date_from):
            return False

        if self.date_to and (message.received is None or message.received > self.date_to):
            return False

        if self.extension:
            wanted_extension = self.extension.lower()
            if not wanted_extension.startswith("."):
                wanted_extension = f".{wanted_extension}"
            if not any(att.extension.lower() == wanted_extension for att in message.attachments):
                return False

        if self.attachment_name_contains:
            needle = self.attachment_name_contains.lower()
            if not any(needle in att.name.lower() for att in message.attachments):
                return False

        return True

    @staticmethod
    def _matches_sender_alias(sender: str, query: str) -> bool:
        sender_clean = re.sub(r"[^a-z0-9]+", "", sender.lower())
        query_clean = re.sub(r"[^a-z0-9]+", "", query.lower())
        if not query_clean:
            return True
        if query_clean in sender_clean:
            return True

        words = re.findall(r"[a-z0-9]+", sender.lower())
        if len(words) >= 2:
            initial_last = f"{words[0][0]}{words[-1]}"
            if query_clean in initial_last:
                return True
            first_last = f"{words[0]}{words[-1]}"
            if query_clean in first_last:
                return True
        return False


@dataclass(slots=True)
class ExtractionResult:
    scanned_messages: int = 0
    matched_messages: int = 0
    exported_attachments: int = 0
    exported_files: list[Path] = field(default_factory=list)
    manifest_path: Path | None = None
