from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Iterable

from ..models import FolderNode, MessageInfo


class BackendError(RuntimeError):
    """Raised when the selected source backend cannot complete an action."""


class MailBackend(ABC):
    @property
    @abstractmethod
    def source_name(self) -> str:
        raise NotImplementedError

    @abstractmethod
    def list_folders(self) -> list[FolderNode]:
        raise NotImplementedError

    @abstractmethod
    def list_messages(self, folder_id: str) -> list[MessageInfo]:
        raise NotImplementedError

    @abstractmethod
    def iter_messages(self, folder_id: str, recursive: bool = False) -> Iterable[MessageInfo]:
        raise NotImplementedError

    @abstractmethod
    def save_attachments(self, message: MessageInfo, output_dir: Path) -> list[Path]:
        raise NotImplementedError

    def ensure_attachment_metadata(self, message: MessageInfo) -> None:
        """Populate attachment metadata (name, extension) when needed by filters/UI."""
        return

    def load_message_preview(self, message: MessageInfo) -> None:
        """Populate message body/headers lazily for preview."""
        return

    def close(self) -> None:
        pass
