from __future__ import annotations

from datetime import datetime
from io import BytesIO
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
import re
from typing import Any, Iterable
import zipfile

from .base import BackendError, MailBackend
from ..models import AttachmentInfo, FolderNode, MessageInfo


def _safe_call(value: Any, default: Any = None) -> Any:
    try:
        if callable(value):
            return value()
        return value
    except Exception:
        return default


def _pick(obj: Any, names: list[str], default: Any = None) -> Any:
    for name in names:
        try:
            attr = getattr(obj, name)
        except Exception:
            continue
        value = _safe_call(attr, default=default)
        if value is not None:
            return value
    return default


def _to_datetime(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
    return None


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        for encoding in ("utf-8", "utf-16-le", "latin-1"):
            try:
                return value.decode(encoding, errors="replace")
            except Exception:
                continue
        return value.decode(errors="replace")
    return str(value)


class _HTMLToTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []
        self._ignore_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"style", "script", "head"}:
            self._ignore_depth += 1
            return
        if self._ignore_depth > 0:
            return
        if tag in {"br", "hr"}:
            self._parts.append("\n")
        elif tag in {"p", "div", "li", "tr"}:
            self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"style", "script", "head"}:
            self._ignore_depth = max(0, self._ignore_depth - 1)
            return
        if self._ignore_depth > 0:
            return
        if tag in {"p", "div", "li", "tr"}:
            self._parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self._ignore_depth > 0:
            return
        if data:
            self._parts.append(data)

    def get_text(self) -> str:
        text = unescape("".join(self._parts))
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        cleaned_lines: list[str] = []
        for line in text.split("\n"):
            stripped = line.strip()
            if not stripped:
                cleaned_lines.append("")
                continue
            if stripped.startswith("/*") or stripped.endswith("*/"):
                continue
            if "{" in stripped and "}" in stripped and ":" in stripped and len(stripped) > 25:
                continue
            cleaned_lines.append(stripped)
        text = "\n".join(cleaned_lines)
        text = text.replace("\u200c", " ").replace("\u200b", " ").replace("\xa0", " ")
        text = re.sub(r"[ \t]{2,}", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def _looks_like_html(text: str) -> bool:
    sample = text.lower()
    return "<html" in sample or "<body" in sample or "<div" in sample or "<p" in sample or "<span" in sample


def _html_to_text(text: str) -> str:
    parser = _HTMLToTextParser()
    try:
        parser.feed(text)
    except Exception:
        return unescape(re.sub(r"<[^>]+>", "", text))
    return parser.get_text()


class OstBackend(MailBackend):
    def __init__(self, ost_path: str | Path):
        self._ost_path = Path(ost_path).expanduser().resolve()
        if not self._ost_path.exists():
            raise BackendError(f"No se encontro el archivo OST: {self._ost_path}")

        try:
            import pypff  # type: ignore
        except ModuleNotFoundError as exc:
            raise BackendError(
                "No se pudo cargar pypff. Ejecuta Iniciar_Extractor.bat (usa Python 3.13 + libpff-python-windows) "
                "o instala manualmente libpff-python-windows en Python 3.13."
            ) from exc

        self._pypff = pypff
        self._file = pypff.file()
        self._file.open(str(self._ost_path))

        self._folder_map: dict[str, Any] = {}
        self._folder_children: dict[str, list[str]] = {}
        self._folder_paths: dict[str, str] = {}
        self._message_cache: dict[str, Any] = {}
        self._folders: list[FolderNode] = []
        self._build_folder_tree()

    @property
    def source_name(self) -> str:
        return str(self._ost_path)

    def close(self) -> None:
        try:
            self._file.close()
        except Exception:
            pass

    def list_folders(self) -> list[FolderNode]:
        return self._folders

    def list_messages(self, folder_id: str) -> list[MessageInfo]:
        return list(self._iter_messages_in_folder(folder_id))

    def _iter_messages_in_folder(self, folder_id: str) -> Iterable[MessageInfo]:
        folder = self._require_folder(folder_id)
        folder_path = self._folder_paths.get(folder_id, "")
        total = int(_pick(folder, ["number_of_sub_messages", "get_number_of_sub_messages"], default=0) or 0)

        for index in range(total):
            try:
                message = folder.get_sub_message(index)
            except Exception:
                continue

            message_class = str(_pick(message, ["message_class", "get_message_class"], default="") or "").lower()
            if message_class and "ipm.note" not in message_class:
                continue

            message_id = f"{folder_id}:{index}"
            self._message_cache[message_id] = message
            try:
                info = self._to_message_info(message_id, folder_id, folder_path, message)
            except Exception:
                # Algunos OST tienen items parcialmente corruptos; se omiten para no bloquear toda la carpeta.
                continue
            yield info

    def iter_messages(self, folder_id: str, recursive: bool = False) -> Iterable[MessageInfo]:
        folder_ids = [folder_id]
        if recursive:
            folder_ids = list(self._iter_descendants(folder_id))

        for current_folder_id in folder_ids:
            yield from self._iter_messages_in_folder(current_folder_id)

    def save_attachments(self, message: MessageInfo, output_dir: Path) -> list[Path]:
        raw_message = self._get_raw_message(message)
        if raw_message is None:
            return []

        output_dir.mkdir(parents=True, exist_ok=True)
        total_attachments = int(_pick(raw_message, ["number_of_attachments", "get_number_of_attachments"], default=0) or 0)
        saved: list[Path] = []

        for index in range(total_attachments):
            try:
                attachment = raw_message.get_attachment(index)
            except Exception:
                continue

            payload = self._attachment_payload(attachment)
            if payload is None:
                continue
            filename = self._attachment_name(attachment, index, payload=payload)
            destination = self._next_available_path(output_dir / filename)
            destination.write_bytes(payload)
            saved.append(destination)

        return saved

    def ensure_attachment_metadata(self, message: MessageInfo) -> None:
        if message.attachments_probed:
            return
        raw_message = self._get_raw_message(message)
        if raw_message is None:
            message.attachments_probed = True
            return
        attachment_count = self._safe_attachment_count(raw_message)
        attachments: list[AttachmentInfo] = []
        for index in range(attachment_count):
            try:
                attachment = raw_message.get_attachment(index)
            except Exception:
                continue

            name = self._attachment_name(attachment, index)
            size = _pick(attachment, ["size", "get_size"], default=None)
            try:
                if size is not None:
                    size = int(size)
            except (TypeError, ValueError):
                size = None

            attachments.append(
                AttachmentInfo(
                    id=f"{message.id}:{index}",
                    name=name,
                    size=size,
                    extension=Path(name).suffix.lower(),
                )
            )
        message.attachments = attachments
        message.attachment_count = max(message.attachment_count, len(attachments))
        message.has_attachments = message.attachment_count > 0
        message.attachments_probed = True

    def load_message_preview(self, message: MessageInfo) -> None:
        if message.preview_loaded:
            return
        raw_message = self._get_raw_message(message)
        if raw_message is None:
            message.preview_loaded = True
            return
        plain_candidate = _to_text(
            _pick(
                raw_message,
                ["plain_text_body", "get_plain_text_body", "body", "get_body"],
                default="",
            )
        )
        html_candidate = _to_text(_pick(raw_message, ["html_body", "get_html_body"], default=""))
        body_text = plain_candidate.strip()
        if not body_text and html_candidate:
            body_text = _html_to_text(html_candidate)
        elif body_text and _looks_like_html(body_text):
            body_text = _html_to_text(body_text)

        headers_text = _to_text(
            _pick(
                raw_message,
                ["transport_headers", "internet_headers", "get_transport_headers"],
                default="",
            )
        )
        message.body_text = body_text
        message.body_html = html_candidate
        message.headers_text = headers_text
        message.preview_loaded = True

    def _build_folder_tree(self) -> None:
        root = self._file.get_root_folder()
        count = int(_pick(root, ["number_of_sub_folders", "get_number_of_sub_folders"], default=0) or 0)
        self._folders = []
        for index in range(count):
            try:
                sub_folder = root.get_sub_folder(index)
            except Exception:
                continue
            folder_id = str(index)
            node = self._build_folder_node(
                folder=sub_folder,
                folder_id=folder_id,
                parent_id=None,
                parent_path="",
            )
            self._folders.append(node)

    def _build_folder_node(self, folder: Any, folder_id: str, parent_id: str | None, parent_path: str) -> FolderNode:
        folder_name = str(_pick(folder, ["name", "get_name"], default="(sin nombre)") or "(sin nombre)")
        folder_path = f"{parent_path}/{folder_name}" if parent_path else folder_name

        node = FolderNode(
            id=folder_id,
            name=folder_name,
            path=folder_path,
            parent_id=parent_id,
            children=[],
        )
        self._folder_map[folder_id] = folder
        self._folder_paths[folder_id] = folder_path

        child_ids: list[str] = []
        sub_count = int(_pick(folder, ["number_of_sub_folders", "get_number_of_sub_folders"], default=0) or 0)
        for child_index in range(sub_count):
            try:
                sub_folder = folder.get_sub_folder(child_index)
            except Exception:
                continue
            child_id = f"{folder_id}/{child_index}"
            child_ids.append(child_id)
            node.children.append(
                self._build_folder_node(
                    folder=sub_folder,
                    folder_id=child_id,
                    parent_id=folder_id,
                    parent_path=folder_path,
                )
            )
        self._folder_children[folder_id] = child_ids
        return node

    def _require_folder(self, folder_id: str) -> Any:
        folder = self._folder_map.get(folder_id)
        if folder is None:
            raise BackendError(f"Carpeta no encontrada: {folder_id}")
        return folder

    def _iter_descendants(self, folder_id: str) -> Iterable[str]:
        stack = [folder_id]
        while stack:
            current = stack.pop()
            yield current
            for child in reversed(self._folder_children.get(current, [])):
                stack.append(child)

    def _to_message_info(self, message_id: str, folder_id: str, folder_path: str, message: Any) -> MessageInfo:
        sender_name = _to_text(_pick(message, ["sender_name"], default="")).strip()
        sender_email = _to_text(_pick(message, ["sender_email_address"], default="")).strip()
        if sender_name and sender_email and sender_email.lower() not in sender_name.lower():
            sender = f"{sender_name} <{sender_email}>"
        else:
            sender = sender_name or sender_email

        if not sender:
            sender = _to_text(_pick(message, ["sent_representing_name"], default="")).strip()
        if not sender:
            sender = _to_text(_pick(message, ["transport_headers"], default="")).strip()
        subject = _pick(message, ["subject"], default="") or "(sin asunto)"
        received = _to_datetime(
            _pick(
                message,
                ["delivery_time", "client_submit_time", "creation_time"],
                default=None,
            )
        )
        raw_has_attachments = _pick(message, ["has_attachments", "get_has_attachments"], default=None)
        has_attachments = bool(raw_has_attachments) if raw_has_attachments is not None else False
        attachment_count = 0

        return MessageInfo(
            id=message_id,
            folder_id=folder_id,
            folder_path=folder_path,
            sender=str(sender or ""),
            subject=str(subject or "(sin asunto)"),
            received=received,
            has_attachments=has_attachments,
            attachment_count=attachment_count,
            attachments=[],
            attachments_probed=False,
            body_text="",
            body_html="",
            headers_text="",
            preview_loaded=False,
            backend_ref=message,
        )

    @staticmethod
    def _safe_attachment_count(message: Any) -> int:
        raw_count = _pick(message, ["number_of_attachments", "get_number_of_attachments"], default=0)
        try:
            return int(raw_count or 0)
        except Exception:
            return 0

    def _get_raw_message(self, message: MessageInfo) -> Any:
        if message.backend_ref is not None:
            return message.backend_ref

        if message.id in self._message_cache:
            return self._message_cache[message.id]

        folder_id, _, index_str = message.id.partition(":")
        folder = self._folder_map.get(folder_id)
        if folder is None:
            return None
        try:
            index = int(index_str)
        except ValueError:
            return None
        try:
            raw_message = folder.get_sub_message(index)
        except Exception:
            return None
        self._message_cache[message.id] = raw_message
        return raw_message

    @staticmethod
    def _attachment_name(attachment: Any, index: int, payload: bytes | None = None) -> str:
        name = _pick(
            attachment,
            [
                "long_filename",
                "filename",
                "name",
                "display_name",
            ],
            default=None,
        )
        extension_hint = OstBackend._guess_extension_from_payload(payload) if payload else None

        if name:
            cleaned_name = str(name).strip()
            if cleaned_name:
                current_suffix = Path(cleaned_name).suffix
                if extension_hint and (not current_suffix or current_suffix.lower() == ".bin"):
                    return f"{Path(cleaned_name).stem}{extension_hint}"
                return cleaned_name

        if extension_hint:
            return f"attachment_{index + 1:03d}{extension_hint}"
        return f"attachment_{index + 1:03d}.bin"

    @staticmethod
    def _guess_extension_from_payload(payload: bytes | None) -> str | None:
        if not payload:
            return None
        head = payload[:8192]
        head_lower = head.lower()

        if head.startswith(b"%PDF-"):
            return ".pdf"
        if head.startswith(b"\xFF\xD8\xFF"):
            return ".jpg"
        if head.startswith(b"\x89PNG\r\n\x1a\n"):
            return ".png"
        if head.startswith((b"GIF87a", b"GIF89a")):
            return ".gif"
        if head.startswith(b"BM"):
            return ".bmp"
        if head.startswith((b"II*\x00", b"MM\x00*")):
            return ".tif"
        if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            # OLE compound file (legacy Office, MSG, etc).
            return ".msg"
        if head.startswith(b"PK\x03\x04"):
            guessed = OstBackend._guess_ooxml_extension(payload)
            return guessed or ".zip"
        if head.startswith(b"Rar!\x1A\x07"):
            return ".rar"
        if head.startswith(b"7z\xBC\xAF\x27\x1C"):
            return ".7z"
        if head.startswith(b"\x1F\x8B"):
            return ".gz"
        if head.startswith(b"PK\x05\x06"):
            return ".zip"
        if head_lower.lstrip().startswith(b"{\\rtf"):
            return ".rtf"
        if b"<html" in head_lower or b"<!doctype html" in head_lower:
            return ".html"
        if head_lower.lstrip().startswith(b"<?xml"):
            return ".xml"
        if OstBackend._looks_textual(payload):
            return ".txt"
        return None

    @staticmethod
    def _guess_ooxml_extension(payload: bytes) -> str | None:
        try:
            with zipfile.ZipFile(BytesIO(payload)) as archive:
                names = [name.lower() for name in archive.namelist()[:200]]
        except Exception:
            return None

        if any(name.startswith("word/") for name in names):
            return ".docx"
        if any(name.startswith("xl/") for name in names):
            return ".xlsx"
        if any(name.startswith("ppt/") for name in names):
            return ".pptx"
        if any(name.startswith("visio/") for name in names):
            return ".vsdx"
        if "[content_types].xml" in names:
            return ".zip"
        return None

    @staticmethod
    def _looks_textual(payload: bytes) -> bool:
        sample = payload[:4096]
        if not sample:
            return False
        text_chars = sum(1 for b in sample if b in b"\t\n\r" or 32 <= b <= 126)
        ratio = text_chars / max(1, len(sample))
        return ratio > 0.9

    @staticmethod
    def _attachment_payload(attachment: Any) -> bytes | None:
        payload = _pick(attachment, ["data", "get_data"], default=None)
        if payload is not None:
            try:
                return bytes(payload)
            except Exception:
                pass

        read_buffer = getattr(attachment, "read_buffer", None)
        if callable(read_buffer):
            size = _pick(attachment, ["size", "get_size"], default=None)
            try:
                size_int = int(size) if size is not None else None
            except (TypeError, ValueError):
                size_int = None

            chunks: list[bytes] = []
            if size_int is None or size_int <= 0:
                size_int = 4 * 1024 * 1024

            remaining = size_int
            while remaining > 0:
                chunk_size = min(1024 * 1024, remaining)
                try:
                    chunk = read_buffer(chunk_size)
                except Exception:
                    break
                if not chunk:
                    break
                chunk_bytes = bytes(chunk)
                chunks.append(chunk_bytes)
                remaining -= len(chunk_bytes)
                if len(chunk_bytes) < chunk_size:
                    break
            if chunks:
                return b"".join(chunks)
        return None

    @staticmethod
    def _next_available_path(path: Path) -> Path:
        if not path.exists():
            return path
        stem = path.stem
        suffix = path.suffix
        parent = path.parent
        index = 1
        while True:
            candidate = parent / f"{stem}_{index}{suffix}"
            if not candidate.exists():
                return candidate
            index += 1
