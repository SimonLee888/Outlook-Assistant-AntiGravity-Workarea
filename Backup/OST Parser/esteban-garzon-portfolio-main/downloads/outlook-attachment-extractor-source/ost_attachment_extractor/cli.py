from __future__ import annotations

import argparse
from datetime import datetime, time
from pathlib import Path
from typing import Iterable

from .backends.base import BackendError
from .backends.ost import OstBackend
from .extractor import extract_attachments
from .models import ExtractionFilters, FolderNode


def flatten_folders(folders: Iterable[FolderNode]) -> list[FolderNode]:
    output: list[FolderNode] = []

    def walk(node: FolderNode) -> None:
        output.append(node)
        for child in node.children:
            walk(child)

    for folder in folders:
        walk(folder)
    return output


def parse_date(raw: str | None, end_of_day: bool) -> datetime | None:
    if not raw:
        return None
    parsed = datetime.strptime(raw, "%Y-%m-%d")
    if end_of_day:
        return datetime.combine(parsed.date(), time(23, 59, 59))
    return datetime.combine(parsed.date(), time(0, 0, 0))


def find_folder_id(backend: OstBackend, folder_path: str) -> str:
    normalized = folder_path.strip().lower().replace("\\", "/")
    all_folders = flatten_folders(backend.list_folders())
    for folder in all_folders:
        if folder.path.lower().replace("\\", "/") == normalized:
            return folder.id
    raise BackendError(f"No se encontro la carpeta: {folder_path}")


def cmd_list_folders(args: argparse.Namespace) -> int:
    backend = OstBackend(args.source)
    try:
        for folder in flatten_folders(backend.list_folders()):
            print(f"{folder.id}\t{folder.path}")
    finally:
        backend.close()
    return 0


def cmd_extract(args: argparse.Namespace) -> int:
    backend = OstBackend(args.source)
    try:
        folder_id = args.folder_id
        if not folder_id:
            if not args.folder_path:
                raise BackendError("Debes indicar --folder-id o --folder-path para extraer.")
            folder_id = find_folder_id(backend, args.folder_path)

        filters = ExtractionFilters(
            sender_contains=args.sender_contains or "",
            subject_contains=args.subject_contains or "",
            attachment_name_contains=args.attachment_name_contains or "",
            extension=args.extension or "",
            has_attachments_only=not args.include_no_attachments,
            date_from=parse_date(args.date_from, end_of_day=False),
            date_to=parse_date(args.date_to, end_of_day=True),
        )
        result = extract_attachments(
            backend=backend,
            folder_id=folder_id,
            output_dir=Path(args.output),
            filters=filters,
            recursive=args.recursive,
            progress_callback=lambda text: print(text),
        )
        print(
            "Extraccion completada | "
            f"escaneados={result.scanned_messages} "
            f"coinciden={result.matched_messages} "
            f"adjuntos={result.exported_attachments}"
        )
        if result.manifest_path:
            print(f"Manifest: {result.manifest_path}")
    finally:
        backend.close()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ost-extractor",
        description="Extrae adjuntos desde archivos OST de Outlook con filtros.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    parser_list = subparsers.add_parser("list-folders", help="Lista carpetas del OST")
    parser_list.add_argument("--source", required=True, help="Ruta del archivo .ost")
    parser_list.set_defaults(func=cmd_list_folders)

    parser_extract = subparsers.add_parser("extract", help="Extrae adjuntos del OST")
    parser_extract.add_argument("--source", required=True, help="Ruta del archivo .ost")
    parser_extract.add_argument("--output", required=True, help="Directorio de salida")
    parser_extract.add_argument("--folder-id", help="ID de carpeta (ver list-folders)")
    parser_extract.add_argument("--folder-path", help="Ruta de carpeta legible (ej. Inbox/Facturas)")
    parser_extract.add_argument("--recursive", action="store_true", help="Incluir subcarpetas")
    parser_extract.add_argument("--sender-contains", help="Filtro por remitente")
    parser_extract.add_argument("--subject-contains", help="Filtro por asunto")
    parser_extract.add_argument("--attachment-name-contains", help="Filtro por nombre de adjunto")
    parser_extract.add_argument("--extension", help="Filtro por extension (ej: .pdf)")
    parser_extract.add_argument("--date-from", help="Fecha inicial YYYY-MM-DD")
    parser_extract.add_argument("--date-to", help="Fecha final YYYY-MM-DD")
    parser_extract.add_argument(
        "--include-no-attachments",
        action="store_true",
        help="No exigir adjuntos en el filtro (normalmente solo procesa correos con adjuntos)",
    )
    parser_extract.set_defaults(func=cmd_extract)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except BackendError as exc:
        print(f"Error de backend: {exc}")
        return 1
    except ValueError as exc:
        print(f"Error de validacion: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

