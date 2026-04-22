from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime, time
from html import escape as html_escape
from pathlib import Path
import re
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from .backends.base import BackendError, MailBackend
from .backends.ost import OstBackend
from .extractor import build_message_output_dir, extract_attachments
from .models import ExtractionFilters, FolderNode, MessageInfo

try:
    from tkinterweb import HtmlFrame  # type: ignore
except Exception:
    HtmlFrame = None


class AttachmentExtractorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook Attachment Extractor")
        self.geometry("1320x780")
        self.minsize(1140, 680)

        self.backend: MailBackend | None = None
        self.folder_lookup: dict[str, FolderNode] = {}
        self.current_messages: list[MessageInfo] = []
        self.message_lookup: dict[str, MessageInfo] = {}
        self.last_output_dir: Path | None = None
        self._busy_count = 0
        self.preview_html_frame = None
        self.preview_text_fallback = None

        self.source_label_var = tk.StringVar(value="Sin fuente")
        self.status_var = tk.StringVar(value="Listo")
        self.recursive_var = tk.BooleanVar(value=True)
        self.sender_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.attachment_name_var = tk.StringVar()
        self.extension_var = tk.StringVar()
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.has_attachments_var = tk.BooleanVar(value=False)
        self.preview_from_var = tk.StringVar(value="")
        self.preview_subject_var = tk.StringVar(value="")
        self.preview_received_var = tk.StringVar(value="")
        self.preview_att_var = tk.StringVar(value="")

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(toolbar, text="Add Source (.ost)", command=self._on_add_source).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Extract", command=self._on_extract).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Checkbutton(toolbar, text="Recursive", variable=self.recursive_var).pack(side=tk.LEFT, padx=(8, 12))

        self.progress = ttk.Progressbar(toolbar, mode="indeterminate", length=180)
        self.progress.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(toolbar, textvariable=self.status_var).pack(side=tk.LEFT)
        ttk.Label(toolbar, textvariable=self.source_label_var).pack(side=tk.RIGHT)

        root_pane = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        root_pane.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        left_frame = ttk.Frame(root_pane, padding=(0, 0, 4, 0))
        right_frame = ttk.Frame(root_pane, padding=(4, 0, 0, 0))
        root_pane.add(left_frame, weight=1)
        root_pane.add(right_frame, weight=3)

        ttk.Label(left_frame, text="Folders").pack(anchor=tk.W)
        self.folder_tree = ttk.Treeview(left_frame, columns=("path",), show="tree")
        self.folder_tree.pack(fill=tk.BOTH, expand=True)
        self.folder_tree.bind("<<TreeviewSelect>>", self._on_folder_selected)

        filter_card = ttk.LabelFrame(right_frame, text="Filters", padding=8)
        filter_card.pack(side=tk.TOP, fill=tk.X)
        self._build_filter_controls(filter_card)

        message_card = ttk.Frame(right_frame)
        message_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(6, 0))
        self._build_message_and_preview(message_card)

        log_card = ttk.LabelFrame(right_frame, text="Log", padding=8)
        log_card.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        self.log_text = tk.Text(log_card, height=4, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.X)

    def _build_filter_controls(self, parent: ttk.LabelFrame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)

        ttk.Label(parent, text="From contains").grid(row=0, column=0, sticky=tk.W, padx=(0, 6), pady=2)
        ttk.Entry(parent, textvariable=self.sender_var).grid(row=0, column=1, sticky=tk.EW, pady=2)
        ttk.Label(parent, text="Subject contains").grid(row=0, column=2, sticky=tk.W, padx=(10, 6), pady=2)
        ttk.Entry(parent, textvariable=self.subject_var).grid(row=0, column=3, sticky=tk.EW, pady=2)

        ttk.Label(parent, text="Attachment name").grid(row=1, column=0, sticky=tk.W, padx=(0, 6), pady=2)
        ttk.Entry(parent, textvariable=self.attachment_name_var).grid(row=1, column=1, sticky=tk.EW, pady=2)
        ttk.Label(parent, text="Extension").grid(row=1, column=2, sticky=tk.W, padx=(10, 6), pady=2)
        ttk.Entry(parent, textvariable=self.extension_var).grid(row=1, column=3, sticky=tk.EW, pady=2)

        ttk.Label(parent, text="Date from (YYYY-MM-DD)").grid(row=2, column=0, sticky=tk.W, padx=(0, 6), pady=2)
        ttk.Entry(parent, textvariable=self.date_from_var).grid(row=2, column=1, sticky=tk.EW, pady=2)
        ttk.Label(parent, text="Date to (YYYY-MM-DD)").grid(row=2, column=2, sticky=tk.W, padx=(10, 6), pady=2)
        ttk.Entry(parent, textvariable=self.date_to_var).grid(row=2, column=3, sticky=tk.EW, pady=2)

        ttk.Checkbutton(parent, text="Only messages with attachments", variable=self.has_attachments_var).grid(
            row=3, column=0, sticky=tk.W, pady=(4, 2)
        )
        button_bar = ttk.Frame(parent)
        button_bar.grid(row=3, column=3, sticky=tk.E, pady=(4, 2))
        ttk.Button(button_bar, text="Clear", command=self._clear_filters).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(button_bar, text="Apply Filters", command=self._apply_filters_async).pack(side=tk.RIGHT)

    def _build_message_and_preview(self, parent: ttk.Frame) -> None:
        pane = ttk.Panedwindow(parent, orient=tk.VERTICAL)
        pane.pack(fill=tk.BOTH, expand=True)

        table_frame = ttk.Frame(pane)
        preview_frame = ttk.LabelFrame(pane, text="Mail Preview", padding=8)
        pane.add(table_frame, weight=3)
        pane.add(preview_frame, weight=2)

        self._build_message_table(table_frame)
        self._build_preview_panel(preview_frame)

    def _build_message_table(self, parent: ttk.Frame) -> None:
        columns = ("from", "subject", "received", "attachments")
        self.message_table = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        self.message_table.heading("from", text="FROM")
        self.message_table.heading("subject", text="SUBJECT")
        self.message_table.heading("received", text="RECEIVED")
        self.message_table.heading("attachments", text="ATT")
        self.message_table.column("from", width=300, anchor=tk.W)
        self.message_table.column("subject", width=500, anchor=tk.W)
        self.message_table.column("received", width=180, anchor=tk.W)
        self.message_table.column("attachments", width=70, anchor=tk.CENTER)
        self.message_table.pack(fill=tk.BOTH, expand=True)
        self.message_table.bind("<<TreeviewSelect>>", self._on_message_selected)
        self.message_table.bind("<Button-3>", self._on_message_right_click)

        self.message_menu = tk.Menu(self, tearoff=0)
        self.message_menu.add_command(label="Extract Attachments", command=self._extract_selected_message_attachments)

    def _build_preview_panel(self, parent: ttk.LabelFrame) -> None:
        meta = ttk.Frame(parent)
        meta.pack(fill=tk.X)
        meta.columnconfigure(1, weight=1)

        ttk.Label(meta, text="From:").grid(row=0, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Label(meta, textvariable=self.preview_from_var).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(meta, text="Subject:").grid(row=1, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Label(meta, textvariable=self.preview_subject_var).grid(row=1, column=1, sticky=tk.W)
        ttk.Label(meta, text="Received:").grid(row=2, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Label(meta, textvariable=self.preview_received_var).grid(row=2, column=1, sticky=tk.W)
        ttk.Label(meta, text="Attachments:").grid(row=3, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Label(meta, textvariable=self.preview_att_var).grid(row=3, column=1, sticky=tk.W)

        if HtmlFrame is not None:
            self.preview_html_frame = HtmlFrame(parent, horizontal_scrollbar="auto")
            self.preview_html_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        else:
            self.preview_text_fallback = tk.Text(parent, wrap=tk.WORD, height=12, state=tk.DISABLED)
            self.preview_text_fallback.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

    def _on_add_source(self) -> None:
        ost_file = filedialog.askopenfilename(
            title="Select OST file",
            filetypes=[("Outlook OST", "*.ost"), ("All files", "*.*")],
        )
        if not ost_file:
            return

        self._run_worker(
            label=f"Opening OST: {ost_file}",
            task=lambda: OstBackend(ost_file),
            on_success=self._set_backend,
        )

    def _set_backend(self, backend: MailBackend) -> None:
        if self.backend:
            self.backend.close()
        self.backend = backend
        self.source_label_var.set(f"Source: {backend.source_name}")
        self._log(f"Fuente cargada: {backend.source_name}")
        self._populate_folder_tree(backend.list_folders())
        self._clear_messages()
        self._auto_select_initial_folder()

    def _populate_folder_tree(self, root_folders: list[FolderNode]) -> None:
        self.folder_lookup.clear()
        self.folder_tree.delete(*self.folder_tree.get_children())

        def add_node(parent_iid: str, node: FolderNode) -> None:
            self.folder_lookup[node.id] = node
            iid = self.folder_tree.insert(parent_iid, tk.END, iid=node.id, text=node.name, open=False)
            for child in node.children:
                add_node(iid, child)

        for root in root_folders:
            add_node("", root)

    def _auto_select_initial_folder(self) -> None:
        inbox_id: str | None = None
        for folder_id, folder in self.folder_lookup.items():
            path_lower = folder.path.lower().replace("\\", "/")
            if path_lower.endswith("/inbox"):
                inbox_id = folder_id
                break
        target_id = inbox_id
        if target_id is None:
            roots = self.folder_tree.get_children("")
            if not roots:
                return
            target_id = roots[0]
        self.folder_tree.selection_set(target_id)
        self.folder_tree.focus(target_id)
        self.folder_tree.see(target_id)
        self._on_folder_selected()

    def _on_folder_selected(self, _event: tk.Event | None = None) -> None:
        backend = self.backend
        if backend is None:
            return

        selected = self.folder_tree.selection()
        if not selected:
            return
        folder_id = selected[0]
        folder_name = self.folder_lookup.get(folder_id).path if folder_id in self.folder_lookup else folder_id
        recursive = self.recursive_var.get()

        self._run_worker(
            label=f"Loading messages from {folder_name} (recursive={recursive})",
            task=lambda: self._collect_messages(backend, folder_id, recursive),
            on_success=self._set_current_messages,
        )

    def _collect_messages(self, backend: MailBackend, folder_id: str, recursive: bool) -> list[MessageInfo]:
        messages: list[MessageInfo] = []
        for idx, message in enumerate(backend.iter_messages(folder_id, recursive=recursive), start=1):
            messages.append(message)
            if idx % 250 == 0:
                self.after(0, lambda n=idx: self.status_var.set(f"Cargando mensajes... {n}"))
        return messages

    def _set_current_messages(self, messages: list[MessageInfo]) -> None:
        self.current_messages = messages
        self._log(f"Mensajes cargados: {len(messages)}")
        self._apply_filters_async()

    def _apply_filters_async(self) -> None:
        backend = self.backend
        all_messages = list(self.current_messages)
        if not all_messages:
            self._render_messages([])
            self.status_var.set("No hay mensajes cargados")
            return

        filters = self._build_filters()
        requires_attachment_details = bool(filters.extension or filters.attachment_name_contains or filters.has_attachments_only)
        requires_sender_chain_scan = bool(filters.sender_contains and backend is not None)

        def task() -> list[MessageInfo]:
            thread_has_attachment: dict[str, bool] = {}
            if filters.has_attachments_only:
                total_seed = len(all_messages)
                for idx, message in enumerate(all_messages, start=1):
                    if requires_sender_chain_scan and backend is not None:
                        sender_probe = (message.sender or "").lower()
                        query = filters.sender_contains.lower()
                        if query not in sender_probe:
                            backend.load_message_preview(message)
                    if backend is not None:
                        try:
                            backend.ensure_attachment_metadata(message)
                        except Exception:
                            pass
                    subject_key = self._normalize_subject(message.subject)
                    if message.has_attachments:
                        thread_has_attachment[subject_key] = True
                    if idx % 1000 == 0:
                        self.after(0, lambda i=idx, t=total_seed: self.status_var.set(f"Indexando hilo {i}/{t}..."))

            matched: list[MessageInfo] = []
            total = len(all_messages)
            for idx, message in enumerate(all_messages, start=1):
                if requires_sender_chain_scan and backend is not None and not filters.has_attachments_only:
                    sender_probe = (message.sender or "").lower()
                    query = filters.sender_contains.lower()
                    if query not in sender_probe:
                        backend.load_message_preview(message)

                if requires_attachment_details and backend is not None and not message.attachments_probed:
                    try:
                        backend.ensure_attachment_metadata(message)
                    except Exception:
                        pass

                thread_key = self._normalize_subject(message.subject)
                message.thread_has_attachments = thread_has_attachment.get(thread_key, False)

                if filters.has_attachments_only and not message.has_attachments and message.thread_has_attachments:
                    original = message.has_attachments
                    message.has_attachments = True
                    is_match = filters.matches(message)
                    message.has_attachments = original
                else:
                    is_match = filters.matches(message)

                if is_match:
                    matched.append(message)
                if idx % 1000 == 0:
                    self.after(0, lambda i=idx, t=total: self.status_var.set(f"Filtrando {i}/{t}..."))
            return matched

        self._run_worker(
            label=f"Applying filters over {len(all_messages)} message(s)...",
            task=task,
            on_success=self._render_messages,
            show_error=False,
        )

    def _render_messages(self, filtered_messages: list[MessageInfo]) -> None:
        self.message_lookup.clear()
        self.message_table.delete(*self.message_table.get_children())

        for message in filtered_messages:
            received = message.received.strftime("%Y-%m-%d %H:%M:%S") if message.received else ""
            if message.attachment_count > 0:
                att_value = str(message.attachment_count)
            elif message.has_attachments:
                att_value = "Y"
            else:
                att_value = ""
            self.message_lookup[message.id] = message
            self.message_table.insert(
                "",
                tk.END,
                iid=message.id,
                values=(message.sender, message.subject, received, att_value),
            )

        self._clear_preview()
        self._log(f"Mostrando {len(filtered_messages)} mensaje(s)")
        self.status_var.set(f"Cargados: {len(self.current_messages)} | Mostrados: {len(filtered_messages)}")

    def _clear_filters(self) -> None:
        self.sender_var.set("")
        self.subject_var.set("")
        self.attachment_name_var.set("")
        self.extension_var.set("")
        self.date_from_var.set("")
        self.date_to_var.set("")
        self.has_attachments_var.set(False)
        self._apply_filters_async()

    def _on_extract(self) -> None:
        backend = self.backend
        if backend is None:
            messagebox.showwarning("No source", "Primero debes cargar un archivo OST.")
            return

        selected = self.folder_tree.selection()
        if not selected:
            messagebox.showwarning("No folder", "Selecciona una carpeta para extraer adjuntos.")
            return
        folder_id = selected[0]

        output_dir = filedialog.askdirectory(title="Select output folder")
        if not output_dir:
            return

        self.last_output_dir = Path(output_dir)
        filters = self._build_filters()
        recursive = self.recursive_var.get()
        destination = Path(output_dir)

        def run_extraction():
            return extract_attachments(
                backend=backend,
                folder_id=folder_id,
                output_dir=destination,
                filters=filters,
                recursive=recursive,
                progress_callback=lambda text: self.after(0, lambda: self._log(text)),
            )

        self._run_worker(
            label="Extracting attachments...",
            task=run_extraction,
            on_success=self._on_extraction_done,
        )

    def _on_extraction_done(self, result) -> None:
        self._log(
            f"Extraccion completa. Escaneados: {result.scanned_messages}, coinciden: {result.matched_messages}, "
            f"adjuntos guardados: {result.exported_attachments}"
        )
        manifest = str(result.manifest_path) if result.manifest_path else "(no manifest)"
        messagebox.showinfo(
            "Extraction completed",
            "Proceso finalizado\n\n"
            f"Mensajes escaneados: {result.scanned_messages}\n"
            f"Mensajes coincidentes: {result.matched_messages}\n"
            f"Adjuntos exportados: {result.exported_attachments}\n"
            f"Manifest: {manifest}",
        )

    def _on_message_selected(self, _event: tk.Event | None = None) -> None:
        message = self._get_selected_message()
        if message is None:
            self._clear_preview()
            return

        received = message.received.strftime("%Y-%m-%d %H:%M:%S") if message.received else ""

        self.preview_from_var.set(message.sender or "")
        self.preview_subject_var.set(message.subject or "")
        self.preview_received_var.set(received)
        self.preview_att_var.set(self._attachment_text_for_preview(message))

        if not message.attachments_probed:
            backend = self.backend
            if backend is not None:
                self.preview_att_var.set("Cargando adjuntos...")
                message_id = message.id
                self._run_worker(
                    label=f"Loading attachment metadata for {message_id}",
                    task=lambda: backend.ensure_attachment_metadata(message),
                    on_success=lambda _result: self._on_attachment_metadata_loaded(message_id),
                    show_error=False,
                    use_busy=False,
                    log=False,
                )

        if message.preview_loaded:
            self._render_preview_body(message)
            return

        self._set_preview_text("Cargando preview...")
        backend = self.backend
        if backend is None:
            return
        message_id = message.id
        self._run_worker(
            label=f"Loading preview for message {message_id}",
            task=lambda: backend.load_message_preview(message),
            on_success=lambda _result: self._on_preview_loaded(message_id),
            show_error=False,
            use_busy=False,
            log=False,
        )

    def _on_preview_loaded(self, message_id: str) -> None:
        selected = self._get_selected_message()
        if selected is None or selected.id != message_id:
            return
        self._render_preview_body(selected)

    def _on_attachment_metadata_loaded(self, message_id: str) -> None:
        selected = self._get_selected_message()
        if selected is None or selected.id != message_id:
            return
        self.preview_att_var.set(self._attachment_text_for_preview(selected))

    def _attachment_text_for_preview(self, message: MessageInfo) -> str:
        if message.attachments:
            text = ", ".join(att.name for att in message.attachments[:5])
            if len(message.attachments) > 5:
                text += f" ... (+{len(message.attachments) - 5})"
            return text
        if message.thread_has_attachments:
            return "Adjuntos en la cadena del hilo"
        if not message.attachments_probed:
            return "Detectando adjuntos..."
        if message.has_attachments:
            return "Tiene adjuntos"
        return "(sin adjuntos)"

    def _render_preview_body(self, message: MessageInfo) -> None:
        if self.preview_html_frame is not None and message.body_html:
            self._set_preview_html(message.body_html)
            return

        body = (message.body_text or "").strip()
        if not body:
            body = "No hay cuerpo de mensaje disponible para preview."
        elif len(body) > 60000:
            body = body[:60000] + "\n\n[preview truncado]"
        self._set_preview_text(body)

    def _on_message_right_click(self, event: tk.Event) -> None:
        row_id = self.message_table.identify_row(event.y)
        if not row_id:
            return
        self.message_table.selection_set(row_id)
        self.message_table.focus(row_id)
        self._on_message_selected()
        try:
            self.message_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.message_menu.grab_release()

    def _extract_selected_message_attachments(self) -> None:
        backend = self.backend
        if backend is None:
            messagebox.showwarning("No source", "Primero debes cargar un archivo OST.")
            return

        message = self._get_selected_message()
        if message is None:
            messagebox.showwarning("No message", "Selecciona un correo.")
            return
        if not message.attachments_probed:
            backend.ensure_attachment_metadata(message)
        if not message.has_attachments and not message.attachments:
            backend.ensure_attachment_metadata(message)
            if not message.has_attachments and not message.attachments:
                if message.thread_has_attachments:
                    extract_thread = messagebox.askyesno(
                        "Thread attachments",
                        "Este correo no tiene adjuntos directos, pero la cadena del hilo si.\n"
                        "Quieres extraer adjuntos de toda la cadena?",
                    )
                    if extract_thread:
                        self._extract_thread_attachments(message, backend)
                        return
                messagebox.showinfo("No attachments", "El correo seleccionado no tiene adjuntos.")
                return

        initial_dir = str(self.last_output_dir) if self.last_output_dir else str(Path.cwd())
        output_dir = filedialog.askdirectory(title="Select output folder", initialdir=initial_dir)
        if not output_dir:
            return
        self.last_output_dir = Path(output_dir)

        destination = build_message_output_dir(Path(output_dir), message)
        self._run_worker(
            label=f"Extracting attachments from selected message: {message.subject}",
            task=lambda: backend.save_attachments(message, destination),
            on_success=lambda saved: self._on_single_extraction_done(message, saved, destination),
        )

    def _extract_thread_attachments(self, anchor_message: MessageInfo, backend: MailBackend) -> None:
        initial_dir = str(self.last_output_dir) if self.last_output_dir else str(Path.cwd())
        output_dir = filedialog.askdirectory(title="Select output folder", initialdir=initial_dir)
        if not output_dir:
            return
        self.last_output_dir = Path(output_dir)
        output_base = Path(output_dir)
        thread_key = self._normalize_subject(anchor_message.subject)

        def task() -> tuple[list[Path], int]:
            saved_files: list[Path] = []
            processed = 0
            for message in self.current_messages:
                if self._normalize_subject(message.subject) != thread_key:
                    continue
                processed += 1
                backend.ensure_attachment_metadata(message)
                if not message.has_attachments and not message.attachments:
                    continue
                destination = build_message_output_dir(output_base, message)
                saved_files.extend(backend.save_attachments(message, destination))
            return saved_files, processed

        self._run_worker(
            label=f"Extracting thread attachments: {anchor_message.subject}",
            task=task,
            on_success=lambda payload: self._on_thread_extraction_done(anchor_message, payload),
        )

    def _on_thread_extraction_done(self, anchor_message: MessageInfo, payload: object) -> None:
        files: list[Path] = []
        processed = 0
        if isinstance(payload, tuple) and len(payload) == 2:
            files = payload[0] if isinstance(payload[0], list) else []
            processed = int(payload[1])
        count = len(files)
        self._log(f"Extraccion de cadena completada. Mensajes del hilo: {processed}, archivos: {count}")
        messagebox.showinfo(
            "Thread extraction completed",
            f"Hilo: {anchor_message.subject}\nMensajes en cadena: {processed}\nAdjuntos exportados: {count}",
        )

    def _on_single_extraction_done(self, message: MessageInfo, saved_files: object, destination: Path) -> None:
        files = saved_files if isinstance(saved_files, list) else []
        count = len(files)
        if count == 0:
            self._log("No se pudo guardar ningun adjunto del correo seleccionado.")
            messagebox.showwarning("Extraction", "No se pudieron exportar adjuntos de ese correo.")
            return
        self._log(f"Correo extraido: {count} adjunto(s) -> {destination}")
        messagebox.showinfo(
            "Extraction completed",
            f"Correo: {message.subject}\nAdjuntos exportados: {count}\nDestino: {destination}",
        )

    def _get_selected_message(self) -> MessageInfo | None:
        selected = self.message_table.selection()
        if not selected:
            return None
        return self.message_lookup.get(selected[0])

    def _build_filters(self) -> ExtractionFilters:
        date_from = self._parse_date(self.date_from_var.get().strip(), end_of_day=False)
        date_to = self._parse_date(self.date_to_var.get().strip(), end_of_day=True)
        return ExtractionFilters(
            sender_contains=self.sender_var.get().strip(),
            subject_contains=self.subject_var.get().strip(),
            attachment_name_contains=self.attachment_name_var.get().strip(),
            extension=self.extension_var.get().strip(),
            has_attachments_only=self.has_attachments_var.get(),
            date_from=date_from,
            date_to=date_to,
        )

    def _parse_date(self, text: str, end_of_day: bool) -> datetime | None:
        if not text:
            return None
        try:
            parsed = datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid date", f"Formato de fecha invalido: {text}. Usa YYYY-MM-DD.")
            return None
        if end_of_day:
            return datetime.combine(parsed.date(), time(23, 59, 59))
        return datetime.combine(parsed.date(), time(0, 0, 0))

    def _run_worker(
        self,
        label: str,
        task: Callable[[], object],
        on_success: Callable[[object], None],
        show_error: bool = True,
        use_busy: bool = True,
        log: bool = True,
    ) -> None:
        if log:
            self._log(label)
        if use_busy:
            self._set_busy(True, label)

        def worker() -> None:
            try:
                result = task()
            except BackendError as exc:
                def handle_backend_error() -> None:
                    if show_error:
                        messagebox.showerror("Backend error", str(exc))
                    self._log(f"Error: {exc}")
                    if use_busy:
                        self._set_busy(False, "Listo")

                self.after(0, handle_backend_error)
                return
            except Exception as exc:
                def handle_error() -> None:
                    if show_error:
                        messagebox.showerror("Unexpected error", str(exc))
                    self._log(f"Error inesperado: {exc}")
                    if use_busy:
                        self._set_busy(False, "Listo")

                self.after(0, handle_error)
                return

            def finish() -> None:
                on_success(result)
                if use_busy:
                    self._set_busy(False, "Listo")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _set_busy(self, busy: bool, text: str) -> None:
        if busy:
            self._busy_count += 1
            if self._busy_count == 1:
                self.progress.start(10)
        else:
            self._busy_count = max(0, self._busy_count - 1)
            if self._busy_count == 0:
                self.progress.stop()
        self.status_var.set(text)

    @staticmethod
    def _normalize_subject(subject: str) -> str:
        base = (subject or "").strip().lower()
        base = re.sub(r"^\s*((re|fw|fwd)\s*:\s*)+", "", base, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", base).strip()

    def _clear_messages(self) -> None:
        self.current_messages = []
        self.message_lookup.clear()
        self.message_table.delete(*self.message_table.get_children())
        self._clear_preview()
        self.status_var.set("Sin mensajes")

    def _clear_preview(self) -> None:
        self.preview_from_var.set("")
        self.preview_subject_var.set("")
        self.preview_received_var.set("")
        self.preview_att_var.set("")
        self._set_preview_text("")

    def _set_preview_text(self, text: str) -> None:
        if self.preview_html_frame is not None:
            html = f"<html><body><pre style='font-family:Calibri,Arial,sans-serif;white-space:pre-wrap'>{html_escape(text)}</pre></body></html>"
            self.preview_html_frame.load_html(html)
            return
        if self.preview_text_fallback is None:
            return
        self.preview_text_fallback.configure(state=tk.NORMAL)
        self.preview_text_fallback.delete("1.0", tk.END)
        self.preview_text_fallback.insert("1.0", text)
        self.preview_text_fallback.configure(state=tk.DISABLED)

    def _set_preview_html(self, html: str) -> None:
        if self.preview_html_frame is None:
            self._set_preview_text(html)
            return
        self.preview_html_frame.load_html(html)

    def _log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{stamp}] {text}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _on_close(self) -> None:
        if self.backend:
            self.backend.close()
        self.destroy()


def run_gui() -> None:
    app = AttachmentExtractorApp()
    app.mainloop()
