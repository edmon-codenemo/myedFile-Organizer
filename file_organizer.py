"""
File Organizer — Fase 2B (Vigilancia inteligente + Tray)
"""

import os
import re
import sys
import time
import shutil
import sqlite3
import argparse
import unittest
import tempfile
import threading
import hashlib
import unicodedata
from pathlib import Path
from datetime import datetime
from tkinter import ttk, filedialog, messagebox

import customtkinter as ctk
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw

try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
APP_NAME = "File Organizer"
APP_DATA_DIR = Path(os.path.expanduser("~")) / ".file_organizer"
APP_DATA_DIR.mkdir(exist_ok=True)
DB_PATH = APP_DATA_DIR / "history.db"
DEFAULT_DEST_NAME = "_Organizado"
DUP_FOLDER_NAME = "_Duplicados"

TEXTS = {
    "es": {
        "title": "Organizador de Archivos",
        "tab_organize": "Organizar", "tab_dups": "Duplicados",
        "tab_watch": "Vigilancia", "tab_history": "Historial",
        "src_folder": "Carpeta origen:", "dest_folder": "Carpeta destino:",
        "browse": "Elegir...", "opt_type": "Por tipo", "opt_date": "Por fecha",
        "opt_content": "Por contenido", "opt_name": "Por nombre",
        "content_rules": "Reglas (ej: artemisa -> artemisa):",
        "btn_preview": "Vista previa", "btn_execute": "Ejecutar", "btn_cancel": "Cancelar",
        "btn_scan_dups": "Buscar duplicados", "btn_move_dups": "Mover duplicados",
        "btn_start_watch": "▶ Iniciar Vigilancia", "btn_stop_watch": "⏹ Detener Vigilancia",
        "status_ready": "Listo.", "status_scanning": "Escaneando...",
        "status_reading": "Leyendo: {file}", "status_preview_ready": "{n} archivos listos.",
        "status_executing": "Moviendo... {cur}/{total}", "status_done": "✓ {n} movidos.",
        "status_undoing": "Deshaciendo...", "status_undone": "✓ Deshecho.",
        "status_scan_dups": "Buscando duplicados...", "status_dups_found": "{n} grupos encontrados.",
        "status_moving_dups": "Moviendo duplicados...", "status_dups_moved": "✓ {n} movidos.",
        "status_watch_active": "👁 Vigilando: {folder}",
        "status_watch_inactive": "Vigilancia detenida.",
        "col_file": "Archivo", "col_dest": "Destino", "col_rule": "Regla",
        "col_date": "Fecha", "col_files": "Archivos", "col_status": "Estado",
        "col_dup_size": "Tamaño", "col_dup_files": "Archivos duplicados",
        "btn_undo": "Deshacer", "status_active": "Activa", "status_undone_label": "Deshecha",
        "confirm_execute": "¿Mover {n} archivos?", "confirm_dups": "¿Mover {n} duplicados?",
        "no_moves": "No hay movimientos.", "no_dups": "No se encontraron duplicados.",
        "ask_folder": "Selecciona una carpeta.", "errors_title": "Errores",
        "watch_msg_title": "Vigilancia", "watch_msg_body": "El programa se minimizará en la bandeja.\nLos archivos nuevos se organizarán automáticamente.",
    },
    "en": {
        "title": "File Organizer", "tab_organize": "Organize", "tab_dups": "Duplicates",
        "tab_watch": "Watch", "tab_history": "History",
        "src_folder": "Source folder:", "dest_folder": "Destination folder:",
        "browse": "Browse...", "opt_type": "By type", "opt_date": "By date",
        "opt_content": "By content", "opt_name": "By name",
        "content_rules": "Rules (e.g., artemisia -> artemisia):",
        "btn_preview": "Preview", "btn_execute": "Execute", "btn_cancel": "Cancel",
        "btn_scan_dups": "Find duplicates", "btn_move_dups": "Move duplicates",
        "btn_start_watch": "▶ Start Watch", "btn_stop_watch": "⏹ Stop Watch",
        "status_ready": "Ready.", "status_scanning": "Scanning...",
        "status_reading": "Reading: {file}", "status_preview_ready": "{n} files ready.",
        "status_executing": "Moving... {cur}/{total}", "status_done": "✓ {n} moved.",
        "status_undoing": "Undoing...", "status_undone": "✓ Undone.",
        "status_scan_dups": "Searching duplicates...", "status_dups_found": "{n} groups found.",
        "status_moving_dups": "Moving duplicates...", "status_dups_moved": "✓ {n} moved.",
        "status_watch_active": "👁 Watching: {folder}",
        "status_watch_inactive": "Watch stopped.",
        "col_file": "File", "col_dest": "Destination", "col_rule": "Rule",
        "col_date": "Date", "col_files": "Files", "col_status": "Status",
        "col_dup_size": "Size", "col_dup_files": "Duplicate files",
        "btn_undo": "Undo", "status_active": "Active", "status_undone_label": "Undone",
        "confirm_execute": "Move {n} files?", "confirm_dups": "Move {n} duplicates?",
        "no_moves": "No moves.", "no_dups": "No duplicates found.",
        "ask_folder": "Select a folder.", "errors_title": "Errors",
        "watch_msg_title": "Watch", "watch_msg_body": "App will minimize to tray.\nNew files will be organized automatically.",
    },
}

TYPE_CATEGORIES = {
    "Documentos": [".doc", ".docx", ".odt", ".rtf", ".txt", ".md"],
    "PDFs": [".pdf"], "eBooks": [".epub", ".mobi", ".azw3"],
    "Imágenes": [".jpg", ".jpeg", ".png", ".gif", ".webp"],
    "Audio": [".mp3", ".wav", ".flac", ".m4a"], "Video": [".mp4", ".mkv", ".avi"],
    "Comprimidos": [".zip", ".rar", ".7z"], "Otros": []
}

def safe_folder_name(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip(". ")[:100] or "Otros"

def unique_destination(dest_path: Path) -> Path:
    if not dest_path.exists(): return dest_path
    c = 1
    while c < 10000:
        candidate = dest_path.parent / f"{dest_path.stem} ({c}){dest_path.suffix}"
        if not candidate.exists(): return candidate
        c += 1
    raise RuntimeError("Colisión extrema")

def normalize_text(text: str) -> str:
    if not text: return ""
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")

def get_file_date(path: Path, source: str) -> datetime:
    st = path.stat()
    ts = getattr(st, "st_birthtime", st.st_ctime) if source == "ctime" else st.st_mtime
    return datetime.fromtimestamp(ts)

def extract_text(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    text = ""
    if ext == ".pdf" and HAS_FITZ:
        try:
            doc = fitz.open(file_path)
            for page in doc: text += page.get_text("text") + "\n"
            doc.close()
        except Exception: pass
    elif ext in (".docx", ".doc") and HAS_DOCX:
        try:
            doc = docx.Document(file_path)
            text = "\n".join(p.text for p in doc.paragraphs)
        except Exception: pass
    return normalize_text(text)

def parse_content_rules(raw_text: str) -> list[dict]:
    rules = []
    for line in raw_text.splitlines():
        if "->" in line:
            left, right = line.split("->", 1)
            folder = safe_folder_name(right.strip())
            kws = [normalize_text(k.strip()) for k in left.split(",") if k.strip()]
            if kws and folder: rules.append({"keywords": kws, "folder": folder})
    return rules

def classify_file(file_path: Path, use_type: bool, use_date: bool, date_source: str, name_rule: dict | None, content_rules: list[dict], lang: str = "es") -> tuple[str, str]:
    parts, rules_desc = [], []
    content_folder = None
    if content_rules:
        text = extract_text(file_path)
        if text:
            for rule in content_rules:
                if any(kw in text for kw in rule["keywords"]):
                    content_folder = rule["folder"]; rules_desc.append(f"Contenido: {content_folder}"); break

    name_folder = None
    if not content_folder and name_rule and name_rule.get("enabled"):
        pass

    category = None
    if not content_folder and not name_folder and use_type:
        ext = file_path.suffix.lower()
        for cat, exts in TYPE_CATEGORIES.items():
            if ext in exts: category = cat; break
        category = category or "Otros"

    if content_folder: parts.append(content_folder)
    elif name_folder: parts.append(name_folder)
    elif category: parts.append(category)

    if use_date:
        try:
            dt = get_file_date(file_path, date_source)
            parts.append(str(dt.year)); parts.append(f"{dt.month:02d}")
        except Exception: pass

    return "/".join(parts) if parts else "Otros", " | ".join(rules_desc) or "Básica"

def scan_files(folder: Path, exclude: Path | None = None) -> list[Path]:
    res = []
    exclude_res = exclude.resolve() if exclude and exclude.exists() else None
    for root, dirs, files in os.walk(folder):
        root_path = Path(root)
        if exclude_res: dirs[:] = [d for d in dirs if (root_path / d).resolve() != exclude_res]
        for name in files:
            fp = root_path / name
            try:
                if not fp.is_symlink() and fp.is_file(): res.append(fp)
            except OSError: pass
    return res

def hash_file(path: Path, chunk_size=8192) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(chunk_size): h.update(chunk)
    return h.hexdigest()

def find_duplicates(folder: Path) -> list[dict]:
    files = scan_files(folder)
    by_size = {}
    for f in files:
        try:
            size = f.stat().st_size
            if size > 0: by_size.setdefault(size, []).append(f)
        except OSError: pass
    dups = []
    for size, group in by_size.items():
        if len(group) < 2: continue
        by_hash = {}
        for f in group:
            try: by_hash.setdefault(hash_file(f), []).append(f)
            except OSError: pass
        for h, fs in by_hash.items():
            if len(fs) > 1: dups.append({"hash": h, "size": size, "files": fs})
    return dups

# =============================================================================
# HISTORIAL
# =============================================================================
class HistoryDB:
    def __init__(self, db_path: Path):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("CREATE TABLE IF NOT EXISTS organizations (id INTEGER PRIMARY KEY, timestamp TEXT, source_folder TEXT, dest_folder TEXT, total_files INTEGER, status TEXT DEFAULT 'active')")
        self.conn.execute("CREATE TABLE IF NOT EXISTS moves (id INTEGER PRIMARY KEY, org_id INTEGER, original_path TEXT, new_path TEXT, rule TEXT, FOREIGN KEY(org_id) REFERENCES organizations(id))")
        self.conn.commit()
    def save_org(self, source: str, dest: str, moves: list[dict]) -> int:
        cur = self.conn.execute("INSERT INTO organizations (timestamp, source_folder, dest_folder, total_files) VALUES (?, ?, ?, ?)", (datetime.now().isoformat(timespec="seconds"), source, dest, len(moves)))
        org_id = cur.lastrowid
        self.conn.executemany("INSERT INTO moves (org_id, original_path, new_path, rule) VALUES (?, ?, ?, ?)", [(org_id, m["source"], m["dest"], m["rule"]) for m in moves])
        self.conn.commit(); return org_id
    def list_orgs(self): return self.conn.execute("SELECT id, timestamp, source_folder, total_files, status FROM organizations ORDER BY id DESC").fetchall()
    def get_moves(self, org_id): return self.conn.execute("SELECT original_path, new_path FROM moves WHERE org_id = ? ORDER BY id DESC", (org_id,)).fetchall()
    def mark_undone(self, org_id): self.conn.execute("UPDATE organizations SET status = 'undone' WHERE id = ?", (org_id,)); self.conn.commit()

# =============================================================================
# WORKERS
# =============================================================================
class ExecuteWorker(threading.Thread):
    def __init__(self, moves, source, dest, db, on_progress, on_done, on_error):
        super().__init__(daemon=True)
        self.moves, self.source, self.dest, self.db = moves, source, dest, db
        self.on_progress, self.on_done, self.on_error = on_progress, on_done, on_error
        self.cancelled = False
    def cancel(self): self.cancelled = True
    def run(self):
        executed, errors = [], []
        for i, mv in enumerate(self.moves):
            if self.cancelled: break
            try:
                src = Path(mv["source"]); dest_dir = Path(mv["dest_folder"]); dest_dir.mkdir(parents=True, exist_ok=True)
                final = unique_destination(dest_dir / src.name)
                shutil.move(str(src), str(final))
                executed.append({"source": str(src), "dest": str(final), "rule": mv["rule"]})
            except Exception as e: errors.append(f"{Path(mv['source']).name}: {e}")
            self.on_progress(i + 1, len(self.moves))
        if executed: self.db.save_org(self.source, self.dest, executed)
        self.on_done(len(executed), errors, self.cancelled)

class UndoWorker(threading.Thread):
    def __init__(self, db, org_id, on_progress, on_done, on_error):
        super().__init__(daemon=True)
        self.db, self.org_id = db, org_id
        self.on_progress, self.on_done, self.on_error = on_progress, on_done, on_error
    def run(self):
        try:
            moves = self.db.get_moves(self.org_id); ok, errors = 0, []
            for i, (orig, new) in enumerate(moves):
                try:
                    new_path, orig_path = Path(new), Path(orig)
                    if new_path.exists():
                        orig_path.parent.mkdir(parents=True, exist_ok=True)
                        final = unique_destination(orig_path)
                        shutil.move(str(new_path), str(final))
                    ok += 1
                except Exception as e: errors.append(f"{Path(new).name}: {e}")
                self.on_progress(i + 1, len(moves))
            self.db.mark_undone(self.org_id); self.on_done(ok, errors)
        except Exception as e: self.on_error(str(e))

class WatchEventHandler(FileSystemEventHandler):
    def __init__(self, app, dest, use_type, use_date, use_content, raw_content):
        self.app = app
        self.dest = dest
        self.use_type = use_type
        self.use_date = use_date
        self.use_content = use_content
        self.content_rules = parse_content_rules(raw_content) if use_content else []
        self._active_files = set()

    def _is_temp(self, path: Path) -> bool:
        temp_exts = {'.tmp', '.partial', '.crdownload', '.download', '.opdownload', '.part', '.filepart'}
        return path.suffix.lower() in temp_exts or path.name.startswith('~')

    def on_created(self, event):
        if event.is_directory: return
        file_path = Path(event.src_path)
        if self._is_temp(file_path): return
        self._queue_file(file_path)

    def on_moved(self, event):
        if event.is_directory: return
        file_path = Path(event.dest_path)
        if self._is_temp(file_path): return
        self._queue_file(file_path)

    def _queue_file(self, file_path: Path):
        if file_path in self._active_files: return
        self._active_files.add(file_path)
        threading.Thread(target=self._process_file, args=(file_path,), daemon=True).start()

    def _process_file(self, file_path: Path):
        try:
            last_size = -1
            stable_count = 0
            while stable_count < 2:
                try: current_size = file_path.stat().st_size
                except FileNotFoundError: return
                if current_size == last_size and current_size > 0: stable_count += 1
                else: stable_count = 0
                last_size = current_size
                time.sleep(1.5)

            rel, rule = classify_file(file_path, self.use_type, self.use_date, "mtime", None, self.content_rules)
            dest_dir = self.dest / rel
            dest_dir.mkdir(parents=True, exist_ok=True)
            final = unique_destination(dest_dir / file_path.name)
            shutil.move(str(file_path), str(final))
            self.app.after(0, self.app._on_watch_move, str(file_path), str(final), rule)
        except Exception as e: print(f"Watch error: {e}")
        finally: self._active_files.discard(file_path)

# =============================================================================
# INTERFAZ
# =============================================================================
class OrganizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.lang = "es"
        self.history = HistoryDB(DB_PATH)
        self.preview_moves = []
        self.duplicates = []
        self.worker = None
        self.watch_observer = None
        self.tray_icon = None

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.geometry("1150x750")
        self.minsize(1000, 650)
        self._build_ui()
        self._refresh_history()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def t(self, key, **kwargs):
        txt = TEXTS[self.lang].get(key, key)
        return txt.format(**kwargs) if kwargs else txt

    def _set_lang(self, lang):
        self.lang = lang
        self._rebuild_ui()

    def _toggle_theme(self):
        mode = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        ctk.set_appearance_mode(mode)
        self.theme_btn.configure(text="☀️" if mode == "light" else "🌙")
        self._style_trees()

    def _rebuild_ui(self):
        for w in self.winfo_children(): w.destroy()
        self._build_ui()
        self._refresh_history()

    def _build_ui(self):
        self.title(self.t("title"))
        top = ctk.CTkFrame(self, height=40); top.pack(fill="x", padx=12, pady=(12, 0))
        ctk.CTkLabel(top, text=f"🗂️ {self.t('title')}", font=("Segoe UI", 16, "bold")).pack(side="left", padx=8)
        self.theme_btn = ctk.CTkButton(top, text="🌙", width=40, command=self._toggle_theme); self.theme_btn.pack(side="right", padx=4)
        ctk.CTkOptionMenu(top, values=["es", "en"], command=self._set_lang, width=80).pack(side="right", padx=4)

        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=8)
        self.tabs.add(self.t("tab_organize")); self.tabs.add(self.t("tab_dups"))
        self.tabs.add(self.t("tab_watch")); self.tabs.add(self.t("tab_history"))

        self._build_tab_organize(self.tabs.tab(self.t("tab_organize")))
        self._build_tab_dups(self.tabs.tab(self.t("tab_dups")))
        self._build_tab_watch(self.tabs.tab(self.t("tab_watch")))
        self._build_tab_history(self.tabs.tab(self.t("tab_history")))

    def _build_tab_organize(self, parent):
        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x", pady=4)
        ctk.CTkLabel(r1, text=self.t("src_folder"), width=120, anchor="w").pack(side="left")
        self.src_var = ctk.StringVar(); ctk.CTkEntry(r1, textvariable=self.src_var).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(r1, text=self.t("browse"), width=90, command=self._pick_src).pack(side="left")

        r2 = ctk.CTkFrame(parent, fg_color="transparent"); r2.pack(fill="x", pady=4)
        ctk.CTkLabel(r2, text=self.t("dest_folder"), width=120, anchor="w").pack(side="left")
        self.dest_var = ctk.StringVar(); ctk.CTkEntry(r2, textvariable=self.dest_var).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(r2, text=self.t("browse"), width=90, command=self._pick_dest).pack(side="left")

        opts = ctk.CTkFrame(parent); opts.pack(fill="x", pady=8)
        self.opt_type_var = ctk.BooleanVar(value=True); ctk.CTkCheckBox(opts, text=self.t("opt_type"), variable=self.opt_type_var).pack(side="left", padx=8, pady=8)
        self.opt_date_var = ctk.BooleanVar(value=False); ctk.CTkCheckBox(opts, text=self.t("opt_date"), variable=self.opt_date_var).pack(side="left", padx=8)
        self.opt_content_var = ctk.BooleanVar(value=False); ctk.CTkCheckBox(opts, text=self.t("opt_content"), variable=self.opt_content_var, command=self._toggle_content).pack(side="left", padx=8)

        self.content_frame = ctk.CTkFrame(parent); self.content_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(self.content_frame, text=self.t("content_rules")).pack(anchor="w", padx=8)
        self.content_rules_text = ctk.CTkTextbox(self.content_frame, height=60)
        self.content_rules_text.pack(fill="x", padx=8, pady=4)
        self.content_rules_text.insert("1.0", "artemisa, ajenjo -> artemisa\nabedul -> Abedul")
        self._toggle_content()

        btn_row = ctk.CTkFrame(parent, fg_color="transparent"); btn_row.pack(fill="x", pady=8)
        self.btn_preview = ctk.CTkButton(btn_row, text=self.t("btn_preview"), command=self._on_preview, width=140); self.btn_preview.pack(side="left", padx=4)
        self.btn_execute = ctk.CTkButton(btn_row, text=self.t("btn_execute"), command=self._on_execute, width=140, state="disabled", fg_color="#2e7d32"); self.btn_execute.pack(side="left", padx=4)
        self.btn_cancel = ctk.CTkButton(btn_row, text=self.t("btn_cancel"), command=self._on_cancel, width=100, state="disabled", fg_color="#c62828"); self.btn_cancel.pack(side="left", padx=4)

        self.progress = ctk.CTkProgressBar(parent, height=8); self.progress.pack(fill="x", pady=4); self.progress.set(0)
        self.status_var = ctk.StringVar(value=self.t("status_ready"))
        ctk.CTkLabel(parent, textvariable=self.status_var, anchor="w").pack(fill="x", pady=2)

        t_frame = ctk.CTkFrame(parent); t_frame.pack(fill="both", expand=True, pady=(8, 0))
        self.tree = ttk.Treeview(t_frame, columns=("file", "dest", "rule"), show="headings")
        self.tree.heading("file", text=self.t("col_file")); self.tree.heading("dest", text=self.t("col_dest")); self.tree.heading("rule", text=self.t("col_rule"))
        self.tree.column("file", width=280); self.tree.column("dest", width=350); self.tree.column("rule", width=300)
        vsb = ttk.Scrollbar(t_frame, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8); vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)

    def _build_tab_dups(self, parent):
        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x", pady=8)
        ctk.CTkLabel(r1, text=self.t("src_folder"), width=120, anchor="w").pack(side="left")
        self.dup_src_var = ctk.StringVar(); ctk.CTkEntry(r1, textvariable=self.dup_src_var).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(r1, text=self.t("browse"), width=90, command=self._pick_dup_src).pack(side="left")
        self.btn_scan_dups = ctk.CTkButton(r1, text=self.t("btn_scan_dups"), command=self._on_scan_dups, width=160); self.btn_scan_dups.pack(side="left", padx=4)
        self.btn_move_dups = ctk.CTkButton(r1, text=self.t("btn_move_dups"), command=self._on_move_dups, width=160, state="disabled", fg_color="#c62828"); self.btn_move_dups.pack(side="left", padx=4)

        self.dup_status_var = ctk.StringVar(value=self.t("status_ready"))
        ctk.CTkLabel(parent, textvariable=self.dup_status_var, anchor="w").pack(fill="x", pady=2)

        t_frame = ctk.CTkFrame(parent); t_frame.pack(fill="both", expand=True, pady=8)
        self.dup_tree = ttk.Treeview(t_frame, columns=("size", "files"), show="headings")
        self.dup_tree.heading("size", text=self.t("col_dup_size")); self.dup_tree.heading("files", text=self.t("col_dup_files"))
        self.dup_tree.column("size", width=100); self.dup_tree.column("files", width=700)
        vsb = ttk.Scrollbar(t_frame, orient="vertical", command=self.dup_tree.yview); self.dup_tree.configure(yscrollcommand=vsb.set)
        self.dup_tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8); vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)

    def _build_tab_watch(self, parent):
        ctk.CTkLabel(parent, text=self.t("tab_watch"), font=("Segoe UI", 14, "bold")).pack(pady=10)
        
        r1 = ctk.CTkFrame(parent, fg_color="transparent"); r1.pack(fill="x", pady=8)
        ctk.CTkLabel(r1, text=self.t("src_folder"), width=120, anchor="w").pack(side="left")
        self.watch_src_var = ctk.StringVar(); ctk.CTkEntry(r1, textvariable=self.watch_src_var).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(r1, text=self.t("browse"), width=90, command=self._pick_watch_src).pack(side="left")

        r2 = ctk.CTkFrame(parent, fg_color="transparent"); r2.pack(fill="x", pady=8)
        ctk.CTkLabel(r2, text=self.t("dest_folder"), width=120, anchor="w").pack(side="left")
        self.watch_dest_var = ctk.StringVar(); ctk.CTkEntry(r2, textvariable=self.watch_dest_var).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(r2, text=self.t("browse"), width=90, command=self._pick_watch_dest).pack(side="left")

        opts = ctk.CTkFrame(parent); opts.pack(fill="x", pady=8)
        self.watch_type_var = ctk.BooleanVar(value=True); ctk.CTkCheckBox(opts, text=self.t("opt_type"), variable=self.watch_type_var).pack(side="left", padx=8, pady=8)
        self.watch_date_var = ctk.BooleanVar(value=False); ctk.CTkCheckBox(opts, text=self.t("opt_date"), variable=self.watch_date_var).pack(side="left", padx=8)
        self.watch_content_var = ctk.BooleanVar(value=True); ctk.CTkCheckBox(opts, text=self.t("opt_content"), variable=self.watch_content_var, command=self._toggle_watch_content).pack(side="left", padx=8)

        self.watch_content_frame = ctk.CTkFrame(parent); self.watch_content_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(self.watch_content_frame, text=self.t("content_rules")).pack(anchor="w", padx=8)
        self.watch_content_rules_text = ctk.CTkTextbox(self.watch_content_frame, height=60)
        self.watch_content_rules_text.pack(fill="x", padx=8, pady=4)
        self.watch_content_rules_text.insert("1.0", "artemisa, ajenjo -> artemisa\nabedul -> Abedul\nginecologia -> Ginecologia")

        btn_row = ctk.CTkFrame(parent, fg_color="transparent"); btn_row.pack(fill="x", pady=20)
        self.btn_start_watch = ctk.CTkButton(btn_row, text=self.t("btn_start_watch"), command=self._on_start_watch, width=180, fg_color="#2e7d32"); self.btn_start_watch.pack(side="left", padx=10)
        self.btn_stop_watch = ctk.CTkButton(btn_row, text=self.t("btn_stop_watch"), command=self._on_stop_watch, width=180, state="disabled", fg_color="#c62828"); self.btn_stop_watch.pack(side="left", padx=10)

        self.watch_status_var = ctk.StringVar(value=self.t("status_watch_inactive"))
        ctk.CTkLabel(parent, textvariable=self.watch_status_var, font=("Segoe UI", 12)).pack(pady=10)

    def _build_tab_history(self, parent):
        self.history_tree = ttk.Treeview(parent, columns=("date", "folder", "files", "status"), show="headings")
        self.history_tree.heading("date", text=self.t("col_date")); self.history_tree.heading("folder", text=self.t("src_folder"))
        self.history_tree.heading("files", text=self.t("col_files")); self.history_tree.heading("status", text=self.t("col_status"))
        self.history_tree.column("date", width=160); self.history_tree.column("folder", width=500); self.history_tree.column("files", width=80, anchor="center"); self.history_tree.column("status", width=100, anchor="center")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=self.history_tree.yview); self.history_tree.configure(yscrollcommand=vsb.set)
        self.history_tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8); vsb.pack(side="right", fill="y", padx=(0, 8), pady=8)
        btn_row = ctk.CTkFrame(parent, fg_color="transparent"); btn_row.pack(fill="x", padx=8, pady=4)
        self.btn_undo = ctk.CTkButton(btn_row, text=self.t("btn_undo"), command=self._on_undo, width=140, state="disabled", fg_color="#c62828"); self.btn_undo.pack(side="left", padx=4)
        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_select)

    def _style_trees(self):
        style = ttk.Style(); style.theme_use("clam")
        dark = ctk.get_appearance_mode() == "Dark"
        bg, fg = ("#2b2b2b", "#ffffff") if dark else ("#ffffff", "#000000")
        head_bg = "#3b3b3b" if dark else "#d9d9d9"
        style.configure("Treeview", background=bg, foreground=fg, fieldbackground=bg, rowheight=26)
        style.configure("Treeview.Heading", background=head_bg, foreground=fg, relief="flat")
        style.map("Treeview", background=[("selected", "#1f6aa5")])

    def _toggle_content(self):
        state = "normal" if self.opt_content_var.get() else "disabled"
        self.content_rules_text.configure(state=state)

    def _toggle_watch_content(self):
        state = "normal" if self.watch_content_var.get() else "disabled"
        self.watch_content_rules_text.configure(state=state)

    def _pick_src(self): f = filedialog.askdirectory(); self.src_var.set(f) if f else None
    def _pick_dest(self): f = filedialog.askdirectory(); self.dest_var.set(f) if f else None
    def _pick_dup_src(self): f = filedialog.askdirectory(); self.dup_src_var.set(f) if f else None
    def _pick_watch_src(self): f = filedialog.askdirectory(); self.watch_src_var.set(f) if f else None
    def _pick_watch_dest(self): f = filedialog.askdirectory(); self.watch_dest_var.set(f) if f else None

    def _on_preview(self):
        src = self.src_var.get().strip()
        if not src or not Path(src).is_dir(): return messagebox.showwarning(self.t("title"), self.t("ask_folder"))
        dest = self.dest_var.get().strip() or str(Path(src) / DEFAULT_DEST_NAME)
        self._set_busy(True, self.btn_preview); self.status_var.set(self.t("status_scanning"))
        threading.Thread(target=self._preview_worker, args=(Path(src), Path(dest)), daemon=True).start()

    def _preview_worker(self, src, dest):
        try:
            files = scan_files(src, exclude=dest)
            raw_content = self.content_rules_text.get("1.0", "end") if self.opt_content_var.get() else ""
            content_rules = parse_content_rules(raw_content)
            moves = []
            for f in files:
                self.after(0, lambda file=f.name: self.status_var.set(self.t("status_reading", file=file)))
                rel, rule = classify_file(f, self.opt_type_var.get(), self.opt_date_var.get(), "mtime", None, content_rules)
                moves.append({"source": str(f), "dest_folder": str(dest / rel), "rel_folder": rel, "rule": rule})
            self.after(0, lambda: self._populate_preview(moves))
        except Exception as e: self.after(0, lambda: self._on_error(str(e)))

    def _populate_preview(self, moves):
        self.preview_moves = moves
        for item in self.tree.get_children(): self.tree.delete(item)
        for mv in moves: self.tree.insert("", "end", values=(Path(mv["source"]).name, mv["rel_folder"], mv["rule"]))
        self.status_var.set(self.t("status_preview_ready", n=len(moves)) if moves else self.t("no_moves"))
        self.btn_execute.configure(state="normal" if moves else "disabled")
        self._set_busy(False, self.btn_preview)

    def _on_execute(self):
        if not self.preview_moves or not messagebox.askyesno(self.t("title"), self.t("confirm_execute", n=len(self.preview_moves))): return
        self._set_busy(True, self.btn_preview); self.btn_cancel.configure(state="normal")
        self.worker = ExecuteWorker(self.preview_moves, self.src_var.get(), self.dest_var.get(), self.history, lambda c, t: self.after(0, self.progress.set, c/t), self._on_execute_done, lambda e: self.after(0, self._on_error, e))
        self.worker.start()

    def _on_execute_done(self, ok, errors, cancelled):
        self._set_busy(False, self.btn_preview); self.btn_cancel.configure(state="disabled")
        self.preview_moves = []
        for item in self.tree.get_children(): self.tree.delete(item)
        self._refresh_history(); self.progress.set(1.0); self.status_var.set(self.t("status_done", n=ok))
        if errors: messagebox.showwarning(self.t("errors_title"), "\n".join(errors[:10]))

    def _on_cancel(self): self.worker.cancel() if self.worker else None

    def _on_scan_dups(self):
        src = self.dup_src_var.get().strip()
        if not src or not Path(src).is_dir(): return messagebox.showwarning(self.t("title"), self.t("ask_folder"))
        self._set_busy(True, self.btn_scan_dups); self.dup_status_var.set(self.t("status_scan_dups"))
        threading.Thread(target=self._dups_worker, args=(Path(src),), daemon=True).start()

    def _dups_worker(self, src):
        try:
            dups = find_duplicates(src)
            self.after(0, lambda: self._on_dups_found(dups))
        except Exception as e: self.after(0, lambda: self._on_error(str(e)))

    def _on_dups_found(self, dups):
        self.duplicates = dups
        for item in self.dup_tree.get_children(): self.dup_tree.delete(item)
        for d in dups: self.dup_tree.insert("", "end", values=(f"{d['size'] / (1024*1024):.2f} MB", " | ".join(f.name for f in d['files'])))
        self._set_busy(False, self.btn_scan_dups)
        self.dup_status_var.set(self.t("status_dups_found", n=len(dups)) if dups else self.t("no_dups"))
        self.btn_move_dups.configure(state="normal" if dups else "disabled")

    def _on_move_dups(self):
        if not self.duplicates or not messagebox.askyesno(self.t("title"), self.t("confirm_dups", n=sum(len(d["files"]) - 1 for d in self.duplicates))): return
        dest = Path(self.dup_src_var.get()) / DUP_FOLDER_NAME
        moves = [{"source": str(f), "dest_folder": str(dest), "rel_folder": DUP_FOLDER_NAME, "rule": "Duplicado"} for d in self.duplicates for f in sorted(d["files"], key=lambda x: x.name)[1:]]
        self._set_busy(True, self.btn_move_dups)
        self.worker = ExecuteWorker(moves, self.dup_src_var.get(), str(dest), self.history, lambda c, t: self.after(0, self.progress.set, c/t), self._on_dups_moved, lambda e: self.after(0, self._on_error, e))
        self.worker.start()

    def _on_dups_moved(self, ok, errors, cancelled):
        self._set_busy(False, self.btn_move_dups); self.btn_move_dups.configure(state="disabled")
        for item in self.dup_tree.get_children(): self.dup_tree.delete(item)
        self._refresh_history(); self.dup_status_var.set(self.t("status_dups_moved", n=ok, folder=DUP_FOLDER_NAME))

    def _on_start_watch(self):
        src = self.watch_src_var.get().strip(); dest = self.watch_dest_var.get().strip()
        if not src or not dest or not Path(src).is_dir(): return messagebox.showwarning(self.t("title"), self.t("ask_folder"))
        
        messagebox.showinfo(self.t("watch_msg_title"), self.t("watch_msg_body"))
        self.withdraw()
        self._create_tray_icon()

        use_content = self.watch_content_var.get()
        raw_content = self.watch_content_rules_text.get("1.0", "end") if use_content else ""
        event_handler = WatchEventHandler(self, Path(dest), self.watch_type_var.get(), self.watch_date_var.get(), use_content, raw_content)
        
        self.watch_observer = Observer()
        self.watch_observer.schedule(event_handler, src, recursive=False)
        self.watch_observer.start()

        self.btn_start_watch.configure(state="disabled"); self.btn_stop_watch.configure(state="normal")
        self.watch_status_var.set(self.t("status_watch_active", folder=src))

    def _on_stop_watch(self):
        if self.watch_observer:
            self.watch_observer.stop(); self.watch_observer.join(); self.watch_observer = None
        self.btn_start_watch.configure(state="normal"); self.btn_stop_watch.configure(state="disabled")
        self.watch_status_var.set(self.t("status_watch_inactive"))
        if self.tray_icon: self.tray_icon.stop(); self.tray_icon = None

    def _on_watch_move(self, src_path, dest_path, rule):
        self.history.save_org(Path(src_path).parent.as_posix(), Path(dest_path).parent.as_posix(), [{"source": src_path, "dest": dest_path, "rule": rule}])
        self._refresh_history()

    def _create_tray_icon(self):
        img = Image.new('RGB', (64, 64), color='#1f6aa5')
        d = ImageDraw.Draw(img); d.text((20, 20), "🗂", fill=(255, 255, 255))
        menu = pystray.Menu(
            pystray.MenuItem("Mostrar ventana", self._show_window),
            pystray.MenuItem("Detener vigilancia", self._on_stop_watch)
        )
        self.tray_icon = pystray.Icon("FileOrganizer", img, "File Organizer", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _show_window(self):
        self.deiconify()
        if self.tray_icon: self.tray_icon.stop(); self.tray_icon = None

    def _refresh_history(self):
        for item in self.history_tree.get_children(): self.history_tree.delete(item)
        for org_id, ts, src, files, status in self.history.list_orgs():
            stat = self.t("status_active") if status == "active" else self.t("status_undone_label")
            self.history_tree.insert("", "end", iid=str(org_id), values=(ts.replace("T"," "), src, files, stat))

    def _on_history_select(self, _):
        sel = self.history_tree.selection()
        if not sel: return
        val = self.history_tree.item(sel[0], "values")
        self.btn_undo.configure(state="normal" if val[-1] == self.t("status_active") else "disabled")

    def _on_undo(self):
        sel = self.history_tree.selection()
        if not sel: return
        self._set_busy(True, self.btn_undo); self.status_var.set(self.t("status_undoing"))
        UndoWorker(self.history, int(sel[0]), lambda c, t: self.after(0, self.progress.set, c/t), lambda ok, errs: self.after(0, self._refresh_history), lambda e: self.after(0, self._on_error, e)).start()

    def _on_error(self, msg):
        self._set_busy(False, self.btn_preview); self._set_busy(False, self.btn_scan_dups)
        messagebox.showerror(self.t("title"), msg)

    def _set_busy(self, busy, btn):
        try: btn.configure(state="disabled" if busy else "normal")
        except: pass

    def _on_close(self):
        if self.watch_observer:
            if messagebox.askyesno(self.t("title"), "La vigilancia está activa. ¿Cerrar de todos modos?"):
                self._on_stop_watch(); self.destroy()
        else:
            self.destroy()

def main():
    app = OrganizerApp(); app.mainloop()

if __name__ == "__main__":
    if "--test" in sys.argv: unittest.main(argv=[sys.argv[0]], verbosity=2)
    else: main()