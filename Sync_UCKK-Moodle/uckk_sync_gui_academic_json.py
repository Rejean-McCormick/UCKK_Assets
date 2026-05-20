import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "UCKK -> Moodle Sync + Academic JSON"
DEFAULT_SOURCE = r"C:\mycode\UCKK\uckk-moodle"
DEFAULT_DEST = r"C:\mycode\UCKK\moodle\moodle\public"
DEFAULT_ACADEMIC_JSON_FOLDER = "academic_registry_json"

EXPECTED_ACADEMIC_JSON_FILES = [
    "archive_templates.json",
    "assembly_templates.json",
    "badges.json",
    "capabilities.json",
    "categories.json",
    "challenge_templates.json",
    "cohorts.json",
    "competencies.json",
    "course_templates.json",
    "courses.json",
    "pathways.json",
    "programs.json",
    "reports.json",
    "roles.json",
]

# Moodle plugin/component folders copied from the UCKK source package into Moodle.
PLUGIN_COMPONENTS = [
    r"admin\tool\uckkintegrity",
    r"admin\tool\uckkseed",
    r"ai\provider\uckk",
    r"blocks\uckk_dashboard",
    r"course\format\uckk",
    r"local\uckk",
    r"mod\uckkarchive",
    r"mod\uckkassembly",
    r"mod\uckkchallenge",
    r"report\uckk",
    r"theme\uckk",
]


def path_from_rel(root: Path, relpath: str) -> Path:
    """Build a path from a Moodle-style relative path, portable enough for Windows/Linux."""
    parts = relpath.replace("/", "\\").split("\\")
    return root.joinpath(*[p for p in parts if p])


def fmt_ts(ts):
    if ts is None:
        return "—"
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def latest_mtime(path: Path):
    if not path.exists():
        return None
    if path.is_file():
        return path.stat().st_mtime
    latest = path.stat().st_mtime
    for root, dirs, files in os.walk(path):
        try:
            latest = max(latest, Path(root).stat().st_mtime)
        except OSError:
            pass
        for name in files:
            p = Path(root) / name
            try:
                latest = max(latest, p.stat().st_mtime)
            except OSError:
                pass
    return latest


def count_files(path: Path):
    if not path.exists():
        return 0
    if path.is_file():
        return 1
    total = 0
    for _, _, files in os.walk(path):
        total += len(files)
    return total


def count_json_files(path: Path):
    if not path.exists() or not path.is_dir():
        return 0
    return sum(1 for p in path.rglob("*.json") if p.is_file())


def validate_json_folder(path: Path):
    """Return (ok, message). Validates expected academic registry JSON files and syntax."""
    if not path.exists():
        return False, "source missing"
    if not path.is_dir():
        return False, "not a folder"

    missing = [name for name in EXPECTED_ACADEMIC_JSON_FILES if not (path / name).is_file()]
    if missing:
        return False, "missing json files:\n" + "\n".join(missing)

    json_files = sorted(p for p in path.rglob("*.json") if p.is_file())
    if not json_files:
        return False, "no json files"

    errors = []
    for p in json_files:
        try:
            with p.open("r", encoding="utf-8-sig") as f:
                json.load(f)
        except Exception as e:
            rel = p.relative_to(path)
            errors.append(f"{rel}: {e}")

    if errors:
        return False, "json invalid:\n" + "\n".join(errors[:20])

    extra = [p.name for p in json_files if p.name not in EXPECTED_ACADEMIC_JSON_FILES]
    if extra:
        return True, f"json ok ({len(json_files)}), extra files: {len(extra)}"

    return True, f"json ok ({len(json_files)}/{len(EXPECTED_ACADEMIC_JSON_FILES)})"


def safe_rmtree(path: Path):
    if path.exists():
        shutil.rmtree(path)


def copy_folder(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        safe_rmtree(dst)
    shutil.copytree(src, dst)


def archive_folder(dst: Path, archive_root: Path, relpath: str, stamp: str):
    if not dst.exists():
        return None
    archive_target = path_from_rel(archive_root / stamp, relpath)
    archive_target.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(dst), str(archive_target))
    return archive_target


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1450x740")
        self.minsize(1180, 620)

        self.source_var = tk.StringVar(value=DEFAULT_SOURCE)
        self.dest_var = tk.StringVar(value=DEFAULT_DEST)
        self.academic_json_folder_var = tk.StringVar(value=DEFAULT_ACADEMIC_JSON_FOLDER)
        self.archive_var = tk.StringVar(value=str(Path(DEFAULT_DEST).parent / "_uckk_archive"))
        self.status_var = tk.StringVar(value="Prêt.")
        self.rows = {}

        self._build_ui()
        self.refresh_table()

    def _sync_items(self):
        academic_json_rel = self.academic_json_folder_var.get().strip() or DEFAULT_ACADEMIC_JSON_FOLDER
        return [(rel, "plugin") for rel in PLUGIN_COMPONENTS] + [(academic_json_rel, "academic-json")]

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        self._path_row(top, "Source UCKK", self.source_var, self._pick_source, 0)
        self._path_row(top, "Destination Moodle", self.dest_var, self._pick_dest, 1)
        self._text_row(top, "Academic JSON folder", self.academic_json_folder_var, 2)
        self._path_row(top, "Archive root", self.archive_var, self._pick_archive, 3)

        hint = ttk.Label(
            top,
            text="Academic JSON folder is synced root-to-root, for example: uckk-moodle\\academic_registry_json -> moodle\\moodle\\academic_registry_json",
            foreground="#555555",
        )
        hint.grid(row=4, column=1, columnspan=2, sticky="w", pady=(2, 0))

        btns = ttk.Frame(self, padding=(10, 0, 10, 10))
        btns.pack(fill="x")

        ttk.Button(btns, text="Refresh", command=self.refresh_table).pack(side="left")
        ttk.Button(btns, text="Validate academic JSON", command=self.validate_academic_json).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Select all", command=lambda: self._set_all(True)).pack(side="left", padx=(16, 0))
        ttk.Button(btns, text="Select none", command=lambda: self._set_all(False)).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Replace selected", command=self.replace_selected).pack(side="left", padx=(20, 0))
        ttk.Button(btns, text="Archive + replace selected", command=self.archive_replace_selected).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Open source", command=lambda: self._open_folder(Path(self.source_var.get()))).pack(side="right")
        ttk.Button(btns, text="Open dest", command=lambda: self._open_folder(Path(self.dest_var.get()))).pack(side="right", padx=(0, 8))

        table_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        table_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(table_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        headers = [
            ("Use", 6),
            ("Kind", 10),
            ("Path", 38),
            ("Source exists", 12),
            ("Src modified", 20),
            ("Src files", 10),
            ("Dest exists", 12),
            ("Dest modified", 20),
            ("Dest files", 10),
            ("Status", 24),
        ]

        header = ttk.Frame(self.inner)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        for idx, (label, width) in enumerate(headers):
            ttk.Label(header, text=label, width=width, anchor="w").grid(row=0, column=idx, sticky="w", padx=3)

        self._build_rows()

        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var, anchor="w").pack(fill="x")

        self.academic_json_folder_var.trace_add("write", lambda *_: self._rebuild_rows_and_refresh())

    def _build_rows(self):
        for child in self.inner.winfo_children()[1:]:
            child.destroy()
        self.rows = {}

        for idx, (rel, kind) in enumerate(self._sync_items(), start=1):
            row = ttk.Frame(self.inner)
            row.grid(row=idx, column=0, sticky="ew", pady=1)

            include_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(row, variable=include_var).grid(row=0, column=0, sticky="w", padx=3)
            ttk.Label(row, text=kind, width=10, anchor="w").grid(row=0, column=1, sticky="w", padx=3)
            ttk.Label(row, text=rel, width=38, anchor="w").grid(row=0, column=2, sticky="w", padx=3)

            src_exists = ttk.Label(row, width=12, anchor="w")
            src_exists.grid(row=0, column=3, sticky="w", padx=3)

            src_mod = ttk.Label(row, width=20, anchor="w")
            src_mod.grid(row=0, column=4, sticky="w", padx=3)

            src_files = ttk.Label(row, width=10, anchor="w")
            src_files.grid(row=0, column=5, sticky="w", padx=3)

            dst_exists = ttk.Label(row, width=12, anchor="w")
            dst_exists.grid(row=0, column=6, sticky="w", padx=3)

            dst_mod = ttk.Label(row, width=20, anchor="w")
            dst_mod.grid(row=0, column=7, sticky="w", padx=3)

            dst_files = ttk.Label(row, width=10, anchor="w")
            dst_files.grid(row=0, column=8, sticky="w", padx=3)

            status = ttk.Label(row, width=24, anchor="w")
            status.grid(row=0, column=9, sticky="w", padx=3)

            self.rows[rel] = {
                "kind": kind,
                "include": include_var,
                "src_exists": src_exists,
                "src_mod": src_mod,
                "src_files": src_files,
                "dst_exists": dst_exists,
                "dst_mod": dst_mod,
                "dst_files": dst_files,
                "status": status,
            }

    def _rebuild_rows_and_refresh(self):
        self._build_rows()
        self.refresh_table()

    def _path_row(self, parent, label, var, browse_cmd, row):
        ttk.Label(parent, text=label, width=18).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Button(parent, text="Browse", command=browse_cmd).grid(row=row, column=2, padx=(8, 0), pady=3)
        parent.grid_columnconfigure(1, weight=1)

    def _text_row(self, parent, label, var, row):
        ttk.Label(parent, text=label, width=18).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Label(parent, text="example: academic_registry_json", width=18).grid(row=row, column=2, sticky="w", padx=(8, 0), pady=3)
        parent.grid_columnconfigure(1, weight=1)

    def _pick_source(self):
        path = filedialog.askdirectory(initialdir=self.source_var.get() or os.getcwd())
        if path:
            self.source_var.set(path)
            self.refresh_table()

    def _pick_dest(self):
        path = filedialog.askdirectory(initialdir=self.dest_var.get() or os.getcwd())
        if path:
            self.dest_var.set(path)
            self.refresh_table()

    def _pick_archive(self):
        path = filedialog.askdirectory(initialdir=self.archive_var.get() or os.getcwd())
        if path:
            self.archive_var.set(path)

    def _set_all(self, value: bool):
        for row in self.rows.values():
            row["include"].set(value)

    def _open_folder(self, path: Path):
        if not path.exists():
            messagebox.showwarning(APP_TITLE, f"Dossier introuvable:\n{path}")
            return
        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Impossible d’ouvrir le dossier:\n{path}\n\n{e}")

    def _selected_items(self):
        return [(rel, row["kind"]) for rel, row in self.rows.items() if row["include"].get()]

    def _source_path(self, rel):
        return path_from_rel(Path(self.source_var.get()), rel)

    def _dest_path(self, rel):
        return path_from_rel(Path(self.dest_var.get()), rel)

    def refresh_table(self):
        source_root = Path(self.source_var.get())
        dest_root = Path(self.dest_var.get())

        moodle_markers = [dest_root / "config.php", dest_root / "config-dist.php", dest_root / "version.php", dest_root / "admin" / "cli" / "install.php"]
        if not any(p.exists() for p in moodle_markers):
            self.status_var.set("Attention: la destination ne ressemble pas à une racine Moodle.")
        else:
            self.status_var.set("Comparaison à jour.")

        for rel, widgets in self.rows.items():
            kind = widgets["kind"]
            src = self._source_path(rel)
            dst = self._dest_path(rel)

            src_exists = src.exists()
            dst_exists = dst.exists()
            src_m = latest_mtime(src)
            dst_m = latest_mtime(dst)
            src_n = count_files(src)
            dst_n = count_files(dst)

            if kind == "academic-json" and src_exists:
                ok, json_status = validate_json_folder(src)
                status = json_status if ok else "json invalid"
            elif not src_exists:
                status = "source missing"
            elif not dst_exists:
                status = "new copy"
            elif src_m and dst_m and src_m > dst_m:
                status = "source newer"
            elif src_m and dst_m and src_m < dst_m:
                status = "dest newer"
            else:
                status = "same/newest"

            widgets["src_exists"].configure(text="yes" if src_exists else "no")
            widgets["src_mod"].configure(text=fmt_ts(src_m))
            widgets["src_files"].configure(text=str(src_n))
            widgets["dst_exists"].configure(text="yes" if dst_exists else "no")
            widgets["dst_mod"].configure(text=fmt_ts(dst_m))
            widgets["dst_files"].configure(text=str(dst_n))
            widgets["status"].configure(text=status)

    def validate_academic_json(self):
        academic_json_rel = self.academic_json_folder_var.get().strip() or DEFAULT_ACADEMIC_JSON_FOLDER
        academic_json_path = self._source_path(academic_json_rel)
        ok, message = validate_json_folder(academic_json_path)
        if ok:
            messagebox.showinfo(APP_TITLE, f"Academic JSON validation OK.\n\n{academic_json_path}\n\n{message}")
        else:
            messagebox.showerror(APP_TITLE, f"Academic JSON validation failed.\n\n{academic_json_path}\n\n{message}")
        self.refresh_table()

    def _validate_before_copy(self, selected):
        missing = [rel for rel, kind in selected if not self._source_path(rel).exists()]
        if missing:
            messagebox.showerror(APP_TITLE, "Sources manquantes:\n\n" + "\n".join(missing))
            return False

        academic_json_errors = []
        for rel, kind in selected:
            if kind == "academic-json":
                ok, msg = validate_json_folder(self._source_path(rel))
                if not ok:
                    academic_json_errors.append(f"{rel}:\n{msg}")

        if academic_json_errors:
            messagebox.showerror(
                APP_TITLE,
                "Copie arrêtée: le dossier Academic JSON contient du JSON invalide.\n\n" + "\n\n".join(academic_json_errors),
            )
            return False

        return True

    def replace_selected(self):
        selected = self._selected_items()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Aucun élément sélectionné.")
            return

        if not self._validate_before_copy(selected):
            return

        if not messagebox.askyesno(APP_TITLE, "Écraser les éléments sélectionnés sans archive ?"):
            return

        done = []
        current_rel = ""
        try:
            for rel, kind in selected:
                current_rel = rel
                copy_folder(self._source_path(rel), self._dest_path(rel))
                done.append(rel)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erreur pendant la copie de {current_rel}:\n\n{e}")
        finally:
            self.refresh_table()

        if done:
            messagebox.showinfo(APP_TITLE, "Copie terminée:\n\n" + "\n".join(done))

    def archive_replace_selected(self):
        selected = self._selected_items()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Aucun élément sélectionné.")
            return

        if not self._validate_before_copy(selected):
            return

        archive_root = Path(self.archive_var.get())
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if not messagebox.askyesno(
            APP_TITLE,
            "Archiver puis remplacer les éléments sélectionnés ?\n\n"
            f"Archive root:\n{archive_root}\n\n"
            f"Archive stamp:\n{stamp}",
        ):
            return

        done = []
        archived = []
        current_rel = ""
        try:
            for rel, kind in selected:
                current_rel = rel
                src = self._source_path(rel)
                dst = self._dest_path(rel)
                archived_path = archive_folder(dst, archive_root, rel, stamp)
                if archived_path:
                    archived.append(f"{rel} -> {archived_path}")
                copy_folder(src, dst)
                done.append(rel)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erreur pendant l'opération sur {current_rel}:\n\n{e}")
        finally:
            self.refresh_table()

        msg = []
        if done:
            msg.append("Remplacés:")
            msg.extend(done)
        if archived:
            msg.append("")
            msg.append("Archivés:")
            msg.extend(archived)
        if msg:
            messagebox.showinfo(APP_TITLE, "\n".join(msg))


if __name__ == "__main__":
    app = App()
    app.mainloop()
