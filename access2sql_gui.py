#!/usr/bin/env python3
"""
access2sql_gui.py  –  drag-and-drop GUI frontend for access2sql
"""
import sys
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

from access2sql import VERSION, find_access_files, export_db


class _QueueStream:
    """Route stdout/stderr writes into a thread-safe queue for the log widget."""
    def __init__(self, q: queue.Queue) -> None:
        self._q = q

    def write(self, text: str) -> None:
        if text:
            self._q.put(text)

    def flush(self) -> None:
        pass


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"Access to SQLite  v{VERSION}")
        self.root.minsize(660, 540)

        self._files: list[Path] = []
        self._q: queue.Queue = queue.Queue()
        self._busy = False

        try:
            import pyodbc  # noqa: F401
            self._use_pyodbc = True
            self._backend = "pyodbc"
        except ImportError:
            self._use_pyodbc = False
            self._backend = "mdbtools"

        self._build_ui()
        self._tick()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)

        # Drop zone ───────────────────────────────────────────────────────────
        drop_text = (
            "Drop .accdb / .mdb files or folders here\n(or use the Browse buttons below)"
            if _HAS_DND
            else "Use the Browse buttons below to select files or a folder"
        )
        self._drop = tk.Label(
            outer, text=drop_text,
            relief="groove", font=("Helvetica", 13),
            pady=30, cursor="hand2",
        )
        self._drop.pack(fill=tk.X, pady=(0, 10))
        self._drop.bind("<Button-1>", lambda _: self._browse_files())

        if _HAS_DND:
            self._drop.drop_target_register(DND_FILES)
            self._drop.dnd_bind("<<Drop>>",      self._on_drop)
            self._drop.dnd_bind("<<DragEnter>>", lambda _: self._drop.config(relief="sunken"))
            self._drop.dnd_bind("<<DragLeave>>", lambda _: self._drop.config(relief="groove"))

        # File list ───────────────────────────────────────────────────────────
        self._list_lf = ttk.LabelFrame(outer, text="Files queued: 0", padding=4)
        self._list_lf.pack(fill=tk.BOTH, expand=False, pady=(0, 10))

        self._lb = tk.Listbox(self._list_lf, height=6, selectmode=tk.EXTENDED,
                              activestyle="dotbox")
        vsb = ttk.Scrollbar(self._list_lf, orient=tk.VERTICAL, command=self._lb.yview)
        self._lb.configure(yscrollcommand=vsb.set)
        self._lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons ─────────────────────────────────────────────────────────────
        br = ttk.Frame(outer)
        br.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(br, text="Browse files…",   command=self._browse_files).pack(side=tk.LEFT)
        ttk.Button(br, text="Browse folder…",  command=self._browse_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(br, text="Remove selected", command=self._remove_sel).pack(side=tk.LEFT)
        ttk.Button(br, text="Clear",           command=self._clear).pack(side=tk.LEFT, padx=4)

        self._gen_btn = ttk.Button(br, text="Generate Output",
                                   command=self._generate, state=tk.DISABLED)
        self._gen_btn.pack(side=tk.RIGHT)

        # Log area ────────────────────────────────────────────────────────────
        log_lf = ttk.LabelFrame(outer, text="Log", padding=4)
        log_lf.pack(fill=tk.BOTH, expand=True)

        self._log = tk.Text(
            log_lf, state=tk.DISABLED, wrap=tk.WORD,
            bg="#1e1e1e", fg="#d4d4d4",
            font=("Courier", 11), relief=tk.FLAT,
        )
        lsb = ttk.Scrollbar(log_lf, orient=tk.VERTICAL, command=self._log.yview)
        self._log.configure(yscrollcommand=lsb.set)
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Status bar ──────────────────────────────────────────────────────────
        self._status = tk.StringVar(value=f"Backend: {self._backend}")
        ttk.Label(outer, textvariable=self._status,
                  foreground="gray", font=("Helvetica", 10)).pack(anchor=tk.W, pady=(4, 0))

    # ── file management ───────────────────────────────────────────────────────

    def _on_drop(self, event) -> None:
        self._drop.config(relief="groove")
        self._add_paths([Path(p) for p in self.root.tk.splitlist(event.data)])

    def _browse_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Access database(s)",
            filetypes=[("Access databases", "*.accdb *.mdb"), ("All files", "*.*")],
        )
        self._add_paths([Path(p) for p in paths])

    def _browse_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder to scan for Access databases")
        if folder:
            self._add_paths([Path(folder)])

    def _add_paths(self, paths: list[Path]) -> None:
        existing = set(self._files)
        for p in paths:
            if p.is_dir():
                candidates = find_access_files(p)
            elif p.suffix.lower() in (".accdb", ".mdb"):
                candidates = [p]
            else:
                continue
            for db in candidates:
                if db not in existing:
                    self._files.append(db)
                    self._lb.insert(tk.END, f"{db.name}  —  {db.parent}")
                    existing.add(db)
        self._refresh()

    def _remove_sel(self) -> None:
        for idx in reversed(self._lb.curselection()):
            self._lb.delete(idx)
            self._files.pop(idx)
        self._refresh()

    def _clear(self) -> None:
        self._files.clear()
        self._lb.delete(0, tk.END)
        self._refresh()

    def _refresh(self) -> None:
        n = len(self._files)
        self._list_lf.config(text=f"Files queued: {n}")
        self._gen_btn.config(state=tk.NORMAL if n and not self._busy else tk.DISABLED)

    # ── generate ──────────────────────────────────────────────────────────────

    def _generate(self) -> None:
        if not self._files or self._busy:
            return
        self._busy = True
        self._gen_btn.config(state=tk.DISABLED, text="Working…")
        self._status.set("Processing…")

        files      = list(self._files)
        use_pyodbc = self._use_pyodbc
        q          = self._q

        def _worker() -> None:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = _QueueStream(q)
            try:
                for db in files:
                    export_db(db, use_pyodbc)
                q.put("\nDone.\n")
            except Exception as exc:
                q.put(f"\nERROR: {exc}\n")
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                self._busy = False
                self.root.after(0, self._on_done)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self) -> None:
        self._gen_btn.config(text="Generate Output")
        self._refresh()
        self._status.set(f"Backend: {self._backend}  —  finished")

    # ── log polling ───────────────────────────────────────────────────────────

    def _append_log(self, text: str) -> None:
        self._log.config(state=tk.NORMAL)
        self._log.insert(tk.END, text)
        self._log.see(tk.END)
        self._log.config(state=tk.DISABLED)

    def _tick(self) -> None:
        try:
            while True:
                self._append_log(self._q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(80, self._tick)


# ── entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
