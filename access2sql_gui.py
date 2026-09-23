#!/usr/bin/env python3
"""
access2sql_gui.py  –  customtkinter GUI frontend for access2sql
"""
import sys
import queue
import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

from access2sql import VERSION, find_access_files, export_db

ctk.set_appearance_mode("system")       # follows macOS dark / light mode
ctk.set_default_color_theme("blue")


class _QueueStream:
    """Route stdout/stderr writes into a thread-safe queue for the log widget."""
    def __init__(self, q: queue.Queue) -> None:
        self._q = q
    def write(self, text: str) -> None:
        if text:
            self._q.put(text)
    def flush(self) -> None:
        pass


# Combine CTk root with TkinterDnD when available
if _HAS_DND:
    class _Root(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self):
            super().__init__()
            self.TkdndVersion = TkinterDnD._require(self)
else:
    _Root = ctk.CTk  # type: ignore[misc]


class App(_Root):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Access to SQLite  v{VERSION}")
        self.minsize(700, 580)
        self.geometry("800x680")

        self._files: list[Path] = []
        self._file_frames: list[tuple[Path, ctk.CTkFrame]] = []
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
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)     # log area expands
        PAD = 16

        # ── Drop zone ─────────────────────────────────────────────────────────
        self._drop_frame = ctk.CTkFrame(
            self, corner_radius=12,
            fg_color=("gray88", "gray18"),
            border_width=2, border_color=("gray72", "gray32"),
        )
        self._drop_frame.grid(row=0, column=0, padx=PAD, pady=(PAD, 8), sticky="ew")
        self._drop_frame.grid_columnconfigure(0, weight=1)

        drop_text = (
            "Drop .accdb / .mdb files or folders here\n(or use the Browse buttons below)"
            if _HAS_DND else
            "Use the Browse buttons below to select files or a folder"
        )
        self._drop_label = ctk.CTkLabel(
            self._drop_frame, text=drop_text,
            font=ctk.CTkFont(size=14), pady=26, cursor="hand2",
        )
        self._drop_label.grid(row=0, column=0, padx=16, sticky="ew")

        for w in (self._drop_frame, self._drop_label):
            w.bind("<Button-1>", lambda _: self._browse_files())
            w.bind("<Enter>", lambda _: self._drop_frame.configure(
                fg_color=("gray80", "gray26"), border_color=("gray60", "gray45")))
            w.bind("<Leave>", self._drop_leave)

        if _HAS_DND:
            _BLUE = ("#3a7ebf", "#3a7ebf")
            _GREY = ("gray72", "gray32")
            for w in (self._drop_frame, self._drop_label):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>",      self._on_drop)
                w.dnd_bind("<<DragEnter>>",
                           lambda _: self._drop_frame.configure(border_color=_BLUE))
                w.dnd_bind("<<DragLeave>>",
                           lambda _: self._drop_frame.configure(border_color=_GREY))

        # ── File list ─────────────────────────────────────────────────────────
        list_card = ctk.CTkFrame(self, corner_radius=10)
        list_card.grid(row=1, column=0, padx=PAD, pady=(0, 8), sticky="ew")
        list_card.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(list_card, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=10, pady=(8, 2), sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)

        self._list_header = ctk.CTkLabel(
            hdr, text="Files queued: 0",
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
        )
        self._list_header.grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            hdr, text="Clear all", width=78, height=26, corner_radius=6,
            fg_color="transparent", hover_color=("gray80", "gray30"),
            text_color=("gray45", "gray65"), command=self._clear,
        ).grid(row=0, column=1, sticky="e")

        self._scroll_frame = ctk.CTkScrollableFrame(
            list_card, height=130, corner_radius=6,
            fg_color=("gray82", "gray14"),
        )
        self._scroll_frame.grid(row=1, column=0, padx=8, pady=(0, 10), sticky="ew")
        self._scroll_frame.grid_columnconfigure(0, weight=1)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=2, column=0, padx=PAD, pady=(0, 8), sticky="ew")
        btn_row.grid_columnconfigure(2, weight=1)   # spacer pushes Generate right

        kw = dict(corner_radius=8, height=34)
        ctk.CTkButton(btn_row, text="Browse files…",  command=self._browse_files,
                      width=130, **kw).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(btn_row, text="Browse folder…", command=self._browse_folder,
                      width=130, **kw).grid(row=0, column=1)
        # column 2 = spacer
        self._gen_btn = ctk.CTkButton(
            btn_row, text="Generate Output",
            command=self._generate, width=160, state="disabled", **kw,
        )
        self._gen_btn.grid(row=0, column=3)

        # ── Log ───────────────────────────────────────────────────────────────
        log_card = ctk.CTkFrame(self, corner_radius=10)
        log_card.grid(row=3, column=0, padx=PAD, pady=(0, 8), sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_card, text="Log",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     anchor="w").grid(row=0, column=0, padx=12, pady=(8, 2), sticky="w")

        self._log = ctk.CTkTextbox(
            log_card, wrap="word", state="disabled",
            font=ctk.CTkFont(family="Courier", size=11),
            fg_color=("#1e1e1e", "#1e1e1e"),
            text_color=("#d4d4d4", "#d4d4d4"),
            corner_radius=6,
        )
        self._log.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="nsew")

        # ── Status bar ────────────────────────────────────────────────────────
        self._status_var = ctk.StringVar(value=f"Backend: {self._backend}")
        ctk.CTkLabel(self, textvariable=self._status_var,
                     font=ctk.CTkFont(size=11),
                     text_color=("gray50", "gray55"),
                     ).grid(row=4, column=0, padx=PAD, pady=(0, 10), sticky="w")

    # ── file management ───────────────────────────────────────────────────────

    def _drop_leave(self, event) -> None:
        # Only restore when the cursor actually leaves the entire drop zone,
        # not just moves from the frame onto the child label (or vice versa).
        x = event.x_root - self._drop_frame.winfo_rootx()
        y = event.y_root - self._drop_frame.winfo_rooty()
        if not (0 <= x < self._drop_frame.winfo_width() and
                0 <= y < self._drop_frame.winfo_height()):
            self._drop_frame.configure(
                fg_color=("gray88", "gray18"), border_color=("gray72", "gray32"))

    def _on_drop(self, event) -> None:
        self._drop_frame.configure(border_color=("gray72", "gray32"))
        self._add_paths([Path(p) for p in self.tk.splitlist(event.data)])

    def _browse_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Access database(s)",
            filetypes=[("Access databases", "*.accdb *.mdb"), ("All files", "*.*")],
        )
        self._add_paths([Path(p) for p in paths])

    def _browse_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="Select folder to scan for Access databases")
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
                    self._add_file_row(db)
                    existing.add(db)
        self._refresh()

    def _add_file_row(self, db: Path) -> None:
        row = ctk.CTkFrame(
            self._scroll_frame, corner_radius=6,
            fg_color=("gray90", "gray22"),
        )
        row.pack(fill="x", padx=2, pady=2)
        row.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            row, text=db.name, anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, padx=(10, 8), pady=6)

        ctk.CTkLabel(
            row, text=str(db.parent), anchor="w",
            font=ctk.CTkFont(size=10),
            text_color=("gray50", "gray55"),
        ).grid(row=0, column=1, padx=4, sticky="ew")

        def _remove(db=db, row=row) -> None:
            self._files.remove(db)
            self._file_frames[:] = [(p, f) for p, f in self._file_frames if p != db]
            row.destroy()
            self._refresh()

        ctk.CTkButton(
            row, text="×", width=28, height=26, corner_radius=6,
            command=_remove,
            fg_color="transparent", hover_color=("gray78", "gray32"),
            text_color=("gray40", "gray70"),
        ).grid(row=0, column=2, padx=(4, 6))

        self._file_frames.append((db, row))

    def _clear(self) -> None:
        for _, frame in self._file_frames:
            frame.destroy()
        self._files.clear()
        self._file_frames.clear()
        self._refresh()

    def _refresh(self) -> None:
        n = len(self._files)
        self._list_header.configure(text=f"Files queued: {n}")
        self._gen_btn.configure(
            state="normal" if n and not self._busy else "disabled")

    # ── generate ──────────────────────────────────────────────────────────────

    def _generate(self) -> None:
        if not self._files or self._busy:
            return
        self._busy = True
        self._gen_btn.configure(state="disabled", text="Working…")
        self._status_var.set("Processing…")

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
                self.after(0, self._on_done)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self) -> None:
        self._gen_btn.configure(text="Generate Output")
        self._refresh()
        self._status_var.set(f"Backend: {self._backend}  —  finished")

    # ── log polling ───────────────────────────────────────────────────────────

    def _append_log(self, text: str) -> None:
        self._log.configure(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _tick(self) -> None:
        try:
            while True:
                self._append_log(self._q.get_nowait())
        except queue.Empty:
            pass
        self.after(80, self._tick)


# ── entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
