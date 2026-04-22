#!/usr/bin/env python3
"""
access2sql.py
---------------------------
Use a system folder-picker dialog to choose a root folder, then recursively
find every .accdb / .mdb file, extract all tables + data, and produce:
  - <db_name>.sql   : CREATE TABLE + INSERT statements (SQLite-compatible)

Type mapping from Access → SQLite:
  Text / Memo / Hyperlink  → TEXT COLLATE NOCASE
  Date/Time                → TEXT (stored as ISO-8601, cast with datetime())
  Yes/No (Boolean)         → INTEGER (0/1)  -- SQLite has no BOOLEAN type
  AutoNumber / Long Integer → INTEGER
  Integer                  → INTEGER
  Single / Double / Currency / Decimal → REAL
  OLE Object / Binary      → BLOB
  everything else          → TEXT COLLATE NOCASE

Requirements (install once):
  pip install pyodbc     # on macOS needs mdbtools (brew install mdbtools)
  -- OR --
  pip install pywin32    # Windows only (uses COM / DAO)

This script auto-detects the platform and picks the right backend.
On macOS it falls back to the `mdbtools` CLI utilities (mdb-tables, mdb-export,
mdb-schema) if pyodbc is unavailable.
"""

import os
import sys
import platform
import subprocess
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from pathlib import Path


# ══════════════════════════════════════════════════════════════════════════════
# Platform detection
# ══════════════════════════════════════════════════════════════════════════════

IS_WINDOWS = platform.system() == "Windows"
IS_MAC     = platform.system() == "Darwin"
LAST_FOLDER_FILE = Path.home() / ".access2sql.conf"


# ══════════════════════════════════════════════════════════════════════════════
# Folder chooser (native dialog)
# ══════════════════════════════════════════════════════════════════════════════

def _load_last_folder() -> str | None:
    try:
        p = LAST_FOLDER_FILE.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    if p and Path(p).is_dir():
        return p
    return None


def _save_last_folder(folder: str) -> None:
    try:
        LAST_FOLDER_FILE.write_text(folder, encoding="utf-8")
    except OSError:
        # Non-fatal: extraction should continue even if preference cannot be saved.
        pass


def pick_folder() -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    last_folder = _load_last_folder()
    kwargs = {"title": "Select root folder containing Access databases"}
    if last_folder:
        kwargs["initialdir"] = last_folder
    folder = filedialog.askdirectory(**kwargs)
    root.destroy()
    if not folder:
        print("No folder selected. Exiting.")
        sys.exit(0)
    _save_last_folder(folder)
    return Path(folder)


# ══════════════════════════════════════════════════════════════════════════════
# Find all Access databases recursively
# ══════════════════════════════════════════════════════════════════════════════

def find_access_files(root: Path) -> list[Path]:
    files = []
    for pattern in ("**/*.accdb", "**/*.mdb"):
        files.extend(root.glob(pattern))
    return sorted(set(files))


def unique_output_path(path: Path) -> Path:
    """Return a non-existing path by appending _1, _2, ... if needed."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 1
    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


# ══════════════════════════════════════════════════════════════════════════════
# Type mapping helpers
# ══════════════════════════════════════════════════════════════════════════════

# Map Access type names (lowercase) → SQLite column type declaration
_TYPE_MAP = {
    # Text-like
    "text":       "TEXT COLLATE NOCASE",
    "memo":       "TEXT COLLATE NOCASE",
    "hyperlink":  "TEXT COLLATE NOCASE",
    # Numeric
    "autonumber": "INTEGER",
    "long integer":"INTEGER",
    "integer":    "INTEGER",
    "byte":       "INTEGER",
    "single":     "REAL",
    "double":     "REAL",
    "currency":   "REAL",
    "decimal":    "REAL",
    "numeric":    "REAL",
    # Date
    "date/time":  "TEXT",          # stored as ISO-8601
    "datetime":   "TEXT",
    # Boolean
    "yes/no":     "INTEGER",       # 0 / 1
    "boolean":    "INTEGER",
    "bit":        "INTEGER",
    # Binary
    "ole object": "BLOB",
    "binary":     "BLOB",
    "varbinary":  "BLOB",
}

def access_type_to_sqlite(access_type: str) -> str:
    key = access_type.lower().strip()
    return _TYPE_MAP.get(key, "TEXT COLLATE NOCASE")


# ══════════════════════════════════════════════════════════════════════════════
# Value formatting for INSERT statements
# ══════════════════════════════════════════════════════════════════════════════

def format_value(value, sqlite_type: str) -> str:
    """Return a SQL literal for the value."""
    if value is None:
        return "NULL"

    st = sqlite_type.upper()

    if "INTEGER" in st:
        # Boolean: Access stores as -1 (True) / 0 (False)
        if isinstance(value, bool):
            return "1" if value else "0"
        try:
            v = int(value)
            return "1" if v == -1 else str(v)   # -1 → True in Access
        except (ValueError, TypeError):
            return "NULL"

    if "REAL" in st:
        try:
            return repr(float(value))
        except (ValueError, TypeError):
            return "NULL"

    if "BLOB" in st:
        if isinstance(value, (bytes, bytearray)):
            return f"X'{value.hex()}'"
        return "NULL"

    # TEXT / DATETIME
    if isinstance(value, datetime):
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    if hasattr(value, "isoformat"):          # date / time objects
        return f"'{value.isoformat()}'"

    # Escape single quotes
    escaped = str(value).replace("'", "''")
    return f"'{escaped}'"


# ══════════════════════════════════════════════════════════════════════════════
# Backend: pyodbc (Windows / macOS with mdbtools ODBC driver)
# ══════════════════════════════════════════════════════════════════════════════

def try_pyodbc(accdb: Path):
    """Return (columns_info_per_table, rows_per_table) or raise ImportError."""
    import pyodbc  # noqa: F401 – imported to verify availability

    if IS_WINDOWS:
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"Dbq={accdb};"
        )
    else:
        # macOS/Linux with mdbtools ODBC
        conn_str = f"Driver={{MDBTools}};Dbq={accdb};"

    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()

    tables = [row.table_name for row in cursor.tables(tableType="TABLE")]

    schema: dict[str, list[dict]] = {}
    data:   dict[str, list[list]] = {}

    for table in tables:
        cols_info = []
        cursor.execute(f"SELECT * FROM [{table}] WHERE 1=0")
        for col in cursor.description:
            name      = col[0]
            type_code = col[1]
            # Map pyodbc type_code to Access-like name
            import pyodbc as _p
            if type_code == _p.SQL_TYPE_DATE or type_code == _p.SQL_TYPE_TIMESTAMP:
                atype = "date/time"
            elif type_code in (_p.SQL_SMALLINT, _p.SQL_INTEGER, _p.SQL_BIGINT, _p.SQL_TINYINT):
                atype = "integer"
            elif type_code in (_p.SQL_FLOAT, _p.SQL_REAL, _p.SQL_DOUBLE, _p.SQL_NUMERIC, _p.SQL_DECIMAL):
                atype = "double"
            elif type_code == _p.SQL_BIT:
                atype = "yes/no"
            elif type_code in (_p.SQL_BINARY, _p.SQL_VARBINARY, _p.SQL_LONGVARBINARY):
                atype = "ole object"
            else:
                atype = "text"
            cols_info.append({"name": name, "access_type": atype,
                               "sqlite_type": access_type_to_sqlite(atype)})
        schema[table] = cols_info

        cursor.execute(f"SELECT * FROM [{table}]")
        data[table] = [list(row) for row in cursor.fetchall()]

    conn.close()
    return schema, data


# ══════════════════════════════════════════════════════════════════════════════
# Backend: mdbtools CLI (macOS / Linux)
# ══════════════════════════════════════════════════════════════════════════════

def _run(cmd: list[str], check=True) -> str:
    result = subprocess.run(cmd, capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
    if check and result.returncode != 0:
        raise RuntimeError(f"Command {cmd} failed:\n{result.stderr}")
    return result.stdout


def _mdbtools_available() -> bool:
    return subprocess.run(["which", "mdb-tables"], capture_output=True).returncode == 0


# mdbtools schema type string patterns
_MDB_SCHEMA_TYPE_RE = re.compile(
    r"^\s*`?(\w+)`?\s+([\w /]+?)(?:\s*\([^)]*\))?\s*(?:NOT NULL|NULL|,|$)",
    re.IGNORECASE,
)

def _parse_mdb_schema_type(col_line: str) -> str:
    """Extract Access column type from a mdb-schema DDL line."""
    # mdb-schema output looks like:  `ColumnName`  Text (255),
    # or: `DateField`  DateTime,
    m = re.match(r"\s*`?[^`\s]+`?\s+([\w /]+)", col_line)
    if m:
        return m.group(1).strip()
    return "text"


def try_mdbtools(accdb: Path):
    """Extract schema + data using mdbtools CLI utilities."""
    db = str(accdb)

    # List tables
    raw_tables = _run(["mdb-tables", "-1", db])
    tables = [t.strip() for t in raw_tables.splitlines() if t.strip()]

    # Get DDL for type info
    ddl = _run(["mdb-schema", db, "mysql"])   # mysql dialect closest to SQL

    schema: dict[str, list[dict]] = {}
    data:   dict[str, list[list]] = {}

    # Parse DDL to extract column types per table
    # mdb-schema output: CREATE TABLE `TableName` (\n  `Col` Type(len),\n ...);
    table_ddl_re = re.compile(
        r"CREATE\s+TABLE\s+`?([^`\s(]+)`?\s*\((.+?)\);",
        re.IGNORECASE | re.DOTALL,
    )
    ddl_types: dict[str, dict[str, str]] = {}
    for m in table_ddl_re.finditer(ddl):
        tname = m.group(1)
        body  = m.group(2)
        col_types: dict[str, str] = {}
        for line in body.splitlines():
            line = line.strip().rstrip(",")
            if not line or line.upper().startswith(("PRIMARY", "UNIQUE", "INDEX", "KEY")):
                continue
            cm = re.match(r"`?([^`\s]+)`?\s+([\w /]+)", line)
            if cm:
                cname = cm.group(1)
                ctype = cm.group(2).strip()
                col_types[cname] = ctype
        ddl_types[tname] = col_types

    for table in tables:
        # Export data as CSV
        csv_out = _run(["mdb-export", "-H", db, table], check=False)
        # With -H (no header), first run without -H to get header
        csv_hdr = _run(["mdb-export", db, table], check=False)

        header_line = csv_hdr.splitlines()[0] if csv_hdr.strip() else ""
        col_names = _parse_csv_line(header_line) if header_line else []

        # Build schema from DDL types
        type_map = ddl_types.get(table, {})
        cols_info = []
        for cname in col_names:
            atype = type_map.get(cname, "text")
            cols_info.append({
                "name":        cname,
                "access_type": atype,
                "sqlite_type": access_type_to_sqlite(atype),
            })
        schema[table] = cols_info

        # Parse data rows
        rows = []
        lines = csv_out.splitlines()
        for line in lines:
            if not line.strip():
                continue
            raw_vals = _parse_csv_line(line)
            typed_vals = []
            for i, v in enumerate(raw_vals):
                col = cols_info[i] if i < len(cols_info) else {"sqlite_type": "TEXT COLLATE NOCASE"}
                typed_vals.append(_coerce_value(v, col["sqlite_type"]))
            rows.append(typed_vals)
        data[table] = rows

    return schema, data


def _parse_csv_line(line: str) -> list[str]:
    """Simple CSV parser that handles quoted fields with embedded commas/newlines."""
    import csv
    import io
    reader = csv.reader(io.StringIO(line))
    try:
        return next(reader)
    except StopIteration:
        return []


def _coerce_value(raw: str, sqlite_type: str) -> object:
    """Convert a raw CSV string to a Python value appropriate for the type."""
    if raw == "" or raw.lower() == "null":
        return None
    st = sqlite_type.upper()
    if "INTEGER" in st:
        try:
            return int(raw)
        except ValueError:
            # Boolean textual values from mdbtools
            if raw.lower() in ("true", "yes", "1"):
                return 1
            if raw.lower() in ("false", "no", "0"):
                return 0
            return None
    if "REAL" in st:
        try:
            return float(raw)
        except ValueError:
            return None
    if "BLOB" in st:
        return None  # binary not reliably exported via CSV
    # TEXT / datetime — keep as string
    return raw


# ══════════════════════════════════════════════════════════════════════════════
# SQL generation
# ══════════════════════════════════════════════════════════════════════════════

def quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def build_create_table(table: str, cols: list[dict]) -> str:
    col_defs = []
    for col in cols:
        col_defs.append(f"  {quote_ident(col['name'])} {col['sqlite_type']}")
    return (
        f"CREATE TABLE IF NOT EXISTS {quote_ident(table)} (\n"
        + ",\n".join(col_defs)
        + "\n);"
    )


def build_insert(table: str, cols: list[dict], rows: list[list]) -> list[str]:
    if not rows:
        return []
    col_names = ", ".join(quote_ident(c["name"]) for c in cols)
    stmts = []
    for row in rows:
        values = []
        for i, col in enumerate(cols):
            val = row[i] if i < len(row) else None
            values.append(format_value(val, col["sqlite_type"]))
        stmts.append(
            f"INSERT INTO {quote_ident(table)} ({col_names}) VALUES ({', '.join(values)});"
        )
    return stmts


# ══════════════════════════════════════════════════════════════════════════════
# Per-database export
# ══════════════════════════════════════════════════════════════════════════════

def export_db(accdb: Path, use_pyodbc: bool) -> None:
    stem      = accdb.stem
    out_dir   = accdb.parent
    sql_path  = unique_output_path(out_dir / f"{stem}.sql")

    print(f"\n  → Extracting: {accdb}")

    try:
        if use_pyodbc:
            schema, data = try_pyodbc(accdb)
        else:
            schema, data = try_mdbtools(accdb)
    except Exception as e:
        print(f"    ERROR reading {accdb.name}: {e}")
        return

    sql_lines = [
        "-- Generated by access2sql.py",
        f"-- Source: {accdb}",
        f"-- Date  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "PRAGMA journal_mode=WAL;",
        "PRAGMA foreign_keys=ON;",
        "",
    ]

    for table, cols in schema.items():
        create_stmt = build_create_table(table, cols)
        sql_lines.append(create_stmt)
        sql_lines.append("")

        inserts = build_insert(table, cols, data.get(table, []))
        sql_lines.extend(inserts)
        if inserts:
            sql_lines.append("")

    sql_path.write_text("\n".join(sql_lines), encoding="utf-8")
    print(f"    ✓ SQL    → {sql_path.relative_to(accdb.parent.parent) if accdb.parent.parent != accdb.parent else sql_path.name}")


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("Access → SQLite extractor")
    print("=" * 50)

    # Determine backend
    use_pyodbc = False
    try:
        import pyodbc  # noqa: F401
        use_pyodbc = True
        print("Backend: pyodbc")
    except ImportError:
        if IS_MAC and _mdbtools_available():
            print("Backend: mdbtools CLI  (brew install mdbtools)")
        elif IS_WINDOWS:
            print("ERROR: pyodbc not found. Install with: pip install pyodbc")
            sys.exit(1)
        else:
            print("ERROR: Neither pyodbc nor mdbtools found.")
            print("  macOS : brew install mdbtools")
            print("  Windows: pip install pyodbc")
            sys.exit(1)

    # Folder picker
    root_folder = pick_folder()
    print(f"Root folder: {root_folder}")

    # Find all .accdb / .mdb files
    db_files = find_access_files(root_folder)
    if not db_files:
        print("No .accdb or .mdb files found.")
        sys.exit(0)

    print(f"Found {len(db_files)} database(s):\n")
    for f in db_files:
        print(f"  {f.relative_to(root_folder)}")

    # Export each
    for accdb in db_files:
        export_db(accdb, use_pyodbc)

    print("\nDone.")


if __name__ == "__main__":
    main()
