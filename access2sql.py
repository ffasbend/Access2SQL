#!/usr/bin/env python3
"""
access2sql.py
---------------------------
Use a system folder-picker dialog to choose a root folder, then recursively
find every .accdb / .mdb file, extract all tables + data, and produce:
  - <db_name>.sql   : CREATE TABLE + INSERT statements (SQLite-compatible)

Type mapping from Access → SQLite:
  Text / Memo / Hyperlink  → TEXT COLLATE NOCASE
  Date/Time                → DATETIME
  Yes/No (Boolean)         → BOOLEAN (0/1)
  AutoNumber / Long Integer → INTEGER
  Integer / Number        → INTEGER
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
from typing import Any


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
    "int":        "INTEGER",
    "smallint":   "INTEGER",
    "bigint":     "INTEGER",
    "tinyint":    "INTEGER",
    "byte":       "INTEGER",
    "number":     "INTEGER",
    "single":     "REAL",
    "double":     "REAL",
    "float":      "REAL",
    "currency":   "REAL",
    "decimal":    "REAL",
    "numeric":    "REAL",
    # Date
    "date/time":  "DATETIME",
    "datetime":   "DATETIME",
    # Boolean
    "yes/no":     "BOOLEAN",
    "boolean":    "BOOLEAN",
    "bit":        "BOOLEAN",
    # Binary
    "ole object": "BLOB",
    "binary":     "BLOB",
    "varbinary":  "BLOB",
}

def access_type_to_sqlite(access_type: str) -> str:
    key = access_type.lower().strip()
    key = key.split("(", 1)[0].strip()
    key = re.sub(r"\s+(not\s+null|null)$", "", key).strip()
    return _TYPE_MAP.get(key, "TEXT COLLATE NOCASE")


# ══════════════════════════════════════════════════════════════════════════════
# Value formatting for INSERT statements
# ══════════════════════════════════════════════════════════════════════════════

def format_value(value, sqlite_type: str) -> str:
    """Return a SQL literal for the value."""
    if value is None:
        return "NULL"

    st = sqlite_type.upper()

    if "INTEGER" in st or "BOOLEAN" in st:
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
    if "DATETIME" in st:
        mode = "datetime"
        if isinstance(value, dict):
            # Defensive fallback for accidental dict payloads
            escaped = str(value).replace("'", "''")
            return f"'{escaped}'"
        return format_datetime_value(value, mode)

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

    schema: dict[str, dict[str, object]] = {}
    data:   dict[str, list[list]] = {}

    for table in tables:
        cols_info = []
        primary_key: list[str] = []
        foreign_keys: list[dict[str, object]] = []
        nullable_by_name: dict[str, bool] = {}

        try:
            for row in cursor.columns(table=table):
                column_name = getattr(row, "column_name", None)
                nullable = getattr(row, "nullable", None)
                if column_name is not None and nullable is not None:
                    nullable_by_name[column_name] = bool(nullable)
        except Exception:
            nullable_by_name = {}

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
            cols_info.append({
                "name": name,
                "access_type": atype,
                "sqlite_type": access_type_to_sqlite(atype),
                "not_null": not nullable_by_name.get(name, True),
                "datetime_mode": "auto" if "DATETIME" in access_type_to_sqlite(atype).upper() else None,
                "datetime_include_seconds": None if "DATETIME" in access_type_to_sqlite(atype).upper() else None,
            })

        try:
            pk_rows = sorted(
                cursor.primaryKeys(table=table),
                key=lambda row: getattr(row, "key_seq", 0),
            )
            primary_key = [row.column_name for row in pk_rows if getattr(row, "column_name", None)]
        except Exception:
            primary_key = []

        primary_key_set = set(primary_key)
        for col in cols_info:
            if col["name"] in primary_key_set:
                col["not_null"] = True

        try:
            fk_groups: dict[object, list] = {}
            for row in cursor.foreignKeys(table=table):
                key = getattr(row, "fk_name", None) or (
                    getattr(row, "pktable_name", None),
                    getattr(row, "fktable_name", None),
                )
                fk_groups.setdefault(key, []).append(row)
            for rows in fk_groups.values():
                rows = sorted(rows, key=lambda row: getattr(row, "key_seq", 0))
                fk_columns = [row.fkcolumn_name for row in rows if getattr(row, "fkcolumn_name", None)]
                ref_columns = [row.pkcolumn_name for row in rows if getattr(row, "pkcolumn_name", None)]
                ref_table = getattr(rows[0], "pktable_name", None)
                if fk_columns and ref_columns and ref_table:
                    on_update = _odbc_rule_to_action(getattr(rows[0], "update_rule", None))
                    on_delete = _odbc_rule_to_action(getattr(rows[0], "delete_rule", None))
                    foreign_keys.append({
                        "columns": fk_columns,
                        "ref_table": ref_table,
                        "ref_columns": ref_columns,
                        "on_update": on_update,
                        "on_delete": on_delete,
                    })
        except Exception:
            foreign_keys = []

        schema[table] = {
            "columns": cols_info,
            "primary_key": primary_key,
            "foreign_keys": foreign_keys,
        }

        cursor.execute(f"SELECT * FROM [{table}]")
        data[table] = [list(row) for row in cursor.fetchall()]

    infer_datetime_modes(schema, data)

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

    schema: dict[str, dict[str, object]] = {}
    data:   dict[str, list[list]] = {}

    # Parse DDL to extract column types per table.
    # Depending on the selected dialect, mdb-schema may emit primary keys either
    # inline inside CREATE TABLE or later as ALTER TABLE ... ADD PRIMARY KEY.
    table_ddl_re = re.compile(
        r"CREATE\s+TABLE\s+`?([^`\s(]+)`?\s*\((.+?)\);",
        re.IGNORECASE | re.DOTALL,
    )
    alter_pk_re = re.compile(
        r"ALTER\s+TABLE\s+`?([^`\s(]+)`?\s+ADD\s+PRIMARY\s+KEY\s*\((.+?)\);",
        re.IGNORECASE,
    )
    ddl_types: dict[str, dict[str, str]] = {}
    ddl_primary_keys: dict[str, list[str]] = {}
    ddl_foreign_keys: dict[str, list[dict[str, object]]] = {}
    alter_fk_re = re.compile(
        r"ALTER\s+TABLE\s+`?([^`\s(]+)`?\s+ADD\s+CONSTRAINT\s+`?([^`\s]+)`?\s+"
        r"FOREIGN\s+KEY\s*\((.+?)\)\s+REFERENCES\s+`?([^`\s(]+)`?\s*\((.+?)\)\s*(.*?);",
        re.IGNORECASE,
    )
    for m in table_ddl_re.finditer(ddl):
        tname = m.group(1)
        body  = m.group(2)
        col_types: dict[str, str] = {}
        col_not_null: dict[str, bool] = {}
        primary_key: list[str] = []
        for line in body.splitlines():
            line = line.strip().rstrip(",")
            if not line:
                continue
            normalized_line = line.lstrip(", ")
            upper_line = normalized_line.upper()
            if upper_line.startswith("PRIMARY KEY"):
                pk_match = re.search(r"\((.+)\)", normalized_line)
                if pk_match:
                    primary_key = [part.strip().strip("`").strip('"') for part in pk_match.group(1).split(",")]
                continue
            if upper_line.startswith(("UNIQUE", "INDEX", "KEY")):
                continue
            cm = re.match(r"`?([^`\s]+)`?\s+([\w /]+)", normalized_line)
            if cm:
                cname = cm.group(1)
                ctype = cm.group(2).strip()
                col_types[cname] = ctype
                col_not_null[cname] = "NOT NULL" in upper_line
        ddl_types[tname] = col_types
        ddl_not_null = ddl_foreign_keys.setdefault("__not_null__", {})
        ddl_not_null[tname] = col_not_null
        ddl_primary_keys[tname] = primary_key
        ddl_foreign_keys.setdefault(tname, [])

    for m in alter_pk_re.finditer(ddl):
        tname = m.group(1)
        ddl_primary_keys[tname] = [
            part.strip().strip("`").strip('"')
            for part in m.group(2).split(",")
        ]

    for m in alter_fk_re.finditer(ddl):
        tname = m.group(1)
        fk_columns = [part.strip().strip("`").strip('"') for part in m.group(3).split(",")]
        ref_table = m.group(4).strip().strip("`").strip('"')
        ref_columns = [part.strip().strip("`").strip('"') for part in m.group(5).split(",")]
        tail = m.group(6) or ""
        on_update = _parse_fk_action_from_tail(tail, "update")
        on_delete = _parse_fk_action_from_tail(tail, "delete")
        ddl_foreign_keys.setdefault(tname, []).append({
            "columns": fk_columns,
            "ref_table": ref_table,
            "ref_columns": ref_columns,
            "on_update": on_update,
            "on_delete": on_delete,
        })

    merged_foreign_keys = merge_foreign_keys(ddl_foreign_keys, read_msys_relationships(accdb))

    for table in tables:
        datetime_mode_map = read_datetime_modes_from_mdb_prop(accdb, table)

        # Export data as CSV
        csv_out = _run(["mdb-export", "-H", db, table], check=False)
        # With -H (no header), first run without -H to get header
        csv_hdr = _run(["mdb-export", db, table], check=False)

        header_line = csv_hdr.splitlines()[0] if csv_hdr.strip() else ""
        col_names = _parse_csv_line(header_line) if header_line else []

        # Build schema from DDL types
        type_map = ddl_types.get(table, {})
        not_null_map = ddl_foreign_keys.get("__not_null__", {}).get(table, {})
        primary_key_set = set(ddl_primary_keys.get(table, []))
        cols_info = []
        for cname in col_names:
            atype = type_map.get(cname, "text")
            sqlite_type = access_type_to_sqlite(atype)
            cols_info.append({
                "name":        cname,
                "access_type": atype,
                "sqlite_type": sqlite_type,
                "not_null":    not_null_map.get(cname, False) or cname in primary_key_set,
                "datetime_mode": datetime_mode_map.get(cname, {}).get("mode", "auto") if "DATETIME" in sqlite_type.upper() else None,
                "datetime_include_seconds": datetime_mode_map.get(cname, {}).get("include_seconds") if "DATETIME" in sqlite_type.upper() else None,
            })
        schema[table] = {
            "columns": cols_info,
            "primary_key": ddl_primary_keys.get(table, []),
            "foreign_keys": merged_foreign_keys.get(table, []),
        }

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

    infer_datetime_modes(schema, data)

    return schema, data


def _normalize_access_format(fmt: str) -> str:
    return re.sub(r"\s+", " ", (fmt or "").strip().lower())


def datetime_settings_from_access_format(fmt: str | None) -> dict[str, object]:
    """Map Access Format text to datetime mode and time precision."""
    settings: dict[str, object] = {"mode": None, "include_seconds": None}
    if not fmt:
        return settings
    f = _normalize_access_format(fmt)
    if not f:
        return settings

    # Common Access named formats
    if "short date" in f or "medium date" in f or "long date" in f:
        settings["mode"] = "date"
        return settings
    if "short time" in f:
        settings["mode"] = "time"
        settings["include_seconds"] = False
        return settings
    if "medium time" in f:
        settings["mode"] = "time"
        settings["include_seconds"] = False
        return settings
    if "long time" in f:
        settings["mode"] = "time"
        settings["include_seconds"] = True
        return settings
    if "general date" in f:
        settings["mode"] = "datetime"
        return settings

    has_date_token = bool(re.search(r"\b(d|dd|ddd|dddd|m|mm|mmm|mmmm|yy|yyyy)\b", f))
    has_time_token = bool(re.search(r"\b(h|hh|n|nn|s|ss|am/pm|a/p)\b", f))
    has_second_token = bool(re.search(r"\b(s|ss)\b", f))
    if has_date_token and not has_time_token:
        settings["mode"] = "date"
    elif has_time_token and not has_date_token:
        settings["mode"] = "time"
        settings["include_seconds"] = has_second_token
    elif has_date_token and has_time_token:
        settings["mode"] = "datetime"
        settings["include_seconds"] = has_second_token
    return settings


def read_datetime_modes_from_mdb_prop(accdb: Path, table: str) -> dict[str, dict[str, object]]:
    """Read Access field format settings and return datetime settings per column."""
    try:
        text = _run(["mdb-prop", str(accdb), table], check=False)
    except Exception:
        return {}
    if not text.strip():
        return {}

    mode_map: dict[str, dict[str, object]] = {}
    current_col: str | None = None
    current_props: dict[str, str] = {}

    def flush() -> None:
        nonlocal current_col, current_props
        if not current_col or current_col == "(none)":
            return
        fmt = current_props.get("Format")
        settings = datetime_settings_from_access_format(fmt)
        if settings.get("mode"):
            mode_map[current_col] = settings

    for line in text.splitlines():
        m_name = re.match(r"^name:\s*(.+)\s*$", line)
        if m_name:
            flush()
            current_col = m_name.group(1).strip()
            current_props = {}
            continue
        m_prop = re.match(r"^\s+([^:]+):\s*(.*)$", line)
        if m_prop and current_col is not None:
            key = m_prop.group(1).strip()
            value = m_prop.group(2).strip()
            current_props[key] = value
    flush()
    return mode_map


def read_msys_relationships(accdb: Path) -> dict[str, list[dict[str, object]]]:
    """Read relationship metadata from MSysRelationships, including cascade flags."""
    try:
        text = _run(["mdb-export", str(accdb), "MSysRelationships"], check=False)
    except Exception:
        return {}
    if not text.strip():
        return {}

    import csv
    import io

    by_name: dict[str, list[dict[str, str]]] = {}
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        rel_name = (row.get("szRelationship") or "").strip()
        if not rel_name:
            continue
        by_name.setdefault(rel_name, []).append(row)

    relationships: dict[str, list[dict[str, object]]] = {}
    for rel_rows in by_name.values():
        rel_rows.sort(key=lambda row: int((row.get("icolumn") or row.get("ccolumn") or "0") or "0"))
        sample = rel_rows[0]
        child_table = (sample.get("szObject") or "").strip()
        parent_table = (sample.get("szReferencedObject") or "").strip()
        child_columns = [(row.get("szColumn") or "").strip() for row in rel_rows]
        parent_columns = [(row.get("szReferencedColumn") or "").strip() for row in rel_rows]
        if not child_table or not parent_table or not all(child_columns) or not all(parent_columns):
            continue

        grbit_raw = (sample.get("grbit") or "0").strip()
        try:
            grbit = int(grbit_raw)
        except ValueError:
            grbit = 0

        relationships.setdefault(child_table, []).append({
            "columns": child_columns,
            "ref_table": parent_table,
            "ref_columns": parent_columns,
            "on_update": "CASCADE" if grbit & 256 else None,
            "on_delete": "CASCADE" if grbit & 4096 else None,
        })

    return relationships


def merge_foreign_keys(
    ddl_foreign_keys: dict[str, list[dict[str, object]]],
    relationship_foreign_keys: dict[str, list[dict[str, object]]],
) -> dict[str, list[dict[str, object]]]:
    """Merge FK metadata, preserving parsed keys and supplementing from system tables."""
    merged = {
        table: list(fks)
        for table, fks in ddl_foreign_keys.items()
        if table != "__not_null__"
    }

    for table, rel_fks in relationship_foreign_keys.items():
        bucket = merged.setdefault(table, [])
        for rel_fk in rel_fks:
            existing = None
            for fk in bucket:
                if (
                    fk.get("columns") == rel_fk.get("columns")
                    and fk.get("ref_table") == rel_fk.get("ref_table")
                    and fk.get("ref_columns") == rel_fk.get("ref_columns")
                ):
                    existing = fk
                    break
            if existing is None:
                bucket.append(dict(rel_fk))
                continue
            if rel_fk.get("on_update") == "CASCADE":
                existing["on_update"] = "CASCADE"
            if rel_fk.get("on_delete") == "CASCADE":
                existing["on_delete"] = "CASCADE"

    return merged


def _parse_datetime_like(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    patterns = [
        "%m/%d/%y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%H:%M:%S",
    ]
    for pat in patterns:
        try:
            return datetime.strptime(s, pat)
        except ValueError:
            continue
    return None


def infer_datetime_modes(schema: dict[str, dict[str, object]], data: dict[str, list[list]]) -> None:
    """Infer date/time mode and time precision when metadata is absent."""
    baseline_dates = {
        datetime(1899, 12, 30).date(),
        datetime(1899, 12, 31).date(),
        datetime(1900, 1, 1).date(),
    }
    for table, table_schema in schema.items():
        cols = table_schema.get("columns", [])
        rows = data.get(table, [])
        for idx, col in enumerate(cols):
            if "DATETIME" not in str(col.get("sqlite_type", "")).upper():
                continue
            mode = col.get("datetime_mode")
            include_seconds = col.get("datetime_include_seconds")
            need_mode = mode in (None, "", "auto")
            need_seconds = include_seconds is None
            if not need_mode and not need_seconds:
                continue

            parsed_values: list[datetime] = []
            for row in rows:
                if idx >= len(row):
                    continue
                dt = _parse_datetime_like(row[idx])
                if dt is not None:
                    parsed_values.append(dt)
            if not parsed_values:
                if need_mode:
                    col["datetime_mode"] = "datetime"
                if need_seconds:
                    col["datetime_include_seconds"] = True
                continue

            if need_mode:
                if all(dt.time().strftime("%H:%M:%S") == "00:00:00" for dt in parsed_values):
                    col["datetime_mode"] = "date"
                elif all(dt.date() in baseline_dates for dt in parsed_values):
                    col["datetime_mode"] = "time"
                else:
                    col["datetime_mode"] = "datetime"

            if need_seconds:
                col["datetime_include_seconds"] = any(dt.second != 0 for dt in parsed_values)


def format_datetime_value(value: Any, mode: str, include_seconds: bool | None = None) -> str:
    dt = _parse_datetime_like(value)
    if dt is None:
        escaped = str(value).replace("'", "''")
        return f"'{escaped}'"
    if mode == "date":
        return f"'{dt.strftime('%Y-%m-%d')}'"
    if mode == "time":
        return f"'{dt.strftime('%H:%M:%S' if include_seconds else '%H:%M')}'"
    return f"'{dt.strftime('%Y-%m-%d %H:%M:%S' if include_seconds else '%Y-%m-%d %H:%M')}'"


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
    if "INTEGER" in st or "BOOLEAN" in st:
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


def build_create_table(
    table: str,
    cols: list[dict],
    primary_key: list[str] | None = None,
    foreign_keys: list[dict[str, object]] | None = None,
) -> str:
    col_defs = []
    for col in cols:
        nullability = "NOT NULL" if col.get("not_null") else "NULL"
        col_defs.append(f"  {quote_ident(col['name'])} {col['sqlite_type']} {nullability}")
    if primary_key:
        pk_cols = ", ".join(quote_ident(col) for col in primary_key)
        col_defs.append(f"  PRIMARY KEY ({pk_cols})")
    for foreign_key in foreign_keys or []:
        fk_cols = ", ".join(quote_ident(col) for col in foreign_key["columns"])
        ref_cols = ", ".join(quote_ident(col) for col in foreign_key["ref_columns"])
        ref_table = quote_ident(foreign_key["ref_table"])
        fk_clause = f"  FOREIGN KEY ({fk_cols}) REFERENCES {ref_table} ({ref_cols})"
        if foreign_key.get("on_delete") == "CASCADE":
            fk_clause += " ON DELETE CASCADE"
        if foreign_key.get("on_update") == "CASCADE":
            fk_clause += " ON UPDATE CASCADE"
        col_defs.append(fk_clause)
    return (
        f"CREATE TABLE IF NOT EXISTS {quote_ident(table)} (\n"
        + ",\n".join(col_defs)
        + "\n);"
    )


def _odbc_rule_to_action(rule: Any) -> str | None:
    """Map ODBC foreign key rule value to action name."""
    if rule is None:
        return None
    if isinstance(rule, str):
        value = rule.strip().upper().replace(" ", "_")
        if value in {"CASCADE", "NO_ACTION", "RESTRICT", "SET_NULL", "SET_DEFAULT"}:
            return value
        return None
    try:
        num = int(rule)
    except (ValueError, TypeError):
        return None

    # ODBC constants: SQL_CASCADE=0, SQL_RESTRICT=1, SQL_SET_NULL=2,
    # SQL_NO_ACTION=3, SQL_SET_DEFAULT=4
    mapping = {
        0: "CASCADE",
        1: "RESTRICT",
        2: "SET_NULL",
        3: "NO_ACTION",
        4: "SET_DEFAULT",
    }
    return mapping.get(num)


def _parse_fk_action_from_tail(tail: str, action: str) -> str | None:
    """Parse ON UPDATE/ON DELETE action from FK tail SQL."""
    m = re.search(rf"ON\s+{action}\s+(CASCADE|RESTRICT|NO\s+ACTION|SET\s+NULL|SET\s+DEFAULT)", tail, re.IGNORECASE)
    if not m:
        return None
    return m.group(1).upper().replace(" ", "_")


def build_insert(table: str, cols: list[dict], rows: list[list]) -> list[str]:
    if not rows:
        return []
    col_names = ", ".join(quote_ident(c["name"]) for c in cols)
    stmts = []
    for row in rows:
        values = []
        for i, col in enumerate(cols):
            val = row[i] if i < len(row) else None
            if "DATETIME" in str(col.get("sqlite_type", "")).upper():
                values.append(
                    format_datetime_value(
                        val,
                        str(col.get("datetime_mode") or "datetime"),
                        col.get("datetime_include_seconds"),
                    )
                )
            else:
                values.append(format_value(val, col["sqlite_type"]))
        stmts.append(
            f"INSERT INTO {quote_ident(table)} ({col_names}) VALUES ({', '.join(values)});"
        )
    return stmts


def order_tables_by_dependencies(schema: dict[str, dict[str, object]]) -> list[str]:
    """Return tables ordered so referenced tables come before dependent tables."""
    table_names = list(schema.keys())
    remaining = set(table_names)
    dependencies: dict[str, set[str]] = {}

    for table in table_names:
        foreign_keys = schema[table].get("foreign_keys", [])
        refs = {
            fk["ref_table"]
            for fk in foreign_keys
            if fk.get("ref_table") in schema and fk.get("ref_table") != table
        }
        dependencies[table] = refs

    ordered: list[str] = []
    while remaining:
        ready = [table for table in table_names if table in remaining and not dependencies[table]]
        if not ready:
            # Cycles or unresolved metadata: keep stable original order for the rest.
            ordered.extend([table for table in table_names if table in remaining])
            break

        for table in ready:
            ordered.append(table)
            remaining.remove(table)
        for deps in dependencies.values():
            deps.difference_update(ready)

    return ordered


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
        "-- Generated by extract_access_to_sqlite.py",
        f"-- Source: {accdb}",
        f"-- Date  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "PRAGMA journal_mode=WAL;",
        "PRAGMA foreign_keys=ON;",
        "",
    ]

    table_order = order_tables_by_dependencies(schema)

    for table in table_order:
        table_schema = schema[table]
        cols = table_schema["columns"]
        primary_key = table_schema.get("primary_key", [])
        foreign_keys = table_schema.get("foreign_keys", [])
        create_stmt = build_create_table(table, cols, primary_key, foreign_keys)
        sql_lines.append(create_stmt)
        sql_lines.append("")

    for table in table_order:
        table_schema = schema[table]
        cols = table_schema["columns"]

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
