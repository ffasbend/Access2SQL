# access2sql.py

Use a system folder-picker dialog to choose a root folder, then recursively
find every .accdb / .mdb file, extract all tables + data, and produce:
  - <db_name>.sql   : CREATE TABLE + INSERT statements (SQLite-compatible)
  - <db_name>.sqlite: ready-to-use SQLite database

Type mapping from Access → SQLite:

  -  Text / Memo / Hyperlink  → TEXT COLLATE NOCASE
  -  Date/Time                → TEXT (stored as ISO-8601, cast with datetime())
  -  Yes/No (Boolean)         → INTEGER (0/1)  -- SQLite has no BOOLEAN type
  -  AutoNumber / Long Integer → INTEGER
  -  Integer                  → INTEGER
  -  Single / Double / Decimal → REAL
  -  Currency                  → CURRENCY (treated as REAL)
  -  OLE Object / Binary      → BLOB
  -  everything else          → TEXT COLLATE NOCASE

Requirements (install once):
  ```
  pip install pyodbc     # on macOS needs mdbtools (brew install mdbtools)
  ```
  -- OR --
  ```
  pip install pywin32    # Windows only (uses COM / DAO)
  ```

This script auto-detects the platform and picks the right backend.
On macOS it falls back to the `mdbtools` CLI utilities (mdb-tables, mdb-export,
mdb-schema) if pyodbc is unavailable.