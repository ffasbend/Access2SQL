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

## How to install

You can use either Conda (recommended for easiest cross-machine setup) or pip.

### Option A: Conda (recommended)

1. Create the environment:

  conda env create -f environment.yml

2. Activate it:

  conda activate access2sql

3. Run the script:

  python access2sql.py

### Option B: pip + system packages

1. Install Python dependency:

  pip install -r requirements.txt

2. Install required system tools:

   - macOS (Homebrew): install mdbtools and unixodbc
   - Linux: install mdbtools and unixODBC from your package manager
   - Windows: install Microsoft Access Database Engine (ODBC driver)

3. Run the script:

  python access2sql.py

Notes:

- tkinter is included with standard Python distributions on most systems.
- On macOS/Linux, the script can use mdbtools CLI fallback even if pyodbc is not available.

This script auto-detects the platform and picks the right backend.
On macOS it falls back to the `mdbtools` CLI utilities (mdb-tables, mdb-export,
mdb-schema) if pyodbc is unavailable.