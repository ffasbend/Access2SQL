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

---

## GUI

A graphical interface is available in `access2sql_gui.py`.  
Drop `.accdb` / `.mdb` files or folders onto the window, choose TXT or MD output
for the query export, and click **Generate Output**.

Run it directly (no build needed):

```bash
pip install customtkinter tkinterdnd2
python access2sql_gui.py
```

---

## Building a standalone app (no Python required on target machine)

### macOS

**Prerequisites** (developer machine only):

```bash
brew install mdbtools
pip install pyinstaller customtkinter tkinterdnd2
xcode-select --install        # provides sips, iconutil, otool, install_name_tool
```

**Build:**

```bash
python build_mac.py
```

This stages the mdbtools binaries with all their dylib dependencies, fixes rpath
references so they resolve inside the bundle, then runs PyInstaller.

Output: `dist/Access2SQL.app` — copy it to `/Applications` or distribute as-is.

---

### Linux

**Prerequisites:**

```bash
sudo apt install mdbtools        # Debian/Ubuntu; use dnf/pacman on other distros
pip install pyinstaller customtkinter tkinterdnd2
```

**Build:**

```bash
pyinstaller --noconfirm --clean \
    --name Access2SQL \
    --collect-all tkinterdnd2 \
    --collect-all customtkinter \
    --add-data "assets:assets" \
    --icon assets/icon.icns \
    access2sql_gui.py
```

Output: `dist/Access2SQL/` folder containing a self-contained executable.  
Run with `dist/Access2SQL/Access2SQL` or package the folder as a `.tar.gz` for
distribution.

> **Note:** mdbtools must be installed on the target machine separately on Linux —
> it cannot be bundled the same way as on macOS because it links against system
> libraries that vary by distribution.

---

### Windows

**Prerequisites:**

- Install [Microsoft Access Database Engine 2016](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (provides the ODBC driver)
- Install Python 3.10+

```cmd
pip install pyinstaller customtkinter tkinterdnd2 pyodbc
```

**Convert the icon** to `.ico` first (use any PNG→ICO converter, or Pillow):

```python
from PIL import Image
Image.open("assets/icon.png").save("assets/icon.ico", sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])
```

**Build:**

```cmd
pyinstaller --noconfirm --clean ^
    --name Access2SQL ^
    --windowed ^
    --collect-all tkinterdnd2 ^
    --collect-all customtkinter ^
    --add-data "assets;assets" ^
    --icon assets\icon.ico ^
    access2sql_gui.py
```

Output: `dist\Access2SQL\Access2SQL.exe` (or use `--onefile` for a single `.exe`).

> **Note:** The `--add-data` separator is `;` on Windows (`:` on macOS/Linux).

---

## File overview

| File | Purpose |
|---|---|
| `access2sql.py` | Core CLI — table + query export |
| `access2sql_gui.py` | GUI frontend (customtkinter) |
| `build_mac.py` | macOS build script (bundles mdbtools + PyInstaller) |
| `assets/icon.png` | App icon (source PNG) |
| `assets/icon.icns` | App icon for macOS bundle |
| `requirements.txt` | Core runtime dependency (`pyodbc`) |
| `requirements_gui.txt` | GUI/build dependencies (`customtkinter`, `tkinterdnd2`, `pyinstaller`) |
| `environment.yml` | Conda environment definition |