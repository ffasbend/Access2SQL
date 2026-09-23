#!/usr/bin/env python3
"""
build_mac.py  –  Build Access2SQL.app for macOS

What it does:
  1. Copies the mdbtools binaries from Homebrew into a staging area.
  2. Collects every non-system dylib they depend on (recursive otool walk).
  3. Rewrites all dylib references with install_name_tool so they resolve
     relative to the bundle (no Homebrew required on the target machine).
  4. Runs PyInstaller to produce dist/Access2SQL.app with everything bundled.

Requirements (run once):
  brew install mdbtools
  pip install pyinstaller tkinterdnd2
  xcode-select --install          # for otool / install_name_tool
"""
import subprocess
import sys
import shutil
from pathlib import Path

APP_NAME  = "Access2SQL"
STAGING   = Path("build") / "mdb_staging"
MDB_TOOLS = [
    "mdb-tables", "mdb-export", "mdb-schema",
    "mdb-queries", "mdb-sql",   "mdb-prop",
]


# ── helpers ───────────────────────────────────────────────────────────────────

def die(msg: str) -> None:
    print(f"\nERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def check_prereqs() -> None:
    hints = {
        "pyinstaller":       "pip install pyinstaller",
        "mdb-tables":        "brew install mdbtools",
        "otool":             "xcode-select --install",
        "install_name_tool": "xcode-select --install",
    }
    missing = [cmd for cmd in hints if not shutil.which(cmd)]
    if missing:
        for cmd in missing:
            print(f"  missing: {cmd}  ->  {hints[cmd]}")
        die("Install the missing tools above and retry.")


def _deps(binary: Path) -> list[Path]:
    """Non-system dylib dependencies of *binary* (via otool -L)."""
    out = subprocess.run(
        ["otool", "-L", str(binary)], capture_output=True, text=True
    ).stdout
    result: list[Path] = []
    for line in out.splitlines()[1:]:
        path = line.strip().split(" (")[0].strip()
        if not path:
            continue
        if path.startswith(("/usr/lib", "/System", "@")):
            continue
        result.append(Path(path))
    return result


def _fix_refs(binary: Path, lib_dir: Path, loader_prefix: str) -> None:
    """Rewrite all non-system dylib references in *binary* to use @loader_path."""
    for dep in _deps(binary):
        subprocess.run(
            ["install_name_tool", "-change", str(dep),
             f"{loader_prefix}/{dep.name}", str(binary)],
            capture_output=True,
        )


# ── staging ───────────────────────────────────────────────────────────────────

def stage_mdbtools() -> None:
    bin_dir = STAGING / "bin"
    lib_dir = STAGING / "lib"
    shutil.rmtree(STAGING, ignore_errors=True)
    bin_dir.mkdir(parents=True)
    lib_dir.mkdir(parents=True)

    # Copy mdbtools executables
    for name in MDB_TOOLS:
        src = shutil.which(name)
        if not src:
            print(f"  skip    {name}  (not on PATH)")
            continue
        dest = bin_dir / name
        shutil.copy2(src, dest)
        dest.chmod(0o755)
        print(f"  binary  {name}")

    # Collect dylibs with a BFS walk
    processed: set[Path] = set()
    queue: list[Path] = list(bin_dir.iterdir())
    while queue:
        item = queue.pop()
        if item in processed:
            continue
        processed.add(item)
        for dep in _deps(item):
            if not dep.exists():
                continue
            dest = lib_dir / dep.name
            if not dest.exists():
                shutil.copy2(dep, dest)
                dest.chmod(0o755)
                print(f"  dylib   {dep.name}")
                queue.append(dest)

    # Fix rpath references so the bundle is self-contained
    for binary in bin_dir.iterdir():
        _fix_refs(binary, lib_dir, "@loader_path/../lib")
    for lib in lib_dir.iterdir():
        subprocess.run(
            ["install_name_tool", "-id", f"@loader_path/{lib.name}", str(lib)],
            capture_output=True,
        )
        _fix_refs(lib, lib_dir, "@loader_path")

    nb = sum(1 for _ in bin_dir.iterdir())
    nl = sum(1 for _ in lib_dir.iterdir())
    print(f"  staged  {nb} binaries,  {nl} dylibs")


# ── PyInstaller ───────────────────────────────────────────────────────────────

def build() -> None:
    print(f"=== Building {APP_NAME}.app ===\n")

    print("Checking prerequisites…")
    check_prereqs()

    print("\nStaging mdbtools…")
    stage_mdbtools()

    # Build --add-binary arguments
    add_binary: list[str] = []
    for f in sorted((STAGING / "bin").iterdir()):
        add_binary += ["--add-binary", f"{f}:mdbtools/bin"]
    for f in sorted((STAGING / "lib").iterdir()):
        if f.suffix in (".dylib", ".so"):
            add_binary += ["--add-binary", f"{f}:mdbtools/lib"]

    print("\nRunning PyInstaller…")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--name", APP_NAME,
        "--windowed",           # no terminal window on macOS
        "--collect-all", "tkinterdnd2",
        "--collect-all", "customtkinter",
        *add_binary,
        "access2sql_gui.py",
    ]
    subprocess.run(cmd, check=True)

    print(f"\nDone!  Open with:  open dist/{APP_NAME}.app")


if __name__ == "__main__":
    build()
