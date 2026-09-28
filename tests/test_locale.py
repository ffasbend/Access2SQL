"""
Regression tests: table export must work without a UTF-8 locale.

The released app (PyInstaller bundle, started from Finder) runs mdbtools with no
UTF-8 locale. mdbtools could then not parse the non-ASCII table name "Employé",
so CREATE TABLE came out empty and no INSERTs were written (fixed in v1.5.5 by
access2sql._mdb_env()).

Run from the project root:
    python -m unittest discover -s tests -v
"""
import json
import os
import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import access2sql  # noqa: E402

EMPLOYES_DB = ROOT / "tests" / "data" / "Employes.accdb"

MDBTOOLS_MISSING = not access2sql._mdbtools_available()
SKIP_REASON = "mdbtools (mdb-tables, mdb-export, mdb-schema) not found on PATH"


class TestMdbEnv(unittest.TestCase):
    """_mdb_env() adds a UTF-8 locale only when none is set."""

    def test_adds_utf8_locale_when_missing(self) -> None:
        with mock.patch.dict(os.environ, {"PATH": os.environ.get("PATH", "")}, clear=True):
            env = access2sql._mdb_env()
        self.assertRegex(env.get("LC_ALL", ""), r"(?i)utf-?8")

    def test_adds_utf8_locale_for_plain_c_locale(self) -> None:
        with mock.patch.dict(os.environ, {"LANG": "C"}, clear=True):
            env = access2sql._mdb_env()
        self.assertRegex(env.get("LC_ALL", ""), r"(?i)utf-?8")

    def test_keeps_existing_utf8_locale(self) -> None:
        with mock.patch.dict(os.environ, {"LANG": "fr_LU.UTF-8"}, clear=True):
            env = access2sql._mdb_env()
        self.assertNotIn("LC_ALL", env)
        self.assertEqual(env["LANG"], "fr_LU.UTF-8")


# Runs in a child Python with no locale and without Python's C-locale coercion,
# which is how the Python inside the PyInstaller app behaves.
_CHILD_SCRIPT = """
import json, sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
import access2sql
schema, data = access2sql.try_mdbtools(Path(sys.argv[2]))
print(json.dumps({t: [len(schema[t]["columns"]), len(data.get(t, []))] for t in schema}))
"""


@unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
class TestExportWithoutLocale(unittest.TestCase):
    def test_non_ascii_table_name_exports_columns_and_rows(self) -> None:
        env = {
            "PATH": os.environ.get("PATH", ""),
            "HOME": os.environ.get("HOME", ""),
            "PYTHONCOERCECLOCALE": "0",
        }
        result = subprocess.run(
            [sys.executable, "-c", _CHILD_SCRIPT, str(ROOT), str(EMPLOYES_DB)],
            capture_output=True, text=True, encoding="utf-8", env=env, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        counts = json.loads(result.stdout.strip().splitlines()[-1])
        self.assertEqual(counts, {"Employé": [8, 19]},
                         "expected 8 columns and 19 rows in table 'Employé'")


if __name__ == "__main__":
    unittest.main(verbosity=2)
