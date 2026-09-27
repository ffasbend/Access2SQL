"""
Unit tests for saved-query extraction (access2sql.get_saved_query_sql / export_queries).

Every saved query whose name starts with "TEST" in tests/data/test_db01.accdb is
compared, character for character, with the SQL text Access shows in its SQL view.
Other queries in the database are ignored. The expected texts are stored in
tests/data/test_db01_expected_queries.json (JSON, so trailing spaces cannot be
stripped by an editor).

Run from the project root:
    python -m unittest discover -s tests -v
See readme_testing.md for details.
"""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import access2sql  # noqa: E402

DATA_DIR = ROOT / "tests" / "data"
TEST_DB = DATA_DIR / "test_db01.accdb"
EXPECTED_FILE = DATA_DIR / "test_db01_expected_queries.json"
TEST_PREFIX = "TEST"   # only queries whose name starts with this are tested

EXPECTED: dict[str, str] = json.loads(EXPECTED_FILE.read_text(encoding="utf-8"))

MDBTOOLS_MISSING = not (access2sql._mdb_binary_exists("mdb-export")
                        and access2sql._mdb_queries_available())
SKIP_REASON = "mdbtools (mdb-export, mdb-queries) not found on PATH"


@unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
class TestQueryList(unittest.TestCase):
    def test_all_queries_are_covered(self) -> None:
        """Every TEST* query in the DB has an expected text, and vice versa."""
        found = {q for q in access2sql.list_saved_queries(TEST_DB)
                 if q.startswith(TEST_PREFIX)}
        self.assertEqual(found, set(EXPECTED),
                         "Query list changed — update test_db01_expected_queries.json")


@unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
class TestSavedQuerySql(unittest.TestCase):
    """One test per saved query: exact match with the Access SQL view."""
    maxDiff = None


def _make_query_test(name: str, expected: str):
    def test(self: unittest.TestCase) -> None:
        actual = access2sql.get_saved_query_sql(TEST_DB, name)
        # repr on failure makes whitespace differences visible
        self.assertEqual(actual, expected,
                         f"\nQuery {name!r}\nexpected: {expected!r}\nactual:   {actual!r}")
    test.__doc__ = f"Query '{name}' matches Access SQL view"
    return test


for _name, _sql in EXPECTED.items():
    _method = "test_" + "".join(c if c.isalnum() else "_" for c in _name)
    setattr(TestSavedQuerySql, _method, _make_query_test(_name, _sql))


@unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
class TestExportQueriesFiles(unittest.TestCase):
    """export_queries() writes every query into the .txt and .md files."""

    def setUp(self) -> None:
        # Work on a copy: export_queries() writes next to the database.
        self._tmp = tempfile.TemporaryDirectory()
        self.db = Path(self._tmp.name) / TEST_DB.name
        shutil.copy2(TEST_DB, self.db)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_txt_export(self) -> None:
        path = access2sql.export_queries(self.db, "txt")
        self.assertEqual(path.name, "test_db01_queries.txt")
        text = path.read_text(encoding="utf-8")
        for name, sql in EXPECTED.items():
            self.assertIn(f"QUERY NAME: {name}\n{'-' * 60}\n{sql}\n", text, name)

    def test_md_export(self) -> None:
        path = access2sql.export_queries(self.db, "md")
        self.assertEqual(path.name, "test_db01_queries.md")
        text = path.read_text(encoding="utf-8")
        for name, sql in EXPECTED.items():
            self.assertIn(f"## {name}\n\n```sql\n{sql}\n```\n", text, name)

    def test_existing_file_is_not_overwritten(self) -> None:
        first = access2sql.export_queries(self.db, "txt")
        second = access2sql.export_queries(self.db, "txt")
        self.assertNotEqual(first, second)
        self.assertEqual(second.name, "test_db01_queries_1.txt")


if __name__ == "__main__":
    unittest.main(verbosity=2)
