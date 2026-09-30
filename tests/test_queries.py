"""
Unit tests for saved-query extraction (access2sql.get_saved_query_sql / export_queries).

For every test database listed in TEST_DATABASES, each saved query whose name starts
with "TEST_" is compared, character for character, with the SQL text Access shows in
its SQL view. Other queries in the databases are ignored.

The expected texts are stored in tests/data/<db>_expected_queries.json (JSON, so
trailing spaces cannot be stripped by an editor).

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
TEST_PREFIX = "TEST_"   # only queries whose name starts with this are tested

# Test databases: tests/data/<name>.accdb + tests/data/<name>_expected_queries.json
TEST_DATABASES = ["test_db01", "test_db02"]

MDBTOOLS_MISSING = not (access2sql._mdb_binary_exists("mdb-export")
                        and access2sql._mdb_queries_available())
SKIP_REASON = "mdbtools (mdb-export, mdb-queries) not found on PATH"


def _load_expected(db_name: str) -> dict[str, str]:
    path = DATA_DIR / f"{db_name}_expected_queries.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _identifier(text: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in text)


def _make_query_test(db: Path, name: str, expected: str):
    def test(self: unittest.TestCase) -> None:
        actual = access2sql.get_saved_query_sql(db, name)
        # repr on failure makes whitespace differences visible
        self.assertEqual(actual, expected,
                         f"\n{db.name} / query {name!r}\n"
                         f"expected: {expected!r}\nactual:   {actual!r}")
    test.__doc__ = f"{db.name}: query '{name}' matches Access SQL view"
    return test


def _make_test_classes(db_name: str) -> list[type]:
    db = DATA_DIR / f"{db_name}.accdb"
    expected = _load_expected(db_name)
    cls_suffix = "".join(p.capitalize() for p in db_name.split("_"))

    @unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
    class QueryList(unittest.TestCase):
        def test_all_queries_are_covered(self) -> None:
            """Every TEST_ query in the DB has an expected text, and vice versa."""
            found = {q for q in access2sql.list_saved_queries(db)
                     if q.startswith(TEST_PREFIX)}
            self.assertEqual(found, set(expected),
                             f"Query list changed — update {db_name}_expected_queries.json")

    @unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
    class SavedQuerySql(unittest.TestCase):
        """One test per TEST_ query: exact match with the Access SQL view."""
        maxDiff = None

    for name, sql in expected.items():
        setattr(SavedQuerySql, f"test_{_identifier(name)}", _make_query_test(db, name, sql))

    @unittest.skipIf(MDBTOOLS_MISSING, SKIP_REASON)
    class ExportQueriesFiles(unittest.TestCase):
        """export_queries() writes every query into the .txt and .md files."""

        def setUp(self) -> None:
            # Work on a copy: export_queries() writes next to the database.
            self._tmp = tempfile.TemporaryDirectory()
            self.db = Path(self._tmp.name) / db.name
            shutil.copy2(db, self.db)

        def tearDown(self) -> None:
            self._tmp.cleanup()

        def test_txt_export(self) -> None:
            path = access2sql.export_queries(self.db, "txt")
            self.assertEqual(path.name, f"{db_name}_queries.txt")
            text = path.read_text(encoding="utf-8")
            for name, sql in expected.items():
                self.assertIn(f"QUERY NAME: {name}\n{'-' * 60}\n{sql}\n", text, name)

        def test_md_export(self) -> None:
            path = access2sql.export_queries(self.db, "md")
            self.assertEqual(path.name, f"{db_name}_queries.md")
            text = path.read_text(encoding="utf-8")
            for name, sql in expected.items():
                self.assertIn(f"## {name}\n\n```sql\n{sql}\n```\n", text, name)

        def test_existing_file_is_not_overwritten(self) -> None:
            first = access2sql.export_queries(self.db, "txt")
            second = access2sql.export_queries(self.db, "txt")
            self.assertNotEqual(first, second)
            self.assertEqual(second.name, f"{db_name}_queries_1.txt")

    classes = []
    for base, cls in (("TestQueryList", QueryList),
                      ("TestSavedQuerySql", SavedQuerySql),
                      ("TestExportQueriesFiles", ExportQueriesFiles)):
        cls.__name__ = cls.__qualname__ = f"{base}{cls_suffix}"
        classes.append(cls)
    return classes


# Register one set of test classes per database, e.g. TestSavedQuerySqlTestDb02.
for _db_name in TEST_DATABASES:
    for _cls in _make_test_classes(_db_name):
        globals()[_cls.__name__] = _cls
del _db_name, _cls   # else unittest would also collect the loop variable as a test class


if __name__ == "__main__":
    unittest.main(verbosity=2)
