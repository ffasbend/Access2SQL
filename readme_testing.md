# Testing Access2SQL

The unit tests check that:

- saved Access queries are exported **exactly** as Access shows them in its SQL view:
  same keywords (`TOP`, `AS`, `DISTINCT`, …), same line breaks, same spaces;
- tables are exported correctly even **without a UTF-8 locale**, as when the app is
  started from Finder (regression test for tables with non-ASCII names like `Employé`).

## What is tested

| File | Purpose |
|---|---|
| [tests/test_queries.py](tests/test_queries.py) | Saved-query tests (standard-library `unittest`) |
| [tests/test_locale.py](tests/test_locale.py) | Locale / non-ASCII table name tests |
| [tests/data/test_db01.accdb](tests/data/test_db01.accdb) | Test Access database with tables and saved queries |
| [tests/data/test_db01_expected_queries.json](tests/data/test_db01_expected_queries.json) | Expected SQL text for every tested query |
| [tests/data/Employes.accdb](tests/data/Employes.accdb) | Table `Employé` (8 columns, 19 rows) for the locale test |

Only queries whose name starts with **`TEST`** are tested (`TEST_as`, `TEST_top`, …).
The other queries in the database (`Query 1`, `Query 2`, …) are ignored.

The tests are:

| Test class | What it checks |
|---|---|
| `TestQueryList` | The set of `TEST*` queries in the database equals the set in the JSON file, so a new query can't be forgotten |
| `TestSavedQuerySql` | One test per `TEST*` query: `get_saved_query_sql()` returns exactly the expected text |
| `TestExportQueriesFiles` | `export_queries()` writes every query into the `.txt` and `.md` files, and never overwrites an existing file |
| `TestMdbEnv` | `_mdb_env()` adds a UTF-8 locale for mdbtools when none (or plain `C`) is set, and keeps an existing one |
| `TestExportWithoutLocale` | Runs `try_mdbtools()` in a child Python with **no locale** and `PYTHONCOERCECLOCALE=0` (how the Python inside the app behaves); table `Employé` must have 8 columns and 19 rows |

The export tests work on a temporary copy of the database, so running the tests never
creates files in `tests/data/`.

The expected texts are stored as **JSON** on purpose. Several queries have trailing spaces
(e.g. `TEST_like_and`: `WHERE fldNom LIKE '?e*' ⏎`), which an editor could silently strip
from a `.sql` or `.py` file.

## Requirements

- **Python 3.10+.** No extra Python packages are needed.
- **mdbtools**, with `mdb-export` and `mdb-queries` on the `PATH`:

  | Platform | Install |
  |---|---|
  | macOS | `brew install mdbtools` |
  | Linux (Debian/Ubuntu) | `sudo apt install mdbtools` |
  | Conda (any OS) | `conda env create -f environment.yml` (includes mdbtools) |
  | Windows | Not supported: mdbtools isn't available, so all tests are **skipped** |

  Check the install:

  ```bash
  mdb-export --version
  which mdb-queries
  ```

If mdbtools is missing, the tests are reported as **skipped** (not failed), with the
reason *"mdbtools (mdb-export, mdb-queries) not found on PATH"*.

## Running the tests

Always run the tests from the **project root** (the folder that contains `access2sql.py`).

Run all tests:

```bash
python -m unittest discover -s tests -v
```

Run a single test class or a single query:

```bash
python -m unittest tests.test_queries.TestSavedQuerySql -v
python -m unittest tests.test_queries.TestSavedQuerySql.test_TEST_top_all -v
```

Or run the test file directly:

```bash
python tests/test_queries.py
```

**With pytest** (optional, `pip install pytest`). pytest runs the same `unittest` tests:

```bash
pytest tests -v
pytest tests -k TEST_like -v      # only the LIKE queries
```

### Expected output

```
test_TEST_as (test_queries.TestSavedQuerySql.test_TEST_as)
Query 'TEST_as' matches Access SQL view ... ok
...
----------------------------------------------------------------------
Ran 23 tests in 0.22s

OK
```

### When a test fails

Failure messages show both texts with `repr()`, so whitespace differences are visible:

```
Query 'TEST_like_and'
expected: "SELECT *\nFROM tblClients\nWHERE fldNom LIKE '?e*' \nAND fldLocalité = 'Luxembourg';"
actual:   "SELECT *\nFROM tblClients\nWHERE fldNom LIKE '?e*'\n  AND fldLocalité = 'Luxembourg';"
```

## Adding a new test query

1. Open `tests/data/test_db01.accdb` in Microsoft Access.
2. Create a query whose name starts with **`TEST`**, e.g. `TEST_left_join`, and save it.
3. Open the query in **SQL view** and copy the text exactly.
4. Add it to `tests/data/test_db01_expected_queries.json`. Write line breaks as `\n`,
   keep every space, and end with `;` as Access does:

   ```json
   "TEST_left_join": "SELECT *\nFROM tblClients LEFT JOIN tblFactures ON tblClients.idClient = tblFactures.fiClient;"
   ```

5. Run the tests. `TestQueryList` fails until the database and the JSON file contain
   the same `TEST*` queries.

> Take the expected text from **Access's SQL view**, not from the program's output.
> Otherwise the test only confirms what the code already does.

Good candidates that aren't covered yet: joins (`INNER` / `LEFT` / `RIGHT`), queries with
`PARAMETERS`, `DISTINCTROW`, `TOP n PERCENT`, make-table (`SELECT … INTO`),
`INSERT INTO … SELECT` (append from a table), `INSERT` with a column list, and
`UPDATE` / `DELETE` with a join.
