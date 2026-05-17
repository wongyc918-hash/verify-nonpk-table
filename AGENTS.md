# Project: verify-nonpk-table

Oracle → PolarDB (PostgreSQL-compatible) data migration verification tool.
Compares table data between the two databases **without relying on primary keys** — designed for tables that lack PKs.

## Scripts Overview

| File | Approach | Best For |
|------|----------|----------|
| `verify1.py` | `data-diff` CLI (subprocess) | Quick structural diff using the external `data-diff` tool |
| `verify2.py` | Full row fetch + positional diff | Small/medium tables; produces color-coded Excel side-by-side diff |
| `verify3.py` | Chunked MD5 hash streaming | Large tables; low memory footprint (O(CHUNK_SIZE) not O(table)) |
| `verify4.py` | Same as verify3 + server-side cursor | Very large PolarDB tables; adds PostgreSQL keepalive and `stream_results=True` |
| `ora_query.py` | Oracle health check | Sessions, deadlocks, tablespace usage via `v$session` / DBA views |
| `testconn.py` | Standalone calculator | Unrelated utility; ignore for DB work |

## Configuration

Each verify script has a `# --- CONFIGURATION ---` block at the top. Update these before running:
- `ORACLE_CONFIG` — SQLAlchemy URL: `oracle+oracledb://user:pass@host:port/?service_name=svc`
- `POLARDB_CONFIG` — SQLAlchemy URL: `postgresql://user:pass@host:port/dbname`
- `INPUT_EXCEL` — Excel file listing table names (column header: `Table_Name`)
- `OUTPUT_EXCEL` — Output report file
- `MAX_WORKERS` — Parallel threads (start with 2–4)
- `CHUNK_SIZE` (verify3/4) — Rows per streaming chunk (100K for verify3, 1M for verify4)

## Key Patterns

- **Column normalization**: Oracle columns are UPPER-cased; scripts lower-case and sort alphabetically before comparison.
- **Value normalization** (`normalize_col`): datetimes → `%Y-%m-%d %H:%M:%S` string, numerics → `round(float, 6)` string, strings → stripped with `None`/`nan`/`NaT`/`<NA>` replaced by `''`.
- **Excel input format**: Single sheet (`Sheet1`), one column named `Table_Name`. See `tables_to_verify1.xlsx` for an example.
- **Output**: Excel workbook with a summary sheet + one detail sheet per differing table. Cells color-coded: blue header, yellow = Oracle value, light-blue = PolarDB value, orange = mismatch.

## Dependencies

```
pandas
sqlalchemy
oracledb          # cx_Oracle replacement
openpyxl
data-diff         # only needed for verify1.py
```

Activate the local venv before running:
```powershell
& .\venv\Scripts\Activate.ps1
python verify4.py
```

## Non-PK Design Notes

- No `ORDER BY` is used — comparison is unordered (hash counter or positional).
- `verify2.py` positional diff assumes both DBs return rows in the **same order** (no guarantee on large tables — prefer verify3/4 for those).
- `verify3.py` / `verify4.py` use `Counter(hash → count)` so duplicate rows are handled correctly even without a PK.
