import pandas as pd
from sqlalchemy import create_engine, text
import concurrent.futures
import hashlib
from collections import Counter
from openpyxl.styles import PatternFill, Font

# --- CONFIGURATION ---
ORACLE_CONFIG  = "oracle+oracledb://matt:oracle123@localhost:22/?service_name=bcauat2"
POLARDB_CONFIG = "postgresql://matt2:Abcd1234@localhost:23/bcauat2-db-team"
INPUT_EXCEL    = "tables_to_verify3.xlsx"
OUTPUT_EXCEL   = "verification_results.xlsx"
MAX_WORKERS    = 4

CHUNK_SIZE      = 1_000_000   # rows fetched per chunk — controls memory usage
MAX_DETAIL_ROWS = 1_000     # max differing rows written to detail sheet per DB


# ──────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────

def get_engine(connection_string):
    if connection_string.startswith("postgresql"):
        return create_engine(
            connection_string,
            connect_args={
                "keepalives":          1,
                "keepalives_idle":     60,   # send keepalive after 60s idle
                "keepalives_interval": 10,   # retry every 10s
                "keepalives_count":    5,    # drop after 5 missed replies
                "options":             "-c statement_timeout=0 -c idle_in_transaction_session_timeout=0",
            }
        )
    return create_engine(connection_string)


def normalize_col(series):
    """Convert a column to a consistent string representation across DBs."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
    elif pd.api.types.is_numeric_dtype(series):
        return series.astype(float).round(6).astype(str).where(series.notna(), '')
    else:
        return series.astype(str).str.strip().replace({'None': '', 'nan': '', 'NaT': '', '<NA>': ''})


def get_row_count(engine, table_name):
    with engine.connect() as conn:
        result = conn.execute(text(f"SELECT COUNT(*) FROM {table_name}"))
        return result.scalar()


def get_columns(engine, table_name):
    """Return sorted lowercase column names (schema-qualified table_name is fine)."""
    df = pd.read_sql(f"SELECT * FROM {table_name} WHERE 1=0", engine)
    return sorted(c.lower() for c in df.columns)


def hash_row(values):
    """MD5 hash of \x01-separated normalized string values for one row."""
    row_str = '\x01'.join(str(v) for v in values)
    return hashlib.md5(row_str.encode('utf-8')).hexdigest()


def _hash_chunk(chunk, columns):
    """Vectorized row hashing for a normalized (all-string) chunk DataFrame.
    Uses pandas str.cat for the join — faster than a Python-level loop.
    """
    if len(columns) == 1:
        joined = chunk[columns[0]].fillna('')
    else:
        joined = chunk[columns[0]].str.cat([chunk[c] for c in columns[1:]], sep='\x01', na_rep='')
    return [hashlib.md5(s.encode('utf-8')).hexdigest() for s in joined]


# ──────────────────────────────────────────────
# Core: chunked streaming hash counter
# ──────────────────────────────────────────────

def compute_hash_counter(engine, table_name, columns, total_rows=None):
    """
    Stream the full table in chunks of CHUNK_SIZE rows using a server-side cursor
    (stream_results=True) so PostgreSQL yields rows immediately instead of
    buffering the entire result set first.
    Returns Counter(row_hash -> occurrence_count).
    Peak memory: O(CHUNK_SIZE) rows, not O(total_rows).
    """
    query     = f"SELECT {', '.join(columns)} FROM {table_name}"
    counter   = Counter()
    rows_done = 0
    total_str = f"{total_rows:,}" if total_rows else "?"

    with engine.connect().execution_options(stream_results=True) as conn:
        for chunk in pd.read_sql(query, conn, chunksize=CHUNK_SIZE):
            chunk.columns = [c.lower() for c in chunk.columns]
            chunk = chunk.reindex(columns, axis=1)
            for col in columns:
                chunk[col] = normalize_col(chunk[col])
            counter.update(_hash_chunk(chunk, columns))
            rows_done += len(chunk)
            pct = f"{rows_done / total_rows * 100:.1f}%" if total_rows else ""
            print(f"      hashing {rows_done:,} / {total_str} rows {pct}", end='\r', flush=True)
    print()  # newline after progress
    return counter


def fetch_rows_matching_hashes(engine, table_name, columns, target_hashes, limit, total_rows=None):
    """
    Re-scan the table in chunks using a server-side cursor and collect rows whose
    normalized hash is in target_hashes.  Stops early once `limit` rows are collected.
    """
    query         = f"SELECT {', '.join(columns)} FROM {table_name}"
    collected     = []
    matched_total = 0
    rows_done     = 0
    total_str     = f"{total_rows:,}" if total_rows else "?"

    with engine.connect().execution_options(stream_results=True) as conn:
        for chunk in pd.read_sql(query, conn, chunksize=CHUNK_SIZE):
            chunk.columns = [c.lower() for c in chunk.columns]
            chunk = chunk.reindex(columns, axis=1)
            for col in columns:
                chunk[col] = normalize_col(chunk[col])

            chunk_hashes  = pd.Series(_hash_chunk(chunk, columns), index=chunk.index)
            matched       = chunk[chunk_hashes.isin(target_hashes)]
            collected.append(matched)
            matched_total += len(matched)
            rows_done     += len(chunk)
            pct = f"{rows_done / total_rows * 100:.1f}%" if total_rows else ""
            print(f"      scanning {rows_done:,} / {total_str} rows {pct} | found {matched_total:,} mismatches", end='\r', flush=True)

            if matched_total >= limit:
                break
    print()  # newline after progress
    return pd.concat(collected).head(limit).reset_index(drop=True) if collected else pd.DataFrame(columns=columns)


# ──────────────────────────────────────────────
# Per-table verification
# ──────────────────────────────────────────────

def verify_table(table_name):
    print(f"[#] Starting: {table_name}")
    try:
        ora_engine = get_engine(ORACLE_CONFIG)
        plr_engine = get_engine(POLARDB_CONFIG)

        # 1. Row counts (fast)
        count_ora = get_row_count(ora_engine, table_name)
        count_plr = get_row_count(plr_engine, table_name)
        print(f"    {table_name}: Oracle={count_ora:,} rows | PolarDB={count_plr:,} rows")

        # 2. Resolve common normalized columns
        cols_ora = get_columns(ora_engine, table_name)
        cols_plr = get_columns(plr_engine, table_name)
        columns  = sorted(set(cols_ora) & set(cols_plr))

        if not columns:
            return {"table": table_name, "status": "ERROR",
                    "error": "No common columns found between Oracle and PolarDB",
                    "data": None, "row_ora": count_ora, "row_plr": count_plr}

        # 3. Build hash counters by streaming both tables
        print(f"    [{table_name}] Hashing Oracle rows in chunks of {CHUNK_SIZE:,}...")
        counter_ora = compute_hash_counter(ora_engine, table_name, columns, count_ora)

        print(f"    [{table_name}] Hashing PolarDB rows in chunks of {CHUNK_SIZE:,}...")
        counter_plr = compute_hash_counter(plr_engine, table_name, columns, count_plr)

        # 4. Compare hash distributions — identical means tables match exactly
        if counter_ora == counter_plr:
            print(f"    [{table_name}] ✅ MATCH")
            return {"table": table_name, "status": "MATCH", "diff_count": 0,
                    "data": None, "row_ora": count_ora, "row_plr": count_plr, "notes": ""}

        # 5. Find which row hashes differ (count doesn't match between DBs)
        all_hashes      = set(counter_ora) | set(counter_plr)
        mismatch_hashes = frozenset(
            h for h in all_hashes if counter_ora.get(h, 0) != counter_plr.get(h, 0)
        )
        diff_row_count = sum(
            abs(counter_ora.get(h, 0) - counter_plr.get(h, 0)) for h in mismatch_hashes
        )
        print(f"    [{table_name}] ❌ {diff_row_count:,} differing rows detected. Fetching details...")

        # 6. Re-scan both tables to collect the actual mismatched rows
        rows_ora = fetch_rows_matching_hashes(ora_engine, table_name, columns, mismatch_hashes, MAX_DETAIL_ROWS, count_ora)
        rows_plr = fetch_rows_matching_hashes(plr_engine, table_name, columns, mismatch_hashes, MAX_DETAIL_ROWS, count_plr)

        rows_ora = rows_ora.copy(); rows_ora.insert(0, 'Source DB', 'Oracle')
        rows_plr = rows_plr.copy(); rows_plr.insert(0, 'Source DB', 'PolarDB')
        diff_df  = pd.concat([rows_ora, rows_plr]).reset_index(drop=True)

        truncated = len(rows_ora) >= MAX_DETAIL_ROWS or len(rows_plr) >= MAX_DETAIL_ROWS
        notes     = f"Truncated to {MAX_DETAIL_ROWS} rows per DB in detail sheet" if truncated else ""

        return {"table": table_name, "status": "MISMATCH", "diff_count": diff_row_count,
                "data": diff_df, "row_ora": count_ora, "row_plr": count_plr, "notes": notes}

    except Exception as e:
        return {"table": table_name, "status": "ERROR", "error": str(e),
                "data": None, "row_ora": None, "row_plr": None, "notes": ""}


# ──────────────────────────────────────────────
# Excel formatting
# ──────────────────────────────────────────────

def apply_diff_sheet_formatting(ws):
    """Color rows by source DB and style the header."""
    HEADER_FILL = PatternFill("solid", fgColor="4472C4")
    ORACLE_FILL = PatternFill("solid", fgColor="FFF2CC")   # yellow
    POLAR_FILL  = PatternFill("solid", fgColor="DDEBF7")   # light blue
    header_font = Font(bold=True, color="FFFFFF")
    bold_font   = Font(bold=True)

    headers        = [cell.value for cell in ws[1]]
    source_col_idx = next((i + 1 for i, h in enumerate(headers) if h == 'Source DB'), None)

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = header_font

    for row in ws.iter_rows(min_row=2):
        if source_col_idx:
            src_cell = row[source_col_idx - 1]
            fill     = ORACLE_FILL if src_cell.value == 'Oracle' else POLAR_FILL
            for cell in row:
                cell.fill = fill
            src_cell.font = bold_font


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main():
    input_df   = pd.read_excel(INPUT_EXCEL)
    table_list = input_df['table_name'].tolist()

    print(f"[!] {len(table_list)} tables to verify. "
          f"Chunk size: {CHUNK_SIZE:,} rows | Workers: {MAX_WORKERS}\n")

    results_summary  = []
    mismatch_details = {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_table = {executor.submit(verify_table, t): t for t in table_list}

        for future in concurrent.futures.as_completed(future_to_table):
            res = future.result()
            results_summary.append({
                "Table Name":     res['table'],
                "Status":         res['status'],
                "Diff Count":     res.get('diff_count', 0),
                "Rows (Oracle)":  res.get('row_ora', ''),
                "Rows (PolarDB)": res.get('row_plr', ''),
                "Notes":          res.get('notes') or res.get('error', '')
            })
            if res['status'] == "MISMATCH" and res['data'] is not None:
                mismatch_details[res['table']] = res['data']

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        summary_df = pd.DataFrame(results_summary)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Format summary sheet
        ws_sum        = writer.sheets['Summary']
        header_fill   = PatternFill("solid", fgColor="4472C4")
        header_font   = Font(bold=True, color="FFFFFF")
        match_fill    = PatternFill("solid", fgColor="C6EFCE")   # green
        mismatch_fill = PatternFill("solid", fgColor="FCE4D6")   # orange
        error_fill    = PatternFill("solid", fgColor="FFCCCC")   # red

        for cell in ws_sum[1]:
            cell.fill = header_fill
            cell.font = header_font

        status_col = summary_df.columns.get_loc("Status") + 1
        for row in ws_sum.iter_rows(min_row=2):
            st = row[status_col - 1]
            if st.value == "MATCH":       st.fill = match_fill
            elif st.value == "MISMATCH":  st.fill = mismatch_fill
            elif st.value == "ERROR":     st.fill = error_fill

        # Detail sheets
        for table_name, diff_df in mismatch_details.items():
            sheet_name = f"Diff_{table_name}"[:31]
            diff_df.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_diff_sheet_formatting(writer.sheets[sheet_name])

    print(f"\n[!] Done. Results saved to {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
