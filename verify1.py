import pandas as pd
import subprocess
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys
from datetime import datetime

# ========================= CONFIG =========================
ORACLE_URL = "oracle+oracledb://your_user:your_pass@your_oracle_host:1521/your_service_name"
POLARDB_URL = "oracle+oracledb://your_user:your_pass@your_polardb_host:1521/your_polardb_service_name"

EXCEL_FILE = "tables_to_compare.xlsx"
SHEET_NAME = "Sheet1"
TABLE_COLUMN = "Table_Name"

OUTPUT_DIR = Path("comparison_results")
OUTPUT_DIR.mkdir(exist_ok=True)

# Parallelism settings
MAX_WORKERS = 4          # ← Change this: start with 2-4. Increase based on your DB connection limits
                         # (Oracle/PolarDB can usually handle 4-8 concurrent heavy queries safely)

# data-diff options per table
THREADS_PER_TABLE = 2    # Internal threads per data-diff run (for large tables)
LIMIT_DIFFERENCES = 500  # Max differing rows to show per table
# =========================================================

def run_data_diff(table_name: str):
    """Run data-diff for one table."""
    start_time = datetime.now()
    output_file = OUTPUT_DIR / f"{table_name}_diff.txt"
    json_file   = OUTPUT_DIR / f"{table_name}_diff.jsonl"

    cmd = [
        "data-diff",
        ORACLE_URL, table_name,
        POLARDB_URL, table_name,
        "--json",
        "--threads", str(THREADS_PER_TABLE),   # Internal parallelism per table
        "--limit", str(LIMIT_DIFFERENCES) if LIMIT_DIFFERENCES else "0"
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=7200)  # 2-hour timeout

        # Save outputs
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(f"=== DATA-DIFF RESULT FOR TABLE: {table_name} ===\n")
            f.write(f"Started: {start_time}\n")
            f.write(f"Duration: {datetime.now() - start_time}\n\n")
            f.write(result.stdout)
            if result.stderr:
                f.write("\n--- STDERR ---\n" + result.stderr)

        if result.stdout.strip():
            with open(json_file, "w", encoding="utf-8") as f:
                f.write(result.stdout)

        duration = datetime.now() - start_time
        if result.returncode == 0:
            status = "✅ IDENTICAL"
            print(f"✅ {table_name} → IDENTICAL ({duration})")
        else:
            status = "❌ DIFFERENCES FOUND"
            print(f"❌ {table_name} → DIFFERENCES FOUND ({duration}) - see {output_file}")
        return f"{table_name}: {status}"

    except Exception as e:
        print(f"❌ Error on {table_name}: {e}")
        return f"{table_name}: ERROR - {e}"


# ====================== MAIN ======================
print("Reading table list from Excel...")
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, usecols=[TABLE_COLUMN], skiprows=0, header=0)
tables = df[TABLE_COLUMN].dropna().astype(str).str.strip().tolist()

print(f"Found {len(tables)} tables. Running up to {MAX_WORKERS} in parallel...\n")

results_summary = []
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    future_to_table = {executor.submit(run_data_diff, table): table for table in tables}
    
    for future in as_completed(future_to_table):
        result = future.result()
        results_summary.append(result)

# Save final summary
summary_file = OUTPUT_DIR / "SUMMARY.txt"
with open(summary_file, "w", encoding="utf-8") as f:
    f.write(f"=== PARALLEL COMPARISON SUMMARY (Oracle vs PolarDB) ===\n")
    f.write(f"Completed at: {datetime.now()}\n")
    f.write(f"Max concurrent tables: {MAX_WORKERS}\n\n")
    for line in results_summary:
        f.write(line + "\n")

print("\n" + "="*70)
print("✅ ALL TABLE COMPARISONS COMPLETED!")
print(f"Results saved to: {OUTPUT_DIR.absolute()}")
print(f"Summary: {summary_file}")
print("="*70)

for line in sorted(results_summary):
    print("  •", line)