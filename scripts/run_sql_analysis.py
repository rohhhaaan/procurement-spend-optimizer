import os
import sqlite3
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SQL_FILE = os.path.join(BASE_DIR, "sql", "queries.sql")
DB_PATH = os.path.join(DATA_DIR, "procurement_analytics.db")
EXPORT_DIR = os.path.join(DATA_DIR, "sql_exports")
os.makedirs(EXPORT_DIR, exist_ok=True)


def load_csvs_to_sqlite(conn: sqlite3.Connection):
    tables = {
        "suppliers": "suppliers.csv",
        "sku_master": "sku_master.csv",
        "purchase_orders": "purchase_orders.csv",
        "rfq_bom_quotes": "rfq_bom_quotes.csv",
    }
    for table_name, csv_file in tables.items():
        df = pd.read_csv(os.path.join(DATA_DIR, csv_file))
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        print(f"Loaded table '{table_name}' ({len(df):,} rows) into SQLite.")

    # Create analytical indexes for fast joins and aggregations
    cur = conn.cursor()
    cur.execute("CREATE INDEX IF NOT EXISTS idx_po_supplier ON purchase_orders(supplier_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_po_sku ON purchase_orders(sku_id);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_po_month ON purchase_orders(po_month);")
    conn.commit()


def execute_analytical_queries(conn: sqlite3.Connection):
    with open(SQL_FILE, "r", encoding="utf-8") as f:
        sql_script = f.read()

    # Split statements cleanly by semicolon
    raw_statements = [stmt.strip() for stmt in sql_script.split(";") if stmt.strip()]

    query_names = [
        "01_supplier_scorecard",
        "02_category_spend_freight",
        "03_maverick_consolidation_savings",
        "04_inventory_turnover_risk",
        "05_monthly_spend_otd_trend",
    ]

    select_idx = 0
    for stmt in raw_statements:
        # Strip leading comment lines to inspect statement type
        lines = [ln for ln in stmt.splitlines() if not ln.strip().startswith("--") and ln.strip()]
        if not lines:
            continue
        clean_stmt = "\n".join(lines)

        if clean_stmt.upper().startswith(("DROP VIEW", "CREATE VIEW")):
            conn.execute(clean_stmt)
            conn.commit()
        elif clean_stmt.upper().startswith(("WITH", "SELECT")):
            q_name = query_names[select_idx] if select_idx < len(query_names) else f"query_{select_idx+1}"
            select_idx += 1

            df_res = pd.read_sql_query(clean_stmt, conn)
            out_csv = os.path.join(EXPORT_DIR, f"{q_name}.csv")
            df_res.to_csv(out_csv, index=False)

            print(f"\n{'='*75}")
            print(f"RESULT: {q_name.upper()} ({len(df_res)} rows) -> Saved to {q_name}.csv")
            print(f"{'='*75}")
            print(df_res.head(10).to_string(index=False))


def main():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    try:
        load_csvs_to_sqlite(conn)
        execute_analytical_queries(conn)
    finally:
        conn.close()
    print(f"\nSQLite database saved successfully at: {DB_PATH}")


if __name__ == "__main__":
    main()