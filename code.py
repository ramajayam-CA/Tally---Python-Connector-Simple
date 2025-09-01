import pyodbc
import pandas as pd

# =========================
# SETTINGS
# =========================
TALLY_DSN = "TallyODBC64_9000"  # ODBC DSN Name for Tally (adjust as per your ODBC setup)
OUTPUT_FILE = "Trial_Balance.xlsx"

# =========================
# CONNECT TO TALLY
# =========================
try:
    conn = pyodbc.connect(f"DSN={TALLY_DSN}")
    print("✅ Connected to Tally ODBC")
except Exception as e:
    print("❌ Connection failed:", e)
    exit()
#
# =========================
# QUERY TRIAL BALANCE
# =========================
query = """Select $name, $_PrimaryGroup, $Parent, $OpeningBalance, $_ClosingBalance from Ledger"""
# The above is a simple Ledger query — Tally auto-calculates TB when date range is set.

# =========================
# EXECUTE QUERY
# =========================
try:
    df = pd.read_sql(query, conn)
    print(f"✅ Retrieved {len(df)} ledger records")
except Exception as e:
    print("❌ Query failed:", e)
    conn.close()
    exit()

conn.close()

# =========================
# EXPORT TO EXCEL
# =========================
try:
    df.to_csv("Tb.csv")
except Exception as e:
    print("❌ Excel export failed:", e)
