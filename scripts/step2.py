import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

# =====================================================
# This code works on Format Workbook
# =====================================================

# ---------------- PATHS ----------------
BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input_files"

FORMAT_FILE = INPUT_DIR / "Format.xlsx"

# ---------------- LOAD FORMAT FILE ----------------
wb = load_workbook(FORMAT_FILE)
ws = wb["Sheet1"]  # assuming your sheet is named Sheet1

# ---------------- READ INTO DATAFRAME ----------------
df = pd.DataFrame(ws.values)
df.columns = df.iloc[0]  # first row as header
df = df[1:]  # data only

# ---------------- REMOVE ROWS WHERE GROUP IS #N/A ----------------
group_col = "Group"  # Kth column
df = df[df[group_col] != "#N/A"]

# ---------------- BUSINESS RULE FOR CLOSING_ORDER ----------------
closing_col = "Closing_Order"  # Lth column
allowed = {
    "CLOSING ORDER: AWARD ON STIPULATED FINDS AND AWARD (GRANTED)",
    "CLOSING ORDER: C & R (GRANTED)",
    "CLOSING ORDER: DISMISSAL OF CLAIM"
}

df[closing_col] = df[closing_col].where(df[closing_col].isin(allowed), "CIC Pending")

# ---------------- WRITE BACK TO EXCEL ----------------
with pd.ExcelWriter(FORMAT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False)

print("✅ Group cleanup and Closing_Order applied successfully.")
