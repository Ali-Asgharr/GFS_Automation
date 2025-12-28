import pandas as pd
from pathlib import Path

# =====================================================
# This code works on Delta-Workbook
# =====================================================

# ================= PATH =================
BASE_DIR = Path(__file__).resolve().parent.parent
DELTA_FILE = BASE_DIR / "input_files" / "Delta.xlsx"

# ================= COLUMN INDEX MAP =================
# Excel columns start at 0 index:
# A=0, B=1, C=2, ...

WC_COLS = [0, 2, 3, 4, 6, 7, 8, 9, 10, 11, 32, 33, 34, 38]
MED_COLS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
IMR_COLS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]

FINAL_COLUMNS = [
    "Date",
    "Location",
    "Case Id",
    "MRN",
    "Billed",
    "Paid",
    "Outstanding",
    "Settlement Amount",
    "Settlement %Age",
    "ProviderID",
    "Settlement Payment Received?",
    "Settlement Payment",
    "Remaining Payment",
    "Status"
]

# ================= READ SHEETS =================
print("📥 Reading WC Delta...")
wc_df = pd.read_excel(DELTA_FILE, sheet_name="WC Delta", usecols=WC_COLS)
wc_df.columns = FINAL_COLUMNS

print("📥 Reading Med-Legal Delta...")
med_df = pd.read_excel(DELTA_FILE, sheet_name="Med-Legal Delta", usecols=MED_COLS)
med_df.columns = FINAL_COLUMNS

print("📥 Reading IMR Delta...")
imr_df = pd.read_excel(DELTA_FILE, sheet_name="IMR Delta", usecols=IMR_COLS)
imr_df.columns = FINAL_COLUMNS

# ================= COMBINE =================
combined_df = pd.concat([wc_df, med_df, imr_df], ignore_index=True)

# ================= WRITE BACK =================
with pd.ExcelWriter(
    DELTA_FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    combined_df.to_excel(writer, sheet_name="Combined", index=False)
    combined_df.to_excel(writer, sheet_name="2025", index=False)
    combined_df.to_excel(writer, sheet_name="GFS", index=False)

print("✅ SUCCESS: Combined, 2025, and GFS sheets created correctly.")
