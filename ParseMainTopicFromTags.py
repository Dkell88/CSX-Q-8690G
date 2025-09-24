#!/usr/bin/env python3
import sys
import csv
from pathlib import Path

import pandas as pd

# ── PATH SETUP ────────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).resolve().parent
BASE_DIR  = SCRIPT_DIR.parent     # …\10 Conceptual

INPUT_FILE  = BASE_DIR / "WonderWare Tag Export.CSV"
OUTPUT_FILE = BASE_DIR / "WonderWare Tag Export - Main Filtered.xlsx"
# ────────────────────────────────────────────────────────────────────────────────
#----#

def main():
    if not INPUT_FILE.is_file():
        sys.exit(f"❌ Input file not found:\n   {INPUT_FILE}")

    matches = []
    with INPUT_FILE.open(newline='', encoding='utf-8', errors='replace') as f:
        reader = csv.reader(f)
        for row in reader:
            # skip empty or section‐header lines
            if not row or row[0].startswith(':'):
                continue
            # need at least two columns
            if len(row) < 2:
                continue
            # exact match on column B
            if row[1] == "Main":
                matches.append(row)

    if not matches:
        sys.exit("❌ No rows with column B == \"Main\" found.")

    # pad all rows to the same length
    max_cols = max(len(r) for r in matches)
    padded = [r + [""]*(max_cols - len(r)) for r in matches]

    # build a DataFrame
    cols = [f"Col{i+1}" for i in range(max_cols)]
    df = pd.DataFrame(padded, columns=cols)

    # write out
    try:
        df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    except Exception as e:
        sys.exit(f"❌ Failed to write {OUTPUT_FILE!r}:\n   {e}")

    print(f"✅ Exported {len(matches)} rows to:\n   {OUTPUT_FILE}")

if __name__ == "__main__":
    main()