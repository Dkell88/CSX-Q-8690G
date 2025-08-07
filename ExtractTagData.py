#!/usr/bin/env python3

import re
from pathlib import Path
import pandas as pd
from lxml import etree

# ————— Configure your paths here —————
SCRIPT_DIR  = Path(__file__).resolve().parent
BASE_DIR    = SCRIPT_DIR.parent     # …\10 Conceptual

INPUT_EXCEL = BASE_DIR / "WonderWare Tag Export - Main Filtered.xlsx"
INPUT_L5X   = BASE_DIR / "Program Exports" / "P_200_Main_240418_A.L5X"
OUTPUT_FILE = BASE_DIR / "Tag Mapping.xlsx"

def extract_mappings(excel_path: Path, l5x_path: Path, output_path: Path):
    # — Sanity checks — 
    for p,label in [(excel_path,"Excel"), (l5x_path,"L5X")]:
        print(f"→ Checking {label} path: {p}")
        if not p.exists() or not p.is_file():
            raise FileNotFoundError(f"{label} file not found or not a file: {p}")

    # 1) Load your Excel and pull the tags in Col45, preserving order
    df_tags = pd.read_excel(excel_path, engine="openpyxl")
    if "Col45" not in df_tags.columns:
        raise KeyError(f"'Col45' column not found. Available: {df_tags.columns.tolist()}")
    tags = df_tags["Col45"].dropna().astype(str).tolist()

    # 2) Parse the L5X once
    parser = etree.XMLParser(remove_comments=False, recover=True)
    tree   = etree.parse(str(l5x_path), parser)
    root   = tree.getroot()
    rungs  = root.findall(".//Rung")

    # 3) Regex for COP/CPS/Message(Source,Dest,…)
    instr_re = re.compile(
        r'\b(COP|CPS|Message)\s*\(\s*([^,\s\)]+)\s*,\s*([^,\s\)]+)',
        re.IGNORECASE
    )

    records = []
    for tag in tags:
        found = False

        for rung in rungs:
            # — get context —
            rllcontent = rung.getparent()                   # <RLLContent>
            routine    = rllcontent.getparent()             # <Routine Name="...">
            routines   = routine.getparent()                # <Routines>
            program    = routines.getparent()               # <Program Name="...">

            prog_name  = program.get("Name")
            rout_name  = routine.get("Name")
            rung_num   = rung.get("Number")
            text_el    = rung.find("Text")

            if text_el is None or not text_el.text:
                continue

            for instr, src, dst in instr_re.findall(text_el.text):
                if dst == tag:
                    found = True
                    records.append({
                        "Col45":       tag,
                        "Program":     prog_name,
                        "Routine":     rout_name,
                        "Rung":        rung_num,
                        "Instruction": instr.upper(),
                        "Source":      src
                    })

        if not found:
            # Emit a “Not Found” row for this tag
            records.append({
                "Col45":       tag,
                "Program":     "",
                "Routine":     "",
                "Rung":        "",
                "Instruction": "Not Found",
                "Source":      ""
            })

    # 4) Build DataFrame, drop any exact dupes, and write out
    cols = ["Col45","Program","Routine","Rung","Instruction","Source"]
    df_out = pd.DataFrame(records, columns=cols).drop_duplicates()

    if df_out.empty:
        print("⚠️  No tags processed at all.")
    else:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Tag Mapping")
        print(f"✅  Wrote {len(df_out)} rows to '{output_path}'")

if __name__ == "__main__":
    extract_mappings(INPUT_EXCEL, INPUT_L5X, OUTPUT_FILE)
