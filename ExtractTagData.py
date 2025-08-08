#!/usr/bin/env python3

import re
from pathlib import Path

import pandas as pd
from lxml import etree

# — Natural sort helper —
_tag_pat = re.compile(r'^(.*?)(?:\[(\d+)\])?$')
def tag_key(tag: str):
    m = _tag_pat.match(tag)
    if not m:
        return (tag, -1)  # no index → before [0]
    base, idx = m.group(1), m.group(2)
    return (base, int(idx) if idx is not None else -1)

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

    # 1) Load and sort your Excel tags in natural order
    df_tags = pd.read_excel(excel_path, engine="openpyxl")
    if "Col45" not in df_tags.columns:
        raise KeyError(f"'Col45' column not found. Available: {df_tags.columns.tolist()}")
    dest_tags = set(df_tags["Col45"].dropna().astype(str))
    tags      = sorted(dest_tags, key=tag_key)

    # 2) Parse the L5X once
    parser = etree.XMLParser(remove_comments=False, recover=True)
    tree   = etree.parse(str(l5x_path), parser)
    root   = tree.getroot()

    # 3) Prepare COP/CPS regex
    cop_re = re.compile(
        r'\b(COP|CPS)\s*\(\s*'      # Instruction
        r'([^,\s\)]+)\s*,\s*'       # Source tag (with [idx])
        r'([^,\s\)]+)\s*,\s*'       # Dest tag (with [idx])
        r'(\d+)\s*\)',              # Length
        re.IGNORECASE
    )

    records = []
    found   = set()

    # 4) Walk every Rung
    for rung in root.findall(".//Rung"):
        # context
        rll       = rung.getparent()
        rout      = rll.getparent()
        progs     = rout.getparent()
        prog      = progs.getparent()
        prog_name = prog.get("Name")
        rout_name = rout.get("Name")
        rung_num  = rung.get("Number")

        # 4a) COP / CPS in ladder text
        text_el = rung.find("Text")
        if text_el is not None and text_el.text:
            for instr, src_full, dst_full, length_s in cop_re.findall(text_el.text):
                length = int(length_s)

                def split_base(txt):
                    m = re.match(r'^(.+?)\[(\d+)\]$', txt)
                    return (m.group(1), int(m.group(2))) if m else (txt, 0)

                src_base, src_idx0 = split_base(src_full)
                dst_base, dst_idx0 = split_base(dst_full)

                for i in range(length):
                    dst_i = f"{dst_base}[{dst_idx0 + i}]"
                    if dst_i in dest_tags:
                        src_i = f"{src_base}[{src_idx0 + i}]"
                        records.append({
                            "Col45":       dst_i,
                            "Program":     prog_name,
                            "Routine":     rout_name,
                            "Rung":        rung_num,
                            "Instruction": instr.upper(),
                            "Source":      src_i
                        })
                        found.add(dst_i)

        # 4b) MessageParameters in XML
        for mp in rung.findall(".//MessageParameters"):
            local_elem    = mp.get("LocalElement")
            req_len       = mp.get("RequestedLength")
            local_index_s = mp.get("LocalIndex", "0")
            remote_elem   = mp.get("RemoteElement")
            if not local_elem or not req_len:
                continue

            length      = int(req_len)
            local_index = int(local_index_s)

            m = re.match(r'^(.+?)\[(\d+)\]$', local_elem)
            if m:
                dst_base, dst_idx0 = m.group(1), int(m.group(2))
            else:
                dst_base, dst_idx0 = local_elem, local_index

            m2 = re.match(r'^(.+?)\[(\d+)\]$', remote_elem or "")
            if m2:
                src_base, src_idx0 = m2.group(1), int(m2.group(2))
            else:
                src_base, src_idx0 = remote_elem or "", local_index

            for i in range(length):
                dst_i = f"{dst_base}[{dst_idx0 + i}]"
                if dst_i in dest_tags:
                    src_i = f"{src_base}[{src_idx0 + i}]"
                    records.append({
                        "Col45":       dst_i,
                        "Program":     prog_name,
                        "Routine":     rout_name,
                        "Rung":        rung_num,
                        "Instruction": "MESSAGE",
                        "Source":      src_i
                    })
                    found.add(dst_i)

    # 5) Not Found rows
    for tag in tags:
        if tag not in found:
            records.append({
                "Col45":       tag,
                "Program":     "",
                "Routine":     "",
                "Rung":        "",
                "Instruction": "Not Found",
                "Source":      ""
            })

    # 6) Drop dupes and natural-sort the entire output
    df_out = (pd.DataFrame(records,
                           columns=["Col45","Program","Routine","Rung","Instruction","Source"])
              .drop_duplicates())

    tmp = df_out["Col45"].str.extract(r'^(.*?)(?:\[(\d+)\])?$')
    df_out["__base"] = tmp[0]
    df_out["__idx"]  = tmp[1].fillna(-1).astype(int)
    df_out = df_out.sort_values(["__base","__idx"]).drop(columns=["__base","__idx"])

    # 7) Write to Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="Tag Mapping")

    print(f"✅  Wrote {len(df_out)} rows to '{output_path}'")

if __name__ == "__main__":
    extract_mappings(INPUT_EXCEL, INPUT_L5X, OUTPUT_FILE)
