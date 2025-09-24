#!/usr/bin/env python3

import re
from pathlib import Path
import pandas as pd
from lxml import etree

# —— Natural sort helpers (handle [idx] and .bit) ——
_tag_pat = re.compile(r'^(.*?)(?:\[(\d+)\])?(?:\.(\d+))?$')
def tag_sort_key(tag: str):
    m = _tag_pat.match(tag or "")
    base = m.group(1) if m else tag
    idx  = int(m.group(2)) if m and m.group(2) is not None else -1
    bit  = int(m.group(3)) if m and m.group(3) is not None else -1
    return (base, idx, bit)

def split_index(tag: str):
    """Return (base, idx) where idx defaults to 0 if absent."""
    m = re.match(r'^(.+?)\[(\d+)\]$', tag)
    return (m.group(1), int(m.group(2))) if m else (tag, 0)

def strip_index(tag: str):
    """Return base name with any [idx] removed."""
    return re.sub(r'\[\d+\]$', '', tag)

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

    # 1) Load Excel tags; keep both a set and a natural-sorted list
    df_tags = pd.read_excel(excel_path, engine="openpyxl")
    if "Col45" not in df_tags.columns:
        raise KeyError(f"'Col45' column not found. Available: {df_tags.columns.tolist()}")
    dest_tags_set = set(df_tags["Col45"].dropna().astype(str))
    tags_sorted   = sorted(dest_tags_set, key=tag_sort_key)

    # 2) Parse the L5X once
    parser = etree.XMLParser(remove_comments=False, recover=True)
    tree   = etree.parse(str(l5x_path), parser)
    root   = tree.getroot()

    # 3) Build maps: DataType, base Description, bit comments from Tag definitions, CONTROL LEN
    dtype_map = {}
    base_desc_map = {}
    bit_comment_map = {}   # key: "Base[idx].bit" -> comment text
    control_len_map = {}   # key: control tag name -> LEN (int)

    for t in root.findall(".//Tag"):
        base_name = t.get("Name")
        if not base_name:
            continue

        dt = (t.get("DataType") or "").upper()
        if dt:
            dtype_map[base_name] = dt

        desc_el = t.find("Description")
        if desc_el is not None:
            base_desc = (desc_el.get("Text") or desc_el.text or "").strip()
            if base_desc:
                base_desc_map[base_name] = base_desc

        # Bit-level comments: <Comments><Comment Operand="[i].bit">text</Comment>
        comments_el = t.find("Comments")
        if comments_el is not None:
            for c in comments_el.findall("Comment"):
                operand = (c.get("Operand") or "").strip()
                if not operand:
                    continue
                text = (c.get("Text") or c.text or "").strip()
                if not text:
                    continue
                key = f"{base_name}{operand}" if operand.startswith("[") else operand
                bit_comment_map[key] = text

        # CONTROL LEN
        if dt == "CONTROL":
            len_el = t.find(".//DataValueMember[@Name='LEN']")
            if len_el is not None:
                val = len_el.get("Value")
                try:
                    control_len_map[base_name] = int(val)
                except (TypeError, ValueError):
                    pass  # leave missing if not parseable

    # 4) Prepare regexes
    cop_re = re.compile(
        r'\b(COP|CPS)\s*\(\s*'      # Instruction
        r'([^,\s\)]+)\s*,\s*'       # Source (maybe with [i])
        r'([^,\s\)]+)\s*,\s*'       # Dest   (maybe with [i])
        r'(\d+)\s*\)',              # Length
        re.IGNORECASE
    )
    mov_re = re.compile(
        r'\bMOV\s*\(\s*'
        r'([^,\s\)]+)\s*,\s*'       # Source
        r'([^,\s\)]+)\s*\)',        # Destination
        re.IGNORECASE
    )
    ffl_re = re.compile(           # FFL(Source, Dest, Control, ...)
        r'\bFFL\s*\(\s*'
        r'([^,\s\)]+)\s*,\s*'       # Source
        r'([^,\s\)]+)\s*,\s*'       # Dest
        r'([^,\s\)]+)',             # Control
        re.IGNORECASE
    )
    ote_text_re = re.compile(
        r'\bOTE\s*\(\s*([^\s\)]+)\s*\)',  # OTE(operand) from rung text
        re.IGNORECASE
    )

    # 5) Walk every Rung; collect direct mappings + OTE operand index (and rung-level bit comments if present)
    records = []
    found_original_tags = set()   # items exactly as in Col45 (e.g., "X[7]") that we've mapped
    ote_map = {}                  # operand -> list of (Program, Routine, Rung)

    for rung in root.findall(".//Rung"):
        # context
        rll       = rung.getparent()
        rout      = rll.getparent()
        progs     = rout.getparent()
        prog      = progs.getparent()
        prog_name = prog.get("Name")
        rout_name = rout.get("Name")
        rung_num  = rung.get("Number")

        text_el = rung.find("Text")
        txt = text_el.text if (text_el is not None and text_el.text) else ""

        # Track OTE operands present on this rung (for existence / context)
        for operand in ote_text_re.findall(txt):
            operand = operand.strip()
            if operand:
                ote_map.setdefault(operand, []).append((prog_name, rout_name, rung_num))

        # Also scan structured rung XML for bit comments tied directly to operands (keep if tag-level missing)
        for elem in rung.findall(".//*[@Operand]"):
            operand = (elem.get("Operand") or "").strip()
            if not operand:
                continue
            com = elem.find("Comment")
            if com is not None:
                ctext = (com.get("Text") or com.text or "").strip()
                if ctext and operand not in bit_comment_map:
                    bit_comment_map[operand] = ctext

        # COP/CPS (expand by length)
        for instr, src_full, dst_full, length_s in cop_re.findall(txt):
            length = int(length_s)
            src_base, src_i0 = split_index(src_full)
            dst_base, dst_i0 = split_index(dst_full)
            dt = dtype_map.get(dst_base, "")
            base_desc = base_desc_map.get(dst_base, "")

            for k in range(length):
                dst_k = f"{dst_base}[{dst_i0 + k}]"
                if dst_k in dest_tags_set:
                    src_k = f"{src_base}[{src_i0 + k}]"
                    records.append({
                        "Col45":       dst_k,
                        "Description": base_desc,
                        "DataType":    dt,
                        "Program":     prog_name,
                        "Routine":     rout_name,
                        "Rung":        rung_num,
                        "Instruction": instr.upper(),
                        "Source":      src_k,
                    })
                    found_original_tags.add(dst_k)

        # MOV (no length)
        for src_full, dst_full in mov_re.findall(txt):
            if dst_full in dest_tags_set:
                dst_base, _ = split_index(dst_full)
                dt = dtype_map.get(dst_base, "")
                base_desc = base_desc_map.get(dst_base, "")
                records.append({
                    "Col45":       dst_full,
                    "Description": base_desc,
                    "DataType":    dt,
                    "Program":     prog_name,
                    "Routine":     rout_name,
                    "Rung":        rung_num,
                    "Instruction": "MOV",
                    "Source":      src_full,
                })
                found_original_tags.add(dst_full)

        # FFL: expand destination using LEN from control tag
        for src_full, dst_full, ctl_full in ffl_re.findall(txt):
            ctl_base = strip_index(ctl_full)
            ffl_len  = control_len_map.get(ctl_base)
            if not ffl_len:
                continue  # can't expand without LEN
            dst_base, dst_i0 = split_index(dst_full)
            dt = dtype_map.get(dst_base, "")
            base_desc = base_desc_map.get(dst_base, "")

            for k in range(ffl_len):
                dst_k = f"{dst_base}[{dst_i0 + k}]"
                if dst_k in dest_tags_set:
                    # Source is typically scalar; keep as-is
                    records.append({
                        "Col45":       dst_k,
                        "Description": base_desc,
                        "DataType":    dt,
                        "Program":     prog_name,
                        "Routine":     rout_name,
                        "Rung":        rung_num,
                        "Instruction": "FFL",
                        "Source":      src_full,
                    })
                    found_original_tags.add(dst_k)

        # MessageParameters (length via RequestedLength)
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
                dst_base, dst_i0 = m.group(1), int(m.group(2))
            else:
                dst_base, dst_i0 = local_elem, local_index
            dt = dtype_map.get(dst_base, "")
            base_desc = base_desc_map.get(dst_base, "")

            m2 = re.match(r'^(.+?)\[(\d+)\]$', remote_elem or "")
            if m2:
                src_base, src_i0 = m2.group(1), int(m2.group(2))
            else:
                src_base, src_i0 = (remote_elem or ""), local_index

            for k in range(length):
                dst_k = f"{dst_base}[{dst_i0 + k}]"
                if dst_k in dest_tags_set:
                    src_k = f"{src_base}[{src_i0 + k}]"
                    records.append({
                        "Col45":       dst_k,
                        "Description": base_desc,
                        "DataType":    dt,
                        "Program":     prog_name,
                        "Routine":     rout_name,
                        "Rung":        rung_num,
                        "Instruction": "MESSAGE",
                        "Source":      src_k,
                    })
                    found_original_tags.add(dst_k)

    # 6) Bitwise OTE sweep: only emit bits that actually have an OTE
    for orig in tags_sorted:
        if orig in found_original_tags:
            continue  # already mapped directly

        base, idx = split_index(orig)
        dt = dtype_map.get(base, "")
        if dt not in ("INT", "DINT"):
            continue

        max_bit   = 15 if dt == "INT" else 31
        base_desc = base_desc_map.get(base, "")
        emitted   = False

        for bit in range(max_bit + 1):
            operand = f"{base}[{idx}].{bit}"
            ctx = ote_map.get(operand)
            if not ctx:
                continue  # only add bits with OTE

            descr = bit_comment_map.get(operand, base_desc)
            prog_name, rout_name, rung_num = ctx[0]

            records.append({
                "Col45":       operand,
                "Description": descr,
                "DataType":    dt,
                "Program":     prog_name,
                "Routine":     rout_name,
                "Rung":        rung_num,
                "Instruction": "OTE",
                "Source":      "",
            })
            emitted = True

        if emitted:
            found_original_tags.add(orig)

    # 7) “Not Found” rows for anything never matched
    for tag in tags_sorted:
        if tag not in found_original_tags:
            base, _ = split_index(tag)
            dt = dtype_map.get(base, "")
            base_desc = base_desc_map.get(base, "")
            records.append({
                "Col45":       tag,
                "Description": base_desc,
                "DataType":    dt,
                "Program":     "",
                "Routine":     "",
                "Rung":        "",
                "Instruction": "Not Found",
                "Source":      "",
            })

    # 8) De-dup and natural-sort the entire output by Col45 (handles [i] and .bit)
    cols  = ["Col45","Description","DataType","Program","Routine","Rung","Instruction","Source"]
    df_out = pd.DataFrame(records, columns=cols).drop_duplicates()

    tmp = df_out["Col45"].str.extract(r'^(.*?)(?:\[(\d+)\])?(?:\.(\d+))?$')
    df_out["__base"] = tmp[0]
    df_out["__idx"]  = tmp[1].fillna(-1).astype(int)
    df_out["__bit"]  = tmp[2].fillna(-1).astype(int)
    df_out = df_out.sort_values(["__base","__idx","__bit"]).drop(columns=["__base","__idx","__bit"])

    # 9) Write to Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="Tag Mapping")

    print(f"✅  Wrote {len(df_out)} rows to '{output_path}'")

if __name__ == "__main__":
    extract_mappings(INPUT_EXCEL, INPUT_L5X, OUTPUT_FILE)
