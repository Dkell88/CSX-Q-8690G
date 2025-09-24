import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd

# PARAMETERS you can tweak:
INPUT_DIR = r"C:\Users\dkelly\QCA Systems Ltd\CSX Curtis Bay Pier - Documents\Q-8690G - Site Wide Ignition Deployment\05 ENG AUTO\10 Conceptual\Program Exports"
OUTPUT_FILE = r"C:\Users\dkelly\Documents\All Messages Summary.xlsx"
sheet_name = "All Messages"

# The exact MessageParameters attributes you want to capture:
PARAMS = [
    "MessageType",
    "RemoteElement",
    "RequestedLength",
    "ConnectionPath",
    "CommTypeCode",
    "LocalIndex",
    "LocalElement",
    "SourceFile"
]

def parse_l5x_file(filepath):
    """
    Parse one .L5X, return a DataFrame with one row per
    <Data Format="Message"> found, columns = PARAMS.
    """
    try: 
        tree = ET.parse(filepath)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"[WARN] SKipping {filepath}: XML parse erraor: {e}")
        return []
    records = []
    # Namespace-agnostic tag matching:
    def localname(elem):
       return elem.split('}',1)[-1]
    for data in root.iter():
        if localname(data.tag) == "Data" and data.attrib.get("Format") == "Message":
            mp = None
            for child in data.iter():
                if localname(child.tag) == "MessageParameters":
                    mp = child
                    break
            if mp is not None:
                # pull each attribute (None if missing)
                row = {p: mp.attrib.get(p) for p in PARAMS}
                row["Sourse File"] = os.path.basename(filepath)
                records.append(row)
    # print(f"{records}")
    return pd.DataFrame(records)

def build_message_workbook(input_dir, output_file):
    """
    Scans input_dir for .L5X files and writes an Excel workbook
    with one sheet per file (named after the file, truncated to 31 chars).
    """
    allRows = []
    #writer = pd.ExcelWriter(output_file, engine="openpyxl")

    for fullpath in glob.glob(os.path.join(input_dir, "*.L5X")):
        rows = parse_l5x_file(fullpath)
        if not rows.empty:
            allRows.extend(rows)
        else:
            print(f'Skipping file: {fullpath}')
    print(f"{allRows}")
    if allRows:
        df = pd.DataFrame(allRows, columns = PARAMS)
    else: 
    # if no messages found, write an empty sheet with headers:
        df = pd.DataFrame(columns=PARAMS)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    # writer.save()
    writer.close()
    print(f"Wrote results to {output_file!r} total of {len(df)}")

if __name__ == "__main__":
    build_message_workbook(INPUT_DIR, OUTPUT_FILE)
