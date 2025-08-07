import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd

# PARAMETERS you can tweak:
INPUT_DIR = r"C:\Users\dkelly\QCA Systems Ltd\CSX Curtis Bay Pier - Documents\Q-8690G - Site Wide Ignition Deployment\05 ENG AUTO\10 Conceptual\Program Exports"
OUTPUT_FILE = r"C:\Users\dkelly\Documents\message_summary.xlsx"

# The exact MessageParameters attributes you want to capture:
PARAMS = [
    "MessageType",
    "RemoteElement",
    "RequestedLength",
    "ConnectionPath",
    "CommTypeCode",
    "LocalIndex",
    "LocalElement",
]

def parse_l5x_file(filepath):
    """
    Parse one .L5X, return a DataFrame with one row per
    <Data Format="Message"> found, columns = PARAMS.
    """
    tree = ET.parse(filepath)
    root = tree.getroot()
    records = []

    # Namespace-agnostic tag matching:
    def is_data_message(elem):
        tag = elem.tag.split('}')[-1]
        return tag == "Data" and elem.attrib.get("Format") == "Message"

    for data in root.iter():
        if is_data_message(data):
            mp = data.find('.//MessageParameters')
            if mp is not None:
                # pull each attribute (None if missing)
                row = {p: mp.attrib.get(p) for p in PARAMS}
                records.append(row)

    return pd.DataFrame(records)

def build_message_workbook(input_dir, output_file):
    """
    Scans input_dir for .L5X files and writes an Excel workbook
    with one sheet per file (named after the file, truncated to 31 chars).
    """
    writer = pd.ExcelWriter(output_file, engine="openpyxl")

    for fullpath in glob.glob(os.path.join(input_dir, "*.L5X")):
        df = parse_l5x_file(fullpath)
        # Sheet names max out at 31 chars:
        sheet_name = os.path.splitext(os.path.basename(fullpath))[0][:31]
        # if no messages found, write an empty sheet with headers:
        if df.empty:
            df = pd.DataFrame(columns=PARAMS)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    # writer.save()
    writer.close()
    print(f"Wrote results to {output_file!r}")

if __name__ == "__main__":
    build_message_workbook(INPUT_DIR, OUTPUT_FILE)
