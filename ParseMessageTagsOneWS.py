import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd

# PARAMETERS you can tweak:
INPUT_DIR = r"C:\Users\dkelly\QCA Systems Ltd\CSX Curtis Bay Pier - Documents\Q-8690G - Site Wide Ignition Deployment\05 ENG AUTO\10 Conceptual\Program Exports"
OUTPUT_FILE = r"C:\Users\dkelly\QCA Systems Ltd\CSX Curtis Bay Pier - Documents\Q-8690G - Site Wide Ignition Deployment\05 ENG AUTO\10 Conceptual\PLC Message Summary Dump One WS.xlsx"

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

ALL_COLS = ["SourceFile"] + PARAMS  # single worksheet columns


def parse_l5x_file(filepath):
    """
    Parse one .L5X, return a list of dicts:
      - first row: only SourceFile populated (filename row)
      - subsequent rows: one per <Data Format="Message"> with params
    """
    rows = []

    # Filename-only row
    rows.append({"SourceFile": os.path.basename(filepath), **{p: None for p in PARAMS}})

    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
    except ET.ParseError as e:
        # Add a diagnostic row and return
        rows.append({
            "SourceFile": os.path.basename(filepath) + " (PARSE ERROR)",
            **{p: None for p in PARAMS}
        })
        return rows

    # Namespace-agnostic tag matcher
    def is_data_message(elem):
        tag = elem.tag.split('}')[-1]
        return tag == "Data" and elem.attrib.get("Format") == "Message"

    for data in root.iter():
        if is_data_message(data):
            mp = data.find('.//MessageParameters')
            if mp is not None:
                row = {"SourceFile": os.path.basename(filepath)}
                for p in PARAMS:
                    row[p] = mp.attrib.get(p)
                rows.append(row)

    # If no messages were found, keep just the filename row (already added)
    return rows


def build_message_workbook(input_dir, output_file):
    """
    Scans input_dir for .L5X files and writes a single Excel worksheet
    ('Messages') containing:
      - a filename-only row
      - followed by that file's message rows (if any)
    Repeats for each file. All rows end up on one sheet.
    """
    all_rows = []

    files = sorted(glob.glob(os.path.join(input_dir, "*.L5X")))
    for fullpath in files:
        all_rows.extend(parse_l5x_file(fullpath))

        # Optional: spacer row between files (uncomment if desired)
        # all_rows.append({"SourceFile": None, **{p: None for p in PARAMS}})

    # Build DataFrame and enforce column ordering
    df = pd.DataFrame(all_rows, columns=ALL_COLS)

    # Write to single sheet
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Messages", index=False)

    print(f"Processed {len(files)} file(s). Wrote results to {output_file!r}")


if __name__ == "__main__":
    build_message_workbook(INPUT_DIR, OUTPUT_FILE)
