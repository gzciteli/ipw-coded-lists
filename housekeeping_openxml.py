import csv
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


def _col_to_index(col_letters: str) -> int:
    idx = 0
    for c in col_letters:
        idx = idx * 26 + (ord(c.upper()) - ord("A") + 1)
    return idx - 1  # zero-based


def convert_openxml_to_csv(xlsx_path: Path, csv_path: Path) -> None:
    """Convert first worksheet of an .xlsx/.xlsm file to CSV using stdlib."""
    with zipfile.ZipFile(xlsx_path, "r") as z:
        # Shared strings (optional)
        shared_strings = []
        try:
            with z.open("xl/sharedStrings.xml") as f:
                tree = ET.parse(f)
            for si in tree.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                text_parts = []
                for t in si.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
                    text_parts.append(t.text or "")
                shared_strings.append("".join(text_parts))
        except KeyError:
            pass

        # First worksheet
        with z.open("xl/worksheets/sheet1.xml") as f:
            sheet_tree = ET.parse(f)

    sheet_root = sheet_tree.getroot()
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows_out = []

    for row in sheet_root.findall(".//a:sheetData/a:row", ns):
        # Determine row index to handle gaps if needed
        cells = row.findall("a:c", ns)
        if not cells:
            rows_out.append([])
            continue

        max_col = 0
        parsed_cells = []
        for cell in cells:
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            col_idx = _col_to_index(match.group(1)) if match else len(parsed_cells)
            if col_idx > max_col:
                max_col = col_idx

            value = ""
            cell_type = cell.attrib.get("t")
            v = cell.find("a:v", ns)

            if cell_type == "s" and v is not None:
                ss_idx = int(v.text)
                value = shared_strings[ss_idx] if ss_idx < len(shared_strings) else ""
            elif cell_type == "inlineStr":
                t = cell.find(".//a:t", ns)
                value = t.text or "" if t is not None else ""
            elif v is not None and v.text is not None:
                # Do NOT interpret or reformat numbers; keep raw stored text
                value = v.text
            else:
                value = ""

            parsed_cells.append((col_idx, value))

        # Build the row with proper spacing
        row_list = ["" for _ in range(max_col + 1)]
        for col_idx, val in parsed_cells:
            row_list[col_idx] = val
        rows_out.append(row_list)

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows_out)
