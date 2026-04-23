"""
Processing for derived outputs:
- \" - In Distro\": keep Booth ID + contact email; add Company Booth and Contact = \"In Distro List\".
- \" - Booths\": keep only Booth ID and de-duplicate rows.
"""

from pathlib import Path
from typing import List, Optional
import pandas as pd

from terminal_output import print_transforming_file_action


def _find_column(df: pd.DataFrame, candidates: List[str]) -> str:
    """
    Return the first matching column from candidates, using case/space-insensitive matching.
    """
    norm = {col.lower().replace(" ", ""): col for col in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "")
        if key in norm:
            return norm[key]
    return ""


def _report_issue(message: str, issues: Optional[List[str]] = None) -> None:
    if issues is not None:
        issues.append(message)
    print(f"Warning: {message}")

def _get_normalized_booth_id_column(df: pd.DataFrame) -> str:
    booth_col = _find_column(df, ["Booth ID", "BoothID", "Booth Id", "CmpLoginID"])
    if booth_col:
        df.rename(columns={booth_col: "Booth ID"}, inplace=True)
        return "Booth ID"
    return ""


def process_in_distro(filepath: Path, issues: Optional[List[str]] = None) -> Path:
    """
    Trim to Booth ID and email columns, add constant flag column, and overwrite the Excel file.
    """
    filepath = Path(filepath)
    df = pd.read_excel(filepath, dtype=str, keep_default_na=False)

    keep_cols: List[str] = []
    booth_col = _get_normalized_booth_id_column(df)
    if booth_col:
        keep_cols.append("Booth ID")
    else:
        _report_issue(
            "In Distro: Booth ID column was missing, so the output was created without Booth ID values.",
            issues=issues,
        )
    for cand in ("PerEmail", "PriEmail"):
        found = _find_column(df, [cand])
        if found:
            keep_cols.append(found)

    if not keep_cols:
        raise ValueError("Neither Booth ID nor PerEmail/PriEmail columns found.")

    df = df[keep_cols]
    df["Company Booth and Contact"] = "In Distro List"

    df.to_excel(filepath, index=False)
    print_transforming_file_action("Updated In Distro file saved", filepath)
    return filepath


def process_booths(filepath: Path, issues: List[str] | None = None) -> Path:
    """
    Keep only Booth ID (case/space-insensitive) and drop duplicate Booth IDs.
    Also ensure the final row is a sentinel Booth ID of "99999".
    If any existing 99999 rows are found (unexpected), remove them and record a warning.
    """
    filepath = Path(filepath)
    df = pd.read_excel(filepath, dtype=str, keep_default_na=False)

    booth_col = _get_normalized_booth_id_column(df)
    if not booth_col:
        _report_issue(
            "Booths: Booth ID column was missing, so the Booths output was reduced to the required sentinel row only.",
            issues=issues,
        )
        df = pd.DataFrame([{"Booth ID": "99999"}])
        df.to_excel(filepath, index=False)
        print_transforming_file_action("Updated Booths file saved", filepath)
        return filepath

    df = df[["Booth ID"]].drop_duplicates()

    # Normalize to strings for consistent comparisons.
    df["Booth ID"] = df["Booth ID"].astype(str)

    # Remove any existing 99999 rows; flag if encountered (unexpected).
    sentinel = "99999"
    removed_mask = df["Booth ID"] == sentinel
    removed_count = int(removed_mask.sum())
    if removed_count:
        df = df[~removed_mask]
        if issues is not None:
            issues.append(
                f"Booths: removed {removed_count} unexpected Booth ID {sentinel} row(s) before adding final sentinel."
            )

    # Append the sentinel row at the end.
    df = pd.concat([df, pd.DataFrame([{"Booth ID": sentinel}])], ignore_index=True)

    df.to_excel(filepath, index=False)
    print_transforming_file_action("Updated Booths file saved", filepath)
    return filepath


def count_booths(filepath: Path) -> int:
    df = pd.read_excel(filepath, dtype=str, keep_default_na=False)
    return len(df.index)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process derived 'In Distro' or 'Booths' Excel file in-place.")
    parser.add_argument("filepath", type=Path, help="Path to the Excel file.")
    parser.add_argument("--mode", choices=["in_distro", "booths"], default="in_distro")
    args = parser.parse_args()

    if args.mode == "in_distro":
        process_in_distro(args.filepath)
    else:
        process_booths(args.filepath)
