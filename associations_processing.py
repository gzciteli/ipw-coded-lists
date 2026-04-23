"""
Excel-based processing (pandas) for IPW data.

Workflow:
- Reads an Excel file (first sheet).
- Company cleanup: removes selected company columns (non-fatal if missing) and reports status.
- Company Booth processing: validates required company booth fields and adds booth-level columns.
- Individual Booth Credentials: validates required person login fields, duplicates org info per person, and adds person-level booth columns.
- Contact processing: reports required Contact columns (non-fatal) and adds constant-value columns.
- Overwrites the Excel file in place.
"""

from pathlib import Path
from typing import Dict, Iterable, List, Tuple, Optional
import pandas as pd

import config
from terminal_output import print_transforming_file_action


# Company columns to remove (display -> header)
COMPANY_REMOVE: Dict[str, str] = {
    "Company Type": "CmpType",
    "City": "CmpCity",
    "State": "CmpState",
    "Country": "CmpCountry",
    "CmpLgcNum": "CmpLgcNum",
}

# Display name -> sheet header mapping for required columns
CONTACT_REQUIRED: Dict[str, List[str]] = {
    "First Name": ["PerFirstName", "PriFirstName"],
    "Last Name": ["PerLastName", "PriLastName"],
    "Title": ["PerTitle", "PriTitle"],
    "Email": ["PerEmail", "PriEmail"],
}

# Display name -> sheet header mapping for columns we add
CONTACT_ADDITIONS: Dict[str, str] = {
    "Contact Owner": "Contact Owner",
    "IMPORT Source": "IMPORT Source",
    "Is Member?": "Is Member?",
}

# Company Booth requirements and additions
COMPANY_BOOTH_REQUIRED: Dict[str, str] = {
    "Organization Name": "CmpName",
    "Organization LoginID": "CmpLoginID",
    "Organization Pwd": "CmpPwd",
}

COMPANY_BOOTH_ADDITIONS: Dict[str, str] = {
    "Booth Year - Company": "Booth Year - Company",
    "IMPORT Type": "IMPORT Type",
    "IMPORT Category": "IMPORT Category",
    "Booth Owner": "Booth Owner",
    "Booth Label": "Booth Label",
}

# Individual Booth credentials requirements and additions
INDIVIDUAL_BOOTH_REQUIRED: Dict[str, List[str]] = {
    "Person LoginID": ["PerLoginID", "PriLoginID"],
    "Person Pwd": ["PerPwd", "PriPwd"],
}

INDIVIDUAL_BOOTH_ADDITIONS: Dict[str, str] = {
    "Org (Booth) Name - Person": "Org (Booth) Name - Person",
    "Org LoginID - Person": "Org LoginID - Person",
    "Booth Year - Person": "Booth Year - Person",
    "Credential Owner": "Credential Owner",
    "Individual Credentials Label": "Individual Credentials Label",
}

def finalize_column_names(filepath: Path) -> Path:
    """
    Rename selected columns at the very end of processing to avoid breaking upstream logic.
    - Booth Label -> Company Booth and Contact
    - Individual Credentials Label -> Individual Booth Credential and Company Booth
    """
    df = pd.read_excel(filepath, dtype=str, keep_default_na=False)
    rename_map = {
        "Booth Label": "Company Booth and Contact",
        "Individual Credentials Label": "Individual Booth Credential and Company Booth",
    }
    # Only rename columns that are present
    existing_map = {k: v for k, v in rename_map.items() if k in df.columns}
    if existing_map:
        df = df.rename(columns=existing_map)
        df.to_excel(filepath, index=False)
        print_transforming_file_action("Finalized column names in", filepath)
    return filepath


def _print_status_block(title: str, lines: Iterable[str], show_details: bool = True) -> None:
    print(f"\n--- Processing: {title} ---")
    if show_details:
        for line in lines:
            print(line)
        print("---")


def _find_first_present(headers: List[str], candidates: List[str]) -> str:
    for name in candidates:
        if name in headers:
            return name
    return ""


def _append_issue_once(issues: Optional[List[str]], message: str) -> None:
    if issues is None:
        return
    if message not in issues:
        issues.append(message)


def _generate_company_login_id_from_person_login(person_login: str) -> str:
    """
    Build a company Booth ID from a person login value.
    This is intentionally isolated so the derivation rule can be swapped later.
    """
    return str(person_login or "").strip().split(".", 1)[0]


def _ensure_company_login_id(
    headers: List[str],
    rows: List[Dict[str, str]],
    issues: Optional[List[str]] = None,
) -> Tuple[List[str], str]:
    """
    Ensure CmpLoginID exists, deriving it from a person login column when needed.
    Returns (headers, source_description).
    """
    if "CmpLoginID" in headers:
        return headers, "existing"

    person_login_src = _find_first_present(headers, ["PerLoginID", "PriLoginID"])
    if not person_login_src:
        return headers, ""

    headers.append("CmpLoginID")
    for row in rows:
        row["CmpLoginID"] = _generate_company_login_id_from_person_login(row.get(person_login_src, ""))

    _append_issue_once(
        issues,
        f"Created Company ID (CmpLoginID) from {person_login_src} by taking the text before the first period.",
    )
    return headers, f"derived from {person_login_src}"


def process_company_cleanup(headers: List[str], rows: List[Dict[str, str]]) -> Tuple[List[str], Dict[str, str]]:
    """
    Remove selected company columns if present, report status.
    """
    status: Dict[str, str] = {}
    for display, header in COMPANY_REMOVE.items():
        if header in headers:
            headers = [h for h in headers if h != header]
            for row in rows:
                row.pop(header, None)
            status[display] = "removed"
        else:
            status[display] = "missing"

    lines = [f"{display}: {status[display]}" for display in COMPANY_REMOVE.keys()]
    _print_status_block("Company Cleanup", lines, show_details=config.SHOW_PROCESSING_COMPANY_CLEANUP)
    return headers, status


def process_company_booth(
    headers: List[str],
    rows: List[Dict[str, str]],
    booth_year: str,
    booth_owner_email: str,
    booth_label: str,
    import_type: str = "",
    import_category: str = "",
    issues: Optional[List[str]] = None,
) -> Tuple[List[str], Dict[str, str]]:
    """
    Validate required company booth columns and add booth-related columns.
    """
    status: Dict[str, str] = {}

    headers, login_id_source = _ensure_company_login_id(headers, rows, issues=issues)

    # Required columns
    for display, header in COMPANY_BOOTH_REQUIRED.items():
        status[display] = "Present" if header in headers else "Missing"

    if login_id_source and login_id_source != "existing":
        status["Organization LoginID"] = f"Generated ({login_id_source})"

    if status["Organization LoginID"] == "Missing" and status["Organization Pwd"] == "Missing":
        _append_issue_once(
            issues,
            "Missing company booth credential columns: CmpLoginID (Booth ID / Organization LoginID) and CmpPwd (Organization Pwd). Booth outputs were limited.",
        )
    else:
        if status["Organization LoginID"] == "Missing":
            _append_issue_once(
                issues,
                "Missing company booth credential column: CmpLoginID (Booth ID / Organization LoginID). Booth outputs were limited.",
            )
        if status["Organization Pwd"] == "Missing":
            _append_issue_once(
                issues,
                "Missing company booth credential column: CmpPwd (Organization Pwd). Booth outputs may be incomplete.",
            )

    # Additions (always add Booth Year, Owner, Label)
    for display, header, value in [
        ("Booth Year - Company", COMPANY_BOOTH_ADDITIONS["Booth Year - Company"], booth_year),
        ("Booth Owner", COMPANY_BOOTH_ADDITIONS["Booth Owner"], booth_owner_email),
        ("Booth Label", COMPANY_BOOTH_ADDITIONS["Booth Label"], booth_label),
    ]:
        if header in headers and issues is not None:
            issues.append(f"Unexpected pre-existing column: {header}")
        if header not in headers:
            headers.append(header)
        for row in rows:
            row[header] = value
        status[display] = "added"

    # Conditional additions
    if import_type:
        header = COMPANY_BOOTH_ADDITIONS["IMPORT Type"]
        if header in headers and issues is not None:
            issues.append(f"Unexpected pre-existing column: {header}")
        if header not in headers:
            headers.append(header)
        for row in rows:
            row[header] = import_type
        status["IMPORT Type"] = "added"
    else:
        status["IMPORT Type"] = "excluded"

    if import_category:
        header = COMPANY_BOOTH_ADDITIONS["IMPORT Category"]
        if header in headers and issues is not None:
            issues.append(f"Unexpected pre-existing column: {header}")
        if header not in headers:
            headers.append(header)
        for row in rows:
            row[header] = import_category
        status["IMPORT Category"] = "added"
    else:
        status["IMPORT Category"] = "excluded"

    lines = []
    for display in COMPANY_BOOTH_REQUIRED.keys():
        lines.append(f"{display}: {status[display]}")
    lines.append(f"Booth Year - Company: {status['Booth Year - Company']}")
    lines.append(f"Booth Owner: {status['Booth Owner']}")
    lines.append(f"Booth Label: {status['Booth Label']}")
    lines.append(f"IMPORT Type: {status['IMPORT Type']}")
    lines.append(f"IMPORT Category: {status['IMPORT Category']}")
    _print_status_block("Company Booth", lines, show_details=config.SHOW_PROCESSING_COMPANY_BOOTH)

    return headers, status


def process_individual_booth_credentials(
    headers: List[str],
    rows: List[Dict[str, str]],
    booth_year: str,
    booth_owner_email: str,
    booth_label: str,
    issues: Optional[List[str]] = None,
) -> Tuple[List[str], Dict[str, str]]:
    """
    Validate required person login columns (with aliases) and add per-person booth columns.
    """
    status: Dict[str, str] = {}

    # Required with aliases
    for display, candidates in INDIVIDUAL_BOOTH_REQUIRED.items():
        found = _find_first_present(headers, candidates)
        status[display] = f"Present ({found})" if found else "Missing"

    # Additions: clone org info and add booth/personal fields
    org_name_src = _find_first_present(headers, ["CmpName"])
    org_login_src = _find_first_present(headers, ["CmpLoginID"])

    # Org name clone
    org_name_header = INDIVIDUAL_BOOTH_ADDITIONS["Org (Booth) Name - Person"]
    if org_name_header in headers and issues is not None:
        issues.append(f"Unexpected pre-existing column: {org_name_header}")
    if org_name_header not in headers:
        headers.append(org_name_header)
    if org_name_src:
        for row in rows:
            row[org_name_header] = row.get(org_name_src, "")
        status["Org (Booth) Name - Person"] = f"added (from {org_name_src})"
    else:
        for row in rows:
            row[org_name_header] = ""
        status["Org (Booth) Name - Person"] = "added (source missing)"

    # Org login clone
    org_login_header = INDIVIDUAL_BOOTH_ADDITIONS["Org LoginID - Person"]
    if org_login_header in headers and issues is not None:
        issues.append(f"Unexpected pre-existing column: {org_login_header}")
    if org_login_header not in headers:
        headers.append(org_login_header)
    if org_login_src:
        for row in rows:
            row[org_login_header] = row.get(org_login_src, "")
        status["Org LoginID - Person"] = f"added (from {org_login_src})"
    else:
        for row in rows:
            row[org_login_header] = ""
        status["Org LoginID - Person"] = "added (source missing)"

    # Constant additions
    for display, header, value in [
        ("Booth Year - Person", INDIVIDUAL_BOOTH_ADDITIONS["Booth Year - Person"], booth_year),
        ("Credential Owner", INDIVIDUAL_BOOTH_ADDITIONS["Credential Owner"], booth_owner_email),
        ("Individual Credentials Label", INDIVIDUAL_BOOTH_ADDITIONS["Individual Credentials Label"], booth_label),
    ]:
        if header in headers and issues is not None:
            issues.append(f"Unexpected pre-existing column: {header}")
        if header not in headers:
            headers.append(header)
        for row in rows:
            row[header] = value
        status[display] = "added"

    lines = []
    for display in INDIVIDUAL_BOOTH_REQUIRED.keys():
        lines.append(f"{display}: {status[display]}")
    for display in [
        "Org (Booth) Name - Person",
        "Org LoginID - Person",
        "Booth Year - Person",
        "Credential Owner",
        "Individual Credentials Label",
    ]:
        lines.append(f"{display}: {status[display]}")

    _print_status_block(
        "Individual Booth Credentials",
        lines,
        show_details=config.SHOW_PROCESSING_INDIVIDUAL_BOOTH_CREDENTIALS,
    )
    return headers, status


def _load_excel(filepath: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    df = pd.read_excel(filepath, dtype=str, keep_default_na=False)
    headers = list(df.columns)
    rows = df.to_dict(orient="records")
    return headers, rows


def _write_excel(filepath: Path, headers: List[str], rows: List[Dict[str, str]]) -> None:
    df = pd.DataFrame(rows, columns=headers)
    df.to_excel(filepath, index=False)


def process_contact_rows(
    headers: List[str],
    rows: List[Dict[str, str]],
    owner_email: str,
    is_member: str = "",
    import_source: str = "",
    issues: Optional[List[str]] = None,
) -> Tuple[List[str], Dict[str, str]]:
    status: Dict[str, str] = {}
    import_source = import_source or config.DEFAULT_IMPORT_SOURCE

    # Required column checks
    contact_sources: Dict[str, str] = {}
    for display, candidates in CONTACT_REQUIRED.items():
        found = _find_first_present(headers, candidates)
        contact_sources[display] = found
        status[display] = f"Present ({found})" if found else "Missing"

    # Additions
    # Contact Owner
    owner_col = CONTACT_ADDITIONS["Contact Owner"]
    if owner_col in headers and issues is not None:
        issues.append(f"Unexpected pre-existing column: {owner_col}")
    if owner_col not in headers:
        headers.append(owner_col)
    for row in rows:
        row[owner_col] = owner_email
    status["Contact Owner"] = "added"

    # IMPORT Source
    import_col = CONTACT_ADDITIONS["IMPORT Source"]
    if import_col in headers and issues is not None:
        issues.append(f"Unexpected pre-existing column: {import_col}")
    if import_col not in headers:
        headers.append(import_col)
    for row in rows:
        row[import_col] = import_source
    status["IMPORT Source"] = "added"

    # Optional Is Member?
    member_col = CONTACT_ADDITIONS["Is Member?"]
    if is_member and is_member.strip():
        if member_col in headers and issues is not None:
            issues.append(f"Unexpected pre-existing column: {member_col}")
        if member_col not in headers:
            headers.append(member_col)
        for row in rows:
            row[member_col] = is_member
        status["Is Member?"] = "added"
    else:
        status["Is Member?"] = "excluded"

    # Print user-friendly block
    lines = []
    for display in CONTACT_REQUIRED.keys():
        lines.append(f"{display}: {status[display]}")
    lines.append(f"Contact Owner: {status['Contact Owner']}")
    lines.append(f"IMPORT Source: {status['IMPORT Source']}")
    lines.append(f"Is Member?: {status['Is Member?']}")
    _print_status_block("Contact", lines, show_details=config.SHOW_PROCESSING_CONTACT)

    return headers, status


def process_contact_file(
    filepath: Path,
    owner_email: str,
    is_member: str = "",
    import_source: str = "",
    booth_year: str = "",
    booth_label: str = "",
    import_type: str = "",
    import_category: str = "",
    issues: Optional[List[str]] = None,
) -> Path:
    """
    Load Excel, process steps, and overwrite the file.
    """
    if issues is None:
        issues = []
    import_source = import_source or config.DEFAULT_IMPORT_SOURCE
    filepath = Path(filepath)
    headers, rows = _load_excel(filepath)
    headers, _ = process_company_cleanup(headers, rows)
    headers, _ = process_company_booth(
        headers,
        rows,
        booth_year=booth_year,
        booth_owner_email=owner_email,
        booth_label=booth_label,
        import_type=import_type,
        import_category=import_category,
        issues=issues,
    )
    headers, _ = process_individual_booth_credentials(
        headers,
        rows,
        booth_year=booth_year,
        booth_owner_email=owner_email,
        booth_label=booth_label,
        issues=issues,
    )
    headers, _ = process_contact_rows(
        headers,
        rows,
        owner_email=owner_email,
        is_member=is_member,
        import_source=import_source,
        issues=issues,
    )
    _write_excel(filepath, headers, rows)
    print_transforming_file_action("Updated file saved", filepath)
    return filepath


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process IPW Excel file in-place (first sheet).")
    parser.add_argument("filepath", type=Path, help="Path to the Excel file to process.")
    parser.add_argument("--owner-email", required=True, help="Value for Contact Owner column.")
    parser.add_argument("--is-member", default="", help="Value for Is Member? column (leave blank to skip adding).")
    parser.add_argument("--import-source", default=config.DEFAULT_IMPORT_SOURCE, help="Value for IMPORT Source column.")

    args = parser.parse_args()
    process_contact_file(args.filepath, owner_email=args.owner_email, is_member=args.is_member, import_source=args.import_source)
