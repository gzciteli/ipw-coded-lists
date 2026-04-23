import sys
import re
from pathlib import Path

import config
from date_utils import validate_and_normalize_date, looks_like_date_fragment
from housekeeping import prepare_destination_package, prompt_for_existing_source_file
from associations_processing import process_contact_file, finalize_column_names
from inDistro_processing import process_in_distro, process_booths, count_booths
import pandas as pd
import shutil
from terminal_output import print_transforming_file_action, reset_transforming_files_section


def load_source_headers(path: Path):
    return list(pd.read_excel(path, nrows=0).columns)

def _normalize_whitespace(text: str) -> str:
    return " ".join(text.split())


def _clean_audience_segment(text: str) -> str:
    cleaned = text
    for term in config.AUDIENCE_SEGMENT_EXCLUDED_TERMS:
        cleaned = re.sub(re.escape(term), "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*(?:-|;)\s*$", "", cleaned)
    return _normalize_whitespace(cleaned)


def _match_leading_date(stem: str) -> tuple[str, str]:
    match = re.match(r"^\s*(\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4}))(?:\s*(?:-|;)\s*|\s+)(.+?)\s*$", stem)
    if not match:
        return "", stem.strip()

    date_text = match.group(1)
    remainder = match.group(2).strip()
    if looks_like_date_fragment(date_text) and remainder:
        return validate_and_normalize_date(date_text), remainder

    return "", stem.strip()


def _match_trailing_date(stem: str) -> tuple[str, str]:
    match = re.match(r"^\s*(.+?)(?:\s*(?:-|;)\s*|\s+)(\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4}))\s*$", stem)
    if not match:
        return "", stem.strip()

    remainder = match.group(1).strip()
    date_text = match.group(2)
    if remainder and looks_like_date_fragment(date_text):
        return validate_and_normalize_date(date_text), remainder

    return "", stem.strip()


def _strip_boundary_dates(stem: str) -> tuple[str, str]:
    """
    Return (detected_date, audience_candidate).
    Detected date prefers a leading date, then a trailing date.
    Audience candidate strips valid dates from either boundary.
    """
    cleaned = stem.strip()
    if not cleaned:
        return "", ""

    leading_date, without_leading = _match_leading_date(cleaned)
    trailing_date, without_trailing = _match_trailing_date(cleaned)

    detected_date = leading_date or trailing_date
    audience_candidate = cleaned

    if leading_date:
        audience_candidate = without_leading
    trailing_date_after_leading, without_trailing_after_leading = _match_trailing_date(audience_candidate)
    if trailing_date_after_leading:
        audience_candidate = without_trailing_after_leading
    elif not leading_date and trailing_date:
        audience_candidate = without_trailing

    return detected_date, _normalize_whitespace(audience_candidate)


def _looks_like_audience_segment(text: str) -> bool:
    lowered = text.lower()
    return any(keyword in lowered for keyword in config.AUDIENCE_SEGMENT_KEYWORDS)


def _parse_filename_defaults(path: Path) -> tuple[str, str]:
    stem = path.stem.strip()
    if not stem:
        return "", ""

    detected_date, audience_candidate = _strip_boundary_dates(stem)
    audience_candidate = _clean_audience_segment(audience_candidate)
    if _looks_like_audience_segment(audience_candidate):
        return detected_date, _truncate_detected_audience_segment(audience_candidate)

    return detected_date, ""


def _truncate_detected_audience_segment(text: str, limit: int | None = None) -> str:
    if limit is None:
        limit = config.MAX_DETECTED_AUDIENCE_LEN
    cleaned = _normalize_whitespace(text)
    if len(cleaned) <= limit:
        return cleaned

    prefix = cleaned[:limit]
    if len(cleaned) > limit and not cleaned[limit].isspace() and not prefix.endswith(" "):
        if " " in prefix:
            prefix = prefix.rsplit(" ", 1)[0]
        else:
            prefix = ""

    prefix = prefix.rstrip()
    if not prefix:
        prefix = cleaned[:limit].rstrip()

    return f"{prefix}..."


def ensure_booth_id_source_available(path: Path) -> None:
    try:
        headers = load_source_headers(path)
    except PermissionError as exc:
        raise RuntimeError(
            f'Could not open "{path}". Close the Excel file and try again. '
            "The workflow returned to the main menu without making changes."
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Could not inspect source workbook headers: {exc}") from exc

    has_booth_id = "CmpLoginID" in headers or "Booth ID" in headers
    has_person_login = "PerLoginID" in headers or "PriLoginID" in headers

    if not has_booth_id and not has_person_login:
        raise RuntimeError(
            "Booth ID is missing from the source file, and no Person LoginID column was found to generate it. "
            "Expected CmpLoginID/Booth ID or PerLoginID/PriLoginID. The program exited before making any changes."
        )


def prompt_member():
    while True:
        ans = input("Contact - Is Member? (Yes/No, Enter to leave empty) ").strip().lower()
        if ans == "":
            return ""
        if ans in {"yes", "y"}:
            return "Yes"
        if ans in {"no", "n"}:
            return "No"
        print("Please enter Yes or No.")


def detect_booth_label_default(path: Path) -> str:
    """
    Inspect the first sheet headers and suggest a default booth label.
    - If PriFirstName present (and PerFirstName absent), suggest Primary Contact.
    - If PerFirstName present (and PriFirstName absent), suggest Staff Contact.
    - Otherwise, return empty string (no suggestion).
    """
    try:
        cols = load_source_headers(path)
    except Exception:
        return ""
    has_per = "PerFirstName" in cols
    has_pri = "PriFirstName" in cols
    if has_pri and not has_per:
        return "Primary Contact"
    if has_per and not has_pri:
        return "Staff Contact"
    return ""


def prompt_booth_label(default_label: str = ""):
    prompt_base = "Booth - Booth Label"
    if default_label:
        prompt_text = f"{prompt_base} [detected {default_label} | (P)rimary Contact / (S)taff Contact]: "
    else:
        prompt_text = f"{prompt_base} (P)rimary Contact / (S)taff Contact: "

    while True:
        ans = input(prompt_text).strip().lower()
        if ans == "" and default_label:
            return default_label
        if ans in {"primary contact", "primary", "p"}:
            return "Primary Contact"
        if ans in {"staff contact", "staff", "s"}:
            return "Staff Contact"
        print("Please enter Primary Contact or Staff Contact.")


def main():
    file_location: Path = prompt_for_existing_source_file()
    ensure_booth_id_source_available(file_location)

    detected_label = detect_booth_label_default(file_location)
    detected_date, detected_audience_segment = _parse_filename_defaults(file_location)

    while True:
        if detected_date:
            run_date_raw = input(f"Targeted distribution date [detected {detected_date}]: ").strip() or detected_date
        else:
            run_date_raw = input("Targeted distribution date (format M.D.Y): ").strip()
            if run_date_raw == "":
                print("Targeted distribution date cannot be empty.")
                continue
        try:
            run_date = validate_and_normalize_date(run_date_raw)
            break
        except ValueError as exc:
            print(f"Invalid date: {exc}")

    if detected_audience_segment:
        recipient = input(
            f'Targeted audience segment [detected "{detected_audience_segment}"]: '
        ).strip() or detected_audience_segment
    else:
        recipient = input("Targeted audience segment: ").strip()
    while recipient == "":
        if detected_audience_segment:
            recipient = input(
                f'Targeted audience segment [detected "{detected_audience_segment}"] (cannot be empty): '
            ).strip() or detected_audience_segment
        else:
            recipient = input("Targeted audience segment (cannot be empty): ").strip()

    user_email = input(f"Your email [default {config.DEFAULT_USER_EMAIL}]: ").strip() or config.DEFAULT_USER_EMAIL

    contact_is_member = prompt_member()

    booth_year_raw = input(f"Booth - Booth Year [default {config.DEFAULT_BOOTH_YEAR}]: ").strip()
    booth_year = booth_year_raw if booth_year_raw else config.DEFAULT_BOOTH_YEAR

    import_type = input("Booth - IMPORT Type (leave blank to skip): ").strip()
    if import_type:
        import_category = input("Booth - IMPORT Category (blank allowed): ").strip()
    else:
        import_category = ""

    booth_label = prompt_booth_label(detected_label)

    issues = []
    reset_transforming_files_section()

    # Create the final destination package and the processing copy inside it.
    dest_dir, _archival_path, final_path = prepare_destination_package(file_location, run_date, recipient)

    # Process Excel in place.
    process_contact_file(
        final_path,
        owner_email=user_email,
        is_member=contact_is_member,
        import_source=config.DEFAULT_IMPORT_SOURCE,
        booth_year=booth_year,
        booth_label=booth_label,
        import_type=import_type,
        import_category=import_category,
        issues=issues,
    )

    # Create the derived outputs beside the processing copy.
    base_stem = final_path.stem
    ext = final_path.suffix
    base_prefix = base_stem.rsplit(" ; ", 1)[0] if " ; " in base_stem else base_stem

    associations_path = final_path
    in_distro_path = associations_path.with_name(f"{base_prefix} ; 2 In Distro{ext}")
    shutil.copyfile(associations_path, in_distro_path)
    print_transforming_file_action("Created", in_distro_path)

    # Booths derived from In Distro (further trimmed in inDistro_processing).
    booths_path = associations_path.with_name(f"{base_prefix} ; 3 Booths{ext}")
    shutil.copyfile(in_distro_path, booths_path)
    print_transforming_file_action("Created", booths_path)

    # Post-process derived files.
    process_in_distro(in_distro_path, issues=issues)
    process_booths(booths_path, issues=issues)
    booth_count = count_booths(booths_path)
    counted_booths_path = booths_path.with_name(f"{base_prefix} ; 3 Booths({booth_count}){ext}")
    if booths_path != counted_booths_path:
        if counted_booths_path.exists():
            counted_booths_path.unlink()
        booths_path.replace(counted_booths_path)
        booths_path = counted_booths_path
        print_transforming_file_action("Renamed Booths file to", booths_path)
    finalize_column_names(associations_path)

    if config.SHOW_CAPTURED_VALUES:
        print("\n--- Captured Values ---")
        print(f"Original File: {file_location}")
        print(f"Output Folder: {dest_dir}")
        print(f"Date: {run_date}")
        print(f"Email: {user_email}")
        print(f"Contact Is Member?: {contact_is_member}")
        print(f"Booth Year: {booth_year}")
        print(f"IMPORT Type: {import_type}")
        print(f"IMPORT Category: {import_category}")
        print(f"Booth Label: {booth_label}")
    print("\nIssues detected:")
    if issues:
        for issue in issues:
            print(f"- {issue}")
    else:
        print("No issues detected.")
    print("---")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nCancelled by user")
    except RuntimeError as exc:
        sys.exit(str(exc))
