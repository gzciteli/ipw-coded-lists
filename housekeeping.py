"""
Housekeeping utilities for preparing source files.

- Accepts Excel (.xlsx, .xlsm, .xlsb, .xls).
- Prints the original filename.
- Suggests a new filename by dropping a trailing date segment (if present)
  and prefixing with the user-supplied date.
 - Stages working files in the local "source" directory and consumes
  the original input file.
"""

import shutil
from pathlib import Path
from typing import Union

from date_utils import looks_like_date_fragment, validate_and_normalize_date
from terminal_output import print_transforming_file_action

# Excel-only workflow now that pandas is available.
ALLOWED_EXTS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
SOURCE_DIR = Path(__file__).resolve().parent / "source"


def _prompt_file_name_choice() -> str:
    """
    Prompt for Original / Suggested / New.
    Pressing Enter accepts the suggested file name.
    """
    while True:
        choice = input("Use (O)riginal, (S)uggested [Enter], or (N)ew file name? ").strip().lower()
        if choice == "":
            return "suggested"
        if choice in {"o", "s", "n", "original", "suggested", "new"}:
            return choice


def _normalize_path_input(path_str: Union[str, Path]) -> Path:
    """
    Accept quoted Windows paths copied from Explorer and try sensible fallbacks:
    - strip surrounding quotes
    - resolve relative paths against CWD
    - if still missing, look for the filename inside SOURCE_DIR
    """
    if isinstance(path_str, Path):
        candidate = path_str.expanduser()
    else:
        cleaned = path_str.strip().strip('"').strip("'")
        candidate = Path(cleaned).expanduser()
    if not candidate.is_absolute():
        candidate = (Path.cwd() / candidate).resolve()
    if candidate.exists():
        return candidate

    # If not found, try assuming the file is already in SOURCE_DIR
    alt = (SOURCE_DIR / candidate.name).resolve()
    if alt.exists():
        return alt

    return candidate


def ensure_source_dir() -> None:
    """Create the source directory if it doesn't exist."""
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)


def prompt_for_existing_source_file() -> Path:
    """
    Prompt repeatedly until the user provides an existing Excel file.
    The file may be anywhere; it will be moved into SOURCE_DIR later for processing.
    """
    ensure_source_dir()
    while True:
        raw = input("File path (Excel files only): ").strip()
        candidate = _normalize_path_input(raw)

        if not candidate.exists():
            print(f"File not found: {candidate}. Please verify the path and try again.")
            continue

        if candidate.suffix.lower() not in ALLOWED_EXTS:
            print("File must be Excel (.xlsx, .xlsm, .xlsb, .xls).")
            continue

        return candidate


def suggest_new_name(original_path: Path, user_date: str) -> str:
    """
    Build a suggested filename:
    - If the stem ends with _<date>, drop that trailing date.
    - Prefix the (possibly trimmed) stem with the validated user_date.
    - Preserve the original extension.
    """
    stem = original_path.stem
    suffix = original_path.suffix

    base_stem = stem
    if "_" in stem:
        head, tail = stem.rsplit("_", 1)
        if looks_like_date_fragment(tail):
            base_stem = head

    return f"{user_date}_{base_stem}{suffix}"


def rename_into_source_deprecated(original_path_str: str, user_date: str) -> Path:
    """
    Legacy renamer: preserves original stem with date prefix.
    """
    ensure_source_dir()

    original_path = _normalize_path_input(original_path_str)
    if not original_path.exists():
        raise FileNotFoundError(f"File not found: {original_path}")
    if original_path.suffix.lower() not in ALLOWED_EXTS:
        raise ValueError("File must be CSV or Excel (.csv, .xlsx, .xlsm, .xlsb, .xls).")

    # Validate date (in case caller forgot)
    user_date = validate_and_normalize_date(user_date)

    print(f"Original file name: {original_path.name}")
    suggested_name = suggest_new_name(original_path, user_date)
    print(f"Suggested new name: {suggested_name}")

    while True:
        choice = _prompt_file_name_choice()
        if choice.startswith("o"):
            candidate = original_path.name
        elif choice.startswith("s"):
            candidate = suggested_name
        else:  # new
            while True:
                custom_name = input("Type the new file name (no path): ").strip().strip('"').strip("'")
                if custom_name:
                    candidate = Path(custom_name).name
                    if Path(candidate).suffix == "":
                        candidate += original_path.suffix
                    break
                print("Please enter a file name (cannot be empty).")

        target_path = SOURCE_DIR / candidate

        if target_path.exists():
            confirm = input(f"{target_path.name} already exists in 'source'. Overwrite? [y/N]: ").strip().lower()
            if confirm in {"y", "yes"}:
                break
            # User chose not to overwrite; keep the existing file and continue.
            print("Keeping existing file.")
            return target_path
        break

    if original_path.resolve() == target_path.resolve():
        print("File already at desired path; no move needed.")
        return target_path

    target_path.parent.mkdir(parents=True, exist_ok=True)
    original_path.replace(target_path)
    print_transforming_file_action("Moved to", target_path)
    return target_path


def rename_into_source(original_path_str: str, user_date: str, recipient: str) -> Path:
    """
    New renamer: builds suggested name as "<date> ; <recipient> ; 1 Associations<ext>",
    lets the user confirm or override, then moves into source dir.
    """
    ensure_source_dir()

    original_path = _normalize_path_input(original_path_str)
    if not original_path.exists():
        raise FileNotFoundError(f"File not found: {original_path}")
    if original_path.suffix.lower() not in ALLOWED_EXTS:
        raise ValueError("File must be Excel (.xlsx, .xlsm, .xlsb, .xls).")

    user_date = validate_and_normalize_date(user_date)
    clean_recipient = recipient.strip().strip('"').strip("'")
    default_name = f"{user_date} ; {clean_recipient} ; 1 Associations{original_path.suffix}"

    print(f"Original file name: {original_path.name}")
    suggested_name = default_name
    print(f"Suggested new name: {suggested_name}")

    while True:
        choice = _prompt_file_name_choice()
        if choice.startswith("o"):
            candidate = original_path.name
        elif choice.startswith("s"):
            candidate = suggested_name
        else:  # new
            while True:
                custom_name = input("Type the new file name (no path): ").strip().strip('"').strip("'")
                if custom_name:
                    candidate = Path(custom_name).name
                    if Path(candidate).suffix == "":
                        candidate += original_path.suffix
                    break
                print("Please enter a file name (cannot be empty).")

        target_path = SOURCE_DIR / candidate

        if target_path.exists():
            confirm = input(f"{target_path.name} already exists in 'source'. Overwrite? [y/N]: ").strip().lower()
            if confirm in {"y", "yes"}:
                break
            print("Keeping existing file.")
            return target_path
        break

    if original_path.resolve() == target_path.resolve():
        print("File already at desired path; no move needed.")
        return target_path

    target_path.parent.mkdir(parents=True, exist_ok=True)
    original_path.replace(target_path)
    print_transforming_file_action("Moved to", target_path)
    return target_path


def stage_working_file(original_path_str: Union[str, Path], user_date: str, recipient: str) -> Path:
    """
    Move the workbook into SOURCE_DIR/_working for in-place processing.
    The original input file is consumed.
    """
    ensure_source_dir()

    original_path = _normalize_path_input(original_path_str)
    if not original_path.exists():
        raise FileNotFoundError(f"File not found: {original_path}")
    if original_path.suffix.lower() not in ALLOWED_EXTS:
        raise ValueError("File must be Excel (.xlsx, .xlsm, .xlsb, .xls).")

    user_date = validate_and_normalize_date(user_date)
    clean_recipient = recipient.strip().strip('"').strip("'")
    suggested_name = f"{user_date} ; {clean_recipient} ; 1 Associations{original_path.suffix}"
    staging_dir = SOURCE_DIR / "_working"

    print(f"Original file name: {original_path.name}")
    print(f"Suggested new name: {suggested_name}")

    while True:
        choice = _prompt_file_name_choice()
        if choice.startswith("o"):
            candidate = original_path.name
        elif choice.startswith("s"):
            candidate = suggested_name
        else:
            while True:
                custom_name = input("Type the new file name (no path): ").strip().strip('"').strip("'")
                if custom_name:
                    candidate = Path(custom_name).name
                    if Path(candidate).suffix == "":
                        candidate += original_path.suffix
                    break
                print("Please enter a file name (cannot be empty).")

        target_path = staging_dir / candidate

        if original_path.resolve() == target_path.resolve():
            print("The selected name matches the original file. Please choose a different file name.")
            continue

        if target_path.exists():
            confirm = input(f"{target_path.name} already exists in staging. Overwrite? [y/N]: ").strip().lower()
            if confirm in {"y", "yes"}:
                break
            print("Please choose a different file name.")
            continue
        break

    target_path.parent.mkdir(parents=True, exist_ok=True)
    original_path.replace(target_path)
    print_transforming_file_action("Working file staged", target_path)
    return target_path


def prepare_destination_package(original_path_str: Union[str, Path], user_date: str, recipient: str) -> tuple[Path, Path, Path]:
    """
    Create the final destination folder beside the original workbook and prepare:
    - a moved archival copy named "(Original) <original file name>"
    - a processing copy named "<date> ; <recipient> ; 1 Associations<ext>"
    Returns (destination_dir, archival_path, processing_copy_path).
    """
    original_path = _normalize_path_input(original_path_str)
    if not original_path.exists():
        raise FileNotFoundError(f"File not found: {original_path}")
    if original_path.suffix.lower() not in ALLOWED_EXTS:
        raise ValueError("File must be Excel (.xlsx, .xlsm, .xlsb, .xls).")

    user_date = validate_and_normalize_date(user_date)
    clean_recipient = recipient.strip().strip('"').strip("'")
    output_root = original_path.parent
    dest_dir = output_root / f"{user_date} ; {clean_recipient}"
    archival_path = dest_dir / f"(Original) {original_path.name}"
    processing_copy_path = dest_dir / f"{user_date} ; {clean_recipient} ; 1 Associations{original_path.suffix}"

    if dest_dir.exists():
        confirm = input(
            f"{dest_dir.name} already exists. Overwrite the entire destination package? [y/N]: "
        ).strip().lower()
        if confirm not in {"y", "yes"}:
            raise RuntimeError(f"Cancelled. Kept existing destination package: {dest_dir.name}")
        if dest_dir.resolve().parent != output_root.resolve():
            raise RuntimeError(f"Refusing to overwrite unexpected destination folder: {dest_dir}")
        shutil.rmtree(dest_dir)

    dest_dir.mkdir(parents=True, exist_ok=True)
    original_path.replace(archival_path)
    print_transforming_file_action("Archived original as", archival_path)

    shutil.copy2(archival_path, processing_copy_path)
    print_transforming_file_action("Created", processing_copy_path)
    return dest_dir, archival_path, processing_copy_path


def _convert_excel_to_csv_if_needed(path: Path) -> Path:
    """
    If path is an Excel file, convert first sheet to CSV (same directory/stem),
    remove the original Excel file, and return the CSV path.
    """
    if path.suffix.lower() == ".csv":
        return path

    try:
        print(f"Excel file provided ({path.name}); converting to CSV...")

        # 1) Pure stdlib converter for .xlsx/.xlsm (sheet1 only)
        if path.suffix.lower() in {".xlsx", ".xlsm"}:
            try:
                csv_path = path.with_suffix(".csv")
                convert_openxml_to_csv(path, csv_path)
                path.unlink(missing_ok=True)
                print_transforming_file_action("Converted Excel to CSV (built-in)", csv_path)
                return csv_path
            except PermissionError as exc:
                raise RuntimeError("Excel file appears to be open. Please close it and try again.") from exc
            except OSError as exc:
                # Handle WinError 32 explicitly (file in use)
                if getattr(exc, "winerror", None) == 32:
                    raise RuntimeError("Excel file appears to be open. Please close it and try again.") from exc
                print(f"Built-in converter failed: {exc}")
            except Exception as exc:
                print(f"Built-in converter failed: {exc}")

        # 2) pandas if available
        try:
            import pandas as pd
        except ImportError:
            pd = None

        if pd is not None:
            df = pd.read_excel(path, dtype=str, keep_default_na=False)
            csv_path = path.with_suffix(".csv")
            df.to_csv(csv_path, index=False)
            path.unlink(missing_ok=True)
            print_transforming_file_action("Converted Excel to CSV via pandas", csv_path)
            return csv_path

        # No safe fallback without pandas; abort to avoid corrupting values.
        raise RuntimeError("Conversion needs pandas (or provide CSV). Install pandas or save as CSV manually.")
    except Exception as exc:
        raise RuntimeError(f"Failed to convert Excel to CSV: {exc}")


def main() -> None:
    source_file = prompt_for_existing_source_file()

    while True:
        date_input = input("Targeted distribution date (format M.D.Y): ").strip()
        try:
            normalized_date = validate_and_normalize_date(date_input)
            break
        except ValueError as exc:
            print(f"Invalid date: {exc}")

    recipient = input("Targeted audience segment: ").strip()
    while recipient == "":
        recipient = input("Targeted audience segment (cannot be empty): ").strip()

    final_path = stage_working_file(source_file, normalized_date, recipient)
    print_transforming_file_action("Working file ready", final_path)


if __name__ == "__main__":
    main()
