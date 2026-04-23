from pathlib import Path
from typing import Union

import config

_transforming_files_header_printed = False


def reset_transforming_files_section() -> None:
    global _transforming_files_header_printed
    _transforming_files_header_printed = False


def _display_file_label(path: Union[str, Path]) -> str:
    stem = Path(path).stem
    return stem.rsplit(" ; ", 1)[-1]


def print_transforming_file_action(action: str, path: Union[str, Path]) -> None:
    global _transforming_files_header_printed
    if not _transforming_files_header_printed:
        print("\n--- Transforming Files ---")
        _transforming_files_header_printed = True

    if config.SHOW_TRANSFORMING_FILES:
        print(f"{action}: {_display_file_label(path)}")
