"""
Configuration for the IPW coded lists workflow.

Values are persisted in config.json. This module exposes the current live values
as module globals and provides helpers for loading and saving them.
"""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

CONFIG_PATH = Path(__file__).resolve().with_name("config.json")

DEFAULTS: dict[str, Any] = {
    "DEFAULT_USER_EMAIL": "gciteli@ustravel.org",
    "DEFAULT_BOOTH_YEAR": "2026",
    "DEFAULT_IMPORT_SOURCE": "IPW",
    "MAX_DETECTED_AUDIENCE_LEN": 60,
    "AUDIENCE_SEGMENT_KEYWORDS": ["attendee", "media", "buyer", "exhibitor"],
    "AUDIENCE_SEGMENT_EXCLUDED_TERMS": ["(Coded)"],
    "VALID_YEARS": ["26", "27", "2026", "2027"],
    "SHOW_TRANSFORMING_FILES": False,
    "SHOW_CAPTURED_VALUES": False,
    "SHOW_PROCESSING_COMPANY_CLEANUP": False,
    "SHOW_PROCESSING_COMPANY_BOOTH": False,
    "SHOW_PROCESSING_INDIVIDUAL_BOOTH_CREDENTIALS": False,
    "SHOW_PROCESSING_CONTACT": False,
}

CONFIG_MENU_SECTIONS = [
    ("Workflow Defaults", ["DEFAULT_USER_EMAIL", "DEFAULT_BOOTH_YEAR", "DEFAULT_IMPORT_SOURCE"]),
    (
        "Audience Detection",
        [
            "MAX_DETECTED_AUDIENCE_LEN",
            "AUDIENCE_SEGMENT_KEYWORDS",
            "AUDIENCE_SEGMENT_EXCLUDED_TERMS",
            "VALID_YEARS",
        ],
    ),
    (
        "Output Display",
        [
            "SHOW_TRANSFORMING_FILES",
            "SHOW_CAPTURED_VALUES",
            "SHOW_PROCESSING_COMPANY_CLEANUP",
            "SHOW_PROCESSING_COMPANY_BOOTH",
            "SHOW_PROCESSING_INDIVIDUAL_BOOTH_CREDENTIALS",
            "SHOW_PROCESSING_CONTACT",
        ],
    ),
]

LIST_KEYS = {
    "AUDIENCE_SEGMENT_KEYWORDS",
    "AUDIENCE_SEGMENT_EXCLUDED_TERMS",
    "VALID_YEARS",
}
INT_KEYS = {"MAX_DETECTED_AUDIENCE_LEN"}
BOOL_KEYS = {
    "SHOW_TRANSFORMING_FILES",
    "SHOW_CAPTURED_VALUES",
    "SHOW_PROCESSING_COMPANY_CLEANUP",
    "SHOW_PROCESSING_COMPANY_BOOTH",
    "SHOW_PROCESSING_INDIVIDUAL_BOOTH_CREDENTIALS",
    "SHOW_PROCESSING_CONTACT",
}
STRING_KEYS = {
    "DEFAULT_USER_EMAIL",
    "DEFAULT_BOOTH_YEAR",
    "DEFAULT_IMPORT_SOURCE",
}


def _coerce_value(key: str, value: Any) -> Any:
    default = deepcopy(DEFAULTS[key])
    if key in STRING_KEYS:
        return str(value)
    if key in INT_KEYS:
        return int(value)
    if key in BOOL_KEYS:
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            lowered = value.strip().lower()
            if lowered in {"true", "1", "yes", "y", "on"}:
                return True
            if lowered in {"false", "0", "no", "n", "off"}:
                return False
        raise ValueError(f"Invalid boolean value for {key}: {value!r}")
    if key in LIST_KEYS:
        if not isinstance(value, list):
            raise ValueError(f"Invalid list value for {key}: {value!r}")
        return [str(item) for item in value]
    return default


def _load_raw_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return deepcopy(DEFAULTS)
    with CONFIG_PATH.open("r", encoding="utf-8") as handle:
        loaded = json.load(handle)
    if not isinstance(loaded, dict):
        raise ValueError("config.json must contain a JSON object.")
    merged = deepcopy(DEFAULTS)
    for key, value in loaded.items():
        if key in DEFAULTS:
            merged[key] = _coerce_value(key, value)
    return merged


def get_config_data() -> dict[str, Any]:
    return {key: deepcopy(globals()[key]) for key in DEFAULTS}


def save_config_data(data: dict[str, Any]) -> None:
    normalized: dict[str, Any] = {}
    for key, default_value in DEFAULTS.items():
        raw_value = data.get(key, deepcopy(default_value))
        normalized[key] = _coerce_value(key, raw_value)
    with CONFIG_PATH.open("w", encoding="utf-8") as handle:
        json.dump(normalized, handle, indent=2)
        handle.write("\n")
    reload_config()


def reload_config() -> None:
    data = _load_raw_config()
    for key, value in data.items():
        globals()[key] = value


def ensure_config_file() -> None:
    if not CONFIG_PATH.exists():
        save_config_data(deepcopy(DEFAULTS))
    else:
        reload_config()


ensure_config_file()
