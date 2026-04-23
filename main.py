import config
import ipw_workflow


def _print_main_menu() -> None:
    print("\n--- Main Menu ---")
    print("(1) Run IPW Coded List Workflow")
    print("(2) Edit Config")
    print("(3) Help")
    print("(4) Exit")


def _print_help_screen() -> None:
    print("\n--- Help ---")
    print("This is the HELP screen.")
    print("Developed by Gabriel Z. Citeli March 2026")


def _build_config_entries() -> list[tuple[int, str]]:
    entries: list[tuple[int, str]] = []
    index = 1
    for _, keys in config.CONFIG_MENU_SECTIONS:
        for key in keys:
            entries.append((index, key))
            index += 1
    return entries


def _print_config_menu() -> dict[str, str]:
    print("\n--- Edit Config ---")
    print("Key: a = append, r = remove, c = clear")
    print()
    option_map: dict[str, str] = {}
    index = 1
    for section_title, keys in config.CONFIG_MENU_SECTIONS:
        print(section_title)
        for key in keys:
            value = getattr(config, key)
            if key in config.LIST_KEYS:
                print(f"({index}a, {index}r, {index}c) {key} = {value!r}")
                option_map[f"{index}a"] = key
                option_map[f"{index}r"] = key
                option_map[f"{index}c"] = key
            else:
                print(f"({index}) {key} = {value!r}")
                option_map[str(index)] = key
            index += 1
        print()
    return option_map


def _parse_bool(raw_value: str) -> bool:
    lowered = raw_value.strip().lower()
    if lowered in {"true", "1", "yes", "y", "on"}:
        return True
    if lowered in {"false", "0", "no", "n", "off"}:
        return False
    raise ValueError("Enter true/false, yes/no, on/off, or 1/0.")


def _edit_scalar_config(key: str) -> bool:
    current = getattr(config, key)
    raw_value = input(f"New value for {key} [current {current!r}] (Enter to cancel): ").strip()
    if raw_value == "":
        return False

    if key in config.INT_KEYS:
        new_value = int(raw_value)
    elif key in config.BOOL_KEYS:
        new_value = _parse_bool(raw_value)
    else:
        new_value = raw_value

    updated = config.get_config_data()
    updated[key] = new_value
    config.save_config_data(updated)
    print(f"Updated {key} to {getattr(config, key)!r}.")
    return True


def _append_list_value(key: str) -> bool:
    raw_value = input(f"Append value to {key} (Enter to cancel): ").strip()
    if raw_value == "":
        return False
    updated = config.get_config_data()
    updated[key].append(raw_value)
    config.save_config_data(updated)
    print(f"Appended {raw_value!r} to {key}.")
    return True


def _remove_list_value(key: str) -> bool:
    current = list(getattr(config, key))
    if not current:
        print(f"{key} is already empty.")
        return False

    print(f"Current {key}:")
    for idx, value in enumerate(current):
        print(f"{idx}. {value}")

    raw_value = input(f"Remove by value or 0-based index from {key} (Enter to cancel): ").strip()
    if raw_value == "":
        return False

    updated = config.get_config_data()
    values = list(updated[key])
    if raw_value.isdigit():
        remove_index = int(raw_value)
        if remove_index < 0 or remove_index >= len(values):
            raise ValueError(f"Index out of range for {key}.")
        removed = values.pop(remove_index)
    else:
        if raw_value not in values:
            raise ValueError(f"{raw_value!r} was not found in {key}.")
        values.remove(raw_value)
        removed = raw_value
    updated[key] = values
    config.save_config_data(updated)
    print(f"Removed {removed!r} from {key}.")
    return True


def _clear_list_value(key: str) -> bool:
    confirm = input(f"Type CONFIRM to clear {key}: ").strip()
    if confirm != "CONFIRM":
        print("Clear cancelled.")
        return False
    updated = config.get_config_data()
    updated[key] = []
    config.save_config_data(updated)
    print(f"Cleared {key}.")
    return True


def edit_config_menu() -> None:
    while True:
        option_map = _print_config_menu()
        choice = input("Choose a config option (Enter to return): ").strip().lower()
        if choice == "":
            return
        if choice not in option_map:
            print("Invalid config option.")
            continue

        key = option_map[choice]
        try:
            if key in config.LIST_KEYS:
                action = choice[-1]
                if action == "a":
                    changed = _append_list_value(key)
                elif action == "r":
                    changed = _remove_list_value(key)
                else:
                    changed = _clear_list_value(key)
            else:
                changed = _edit_scalar_config(key)
        except ValueError as exc:
            print(f"Invalid value: {exc}")
            continue

        if changed:
            return


def main() -> None:
    while True:
        _print_main_menu()
        choice = input("Choose an option: ").strip().lower()
        print("---")
        if choice == "1":
            print("Starting IPW Coded List Workflow...")
            print("---")
            try:
                ipw_workflow.main()
            except RuntimeError as exc:
                print(str(exc))
                print("---")
        elif choice == "2":
            print()
            edit_config_menu()
        elif choice == "3" or choice == "help":
            print()
            _print_help_screen()
        elif choice == "4":
            print("Exiting.")
            return
        else:
            print("Invalid option.")


if __name__ == "__main__":
    main()
