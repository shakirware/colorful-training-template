from pathlib import Path
from typing import Any

import yaml


DATA_DIR = Path("data")


class ConfigError(Exception):
    """Raised when config files are missing or invalid."""


def load_yaml(path: str | Path) -> Any:
    path = Path(path)

    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")

    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if data is None:
        raise ConfigError(f"Config file is empty: {path}")

    return data


def load_training_maxes() -> dict[str, float]:
    data = load_yaml(DATA_DIR / "training_maxes.yaml")

    if not isinstance(data, dict):
        raise ConfigError("training_maxes.yaml must be a mapping of exercise -> number")

    cleaned: dict[str, float] = {}
    for exercise, value in data.items():
        if not isinstance(exercise, str):
            raise ConfigError("All training max keys must be strings")

        try:
            cleaned[exercise] = float(value)
        except (TypeError, ValueError) as exc:
            raise ConfigError(
                f"Training max for '{exercise}' must be numeric, got {value!r}"
            ) from exc

        if cleaned[exercise] <= 0:
            raise ConfigError(
                f"Training max for '{exercise}' must be > 0, got {cleaned[exercise]}"
            )

    return cleaned


def load_settings() -> dict[str, Any]:
    data = load_yaml(DATA_DIR / "settings.yaml")

    if not isinstance(data, dict):
        raise ConfigError("settings.yaml must be a mapping")

    required_keys = ["start_date", "output_yaml", "output_workbook"]
    for key in required_keys:
        if key not in data:
            raise ConfigError(f"settings.yaml is missing required key: {key}")

    rounding = data.get("rounding", {})
    if rounding and not isinstance(rounding, dict):
        raise ConfigError("settings.rounding must be a mapping if provided")

    return data


def load_program() -> list[dict[str, Any]]:
    data = load_yaml(DATA_DIR / "program.yaml")

    if not isinstance(data, list):
        raise ConfigError("program.yaml must be a list of week objects")

    for week in data:
        if not isinstance(week, dict):
            raise ConfigError("Each week entry in program.yaml must be a mapping")

    return data
