from __future__ import annotations

import math
from typing import Any


class RoundingError(Exception):
    """Raised when rounding configuration is invalid."""


def get_rounding_increment(exercise_name: str, rounding_cfg: dict[str, Any]) -> float:
    """
    Return the rounding increment for an exercise.

    Example config:
    {
        "default": 2.5,
        "Squat": 2.5,
        "Weighted Pull-Ups": 1.25,
        "Weighted Dips": 1.25,
    }
    """
    if not isinstance(rounding_cfg, dict):
        raise RoundingError("rounding config must be a mapping")

    value = rounding_cfg.get(exercise_name, rounding_cfg.get("default", 2.5))

    try:
        increment = float(value)
    except (TypeError, ValueError) as exc:
        raise RoundingError(
            f"Invalid rounding increment for '{exercise_name}': {value!r}"
        ) from exc

    if increment <= 0:
        raise RoundingError(
            f"Rounding increment for '{exercise_name}' must be > 0, got {increment}"
        )

    return increment


def round_to_increment(value: float, increment: float, mode: str = "nearest") -> float:
    """
    Round a number to a given increment.

    Modes:
    - nearest
    - up
    - down

    Examples:
    round_to_increment(76.1, 1.25, "nearest") -> 76.25
    round_to_increment(76.1, 2.5, "down")     -> 75.0
    """
    if increment <= 0:
        raise RoundingError(f"increment must be > 0, got {increment}")

    scaled = value / increment

    if mode == "nearest":
        rounded = round(scaled) * increment
    elif mode == "up":
        rounded = math.ceil(scaled) * increment
    elif mode == "down":
        rounded = math.floor(scaled) * increment
    else:
        raise RoundingError(f"Unknown rounding mode: {mode!r}")

    # Avoid float junk like 75.0000000001
    return round(rounded, 6)
