from __future__ import annotations

from copy import deepcopy
from typing import Any

from colorful_training_template.rounding import (
    get_rounding_increment,
    round_to_increment,
)
from colorful_training_template.rep_table import (
    RepTableError,
    calculate_weight_from_rep_table,
    get_rep_table_percent_for_reps,
)


EXERCISE_ALIASES = {
    "Back Squats": "Squat",
    "Back Squat": "Squat",
    "Paused Squats": "Squat",
    "Paused Squat": "Squat",
    "Dynamic Effort Squats": "Squat",
    "Dynamic Effort Squat": "Squat",
    "Weighted (Ring) Dips": "Weighted Dips",
    "Weighted (Ring) Dip": "Weighted Dips",
}


class CalculationError(Exception):
    """Raised when workout calculation fails."""


def canonical_exercise_name(name: str) -> str:
    return EXERCISE_ALIASES.get(name, name)


def parse_reps_field(reps_value: Any) -> tuple[int, int]:
    """
    Parse reps field like:
    - '4x4' -> (4, 4)
    - '1x3' -> (1, 3)

    Returns:
        (num_sets, reps_per_set)
    """
    if isinstance(reps_value, int):
        return 1, reps_value

    if not isinstance(reps_value, str):
        raise CalculationError(f"Invalid reps value: {reps_value!r}")

    cleaned = reps_value.strip().lower()
    if "x" not in cleaned:
        raise CalculationError(f"Reps must be in 'NxR' format, got: {reps_value!r}")

    left, right = cleaned.split("x", 1)

    try:
        num_sets = int(left.strip())
        reps_per_set = int(right.strip())
    except ValueError as exc:
        raise CalculationError(f"Could not parse reps field: {reps_value!r}") from exc

    if num_sets <= 0 or reps_per_set <= 0:
        raise CalculationError(f"Reps must be positive, got: {reps_value!r}")

    return num_sets, reps_per_set


def expand_sets(sets: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Expand set prescriptions like:
    {'reps': '4x4', 'percentage_1rm': 60}
    into four individual sets with reps=4.
    """
    expanded: list[dict[str, Any]] = []

    for set_data in sets:
        if not isinstance(set_data, dict):
            raise CalculationError(f"Each set must be a mapping, got: {set_data!r}")

        if "reps" not in set_data:
            raise CalculationError(f"Missing 'reps' in set: {set_data!r}")

        num_sets, reps_per_set = parse_reps_field(set_data["reps"])

        for _ in range(num_sets):
            new_set = deepcopy(set_data)
            new_set["reps"] = reps_per_set
            expanded.append(new_set)

    return expanded


def format_weight(weight: float) -> str:
    if float(weight).is_integer():
        return f"{int(weight)}kg"
    return f"{weight:.2f}".rstrip("0").rstrip(".") + "kg"


def validate_percentage(percentage: Any, exercise_name: str) -> float:
    try:
        pct = float(percentage)
    except (TypeError, ValueError) as exc:
        raise CalculationError(
            f"percentage_1rm for '{exercise_name}' must be numeric, got {percentage!r}"
        ) from exc

    if pct <= 0 or pct > 110:
        raise CalculationError(
            f"percentage_1rm for '{exercise_name}' must be > 0 and <= 110, got {pct}"
        )

    return pct


def resolve_calculation_mode(
    *,
    set_data: dict[str, Any],
    exercise: dict[str, Any],
    settings: dict[str, Any],
) -> str:
    """
    Priority:
    1. set-level calculation_mode
    2. exercise-level calculation_mode
    3. global settings calculation_mode
    """
    mode = (
        set_data.get("calculation_mode")
        or exercise.get("calculation_mode")
        or settings.get("calculation_mode", "training_max")
    )

    if mode not in {"training_max", "rep_table"}:
        raise CalculationError(f"Unknown calculation mode: {mode!r}")

    return mode


def calculate_set_weight_training_max_mode(
    *,
    exercise_name: str,
    percentage_1rm: float,
    training_maxes: dict[str, float],
    rounding_cfg: dict[str, float],
) -> str:
    canonical_name = canonical_exercise_name(exercise_name)

    if canonical_name not in training_maxes:
        raise CalculationError(
            f"No training max found for exercise '{exercise_name}' "
            f"(canonical: '{canonical_name}')"
        )

    training_max = float(training_maxes[canonical_name])
    increment = float(get_rounding_increment(canonical_name, rounding_cfg))

    raw_weight = training_max * (percentage_1rm / 100.0)
    rounded_weight = round_to_increment(raw_weight, increment, mode="nearest")

    return format_weight(rounded_weight)


def calculate_set_weight_rep_table_mode(
    *,
    exercise_name: str,
    reps: int,
    percentage_1rm: float,
    rounding_cfg: dict[str, float],
    rep_table_dir: str,
) -> tuple[str, float]:
    """
    Returns:
        (formatted_weight, rep_table_percent_for_reps)

    In rep_table mode:
    - percentage_1rm selects the nearest base-weight row in the .xlsx table
    - reps selects the reps column
    - the table cell is the planned weight
    """
    canonical_name = canonical_exercise_name(exercise_name)
    increment = float(get_rounding_increment(canonical_name, rounding_cfg))

    table_weight = calculate_weight_from_rep_table(
        exercise_name=canonical_name,
        reps=reps,
        percentage_1rm=percentage_1rm,
        rep_table_dir=rep_table_dir,
    )

    rounded_weight = round_to_increment(table_weight, increment, mode="nearest")
    rep_table_percent = get_rep_table_percent_for_reps(
        exercise_name=canonical_name,
        reps=reps,
        rep_table_dir=rep_table_dir,
    )

    return format_weight(rounded_weight), rep_table_percent


def calculate_program(
    *,
    program: list[dict[str, Any]],
    training_maxes: dict[str, float],
    settings: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Walk the whole program structure and calculate set weights for any set that has
    'percentage_1rm'. Accessories without percentages are left untouched.

    Supported modes:
    - training_max (default)
    - rep_table (optional, explicit)

    Rep-table mode automatically falls back to training-max mode when the
    rep table does not support the requested reps, e.g. singles.
    """
    calculated_program = deepcopy(program)
    rounding_cfg = settings.get("rounding", {"default": 2.5})
    rep_table_dir = settings.get("rep_table_dir", "reference/rep_max_tables")

    for week in calculated_program:
        for _, day_list in week.items():
            if not isinstance(day_list, list):
                raise CalculationError("Each week value must be a list of training days")

            for day in day_list:
                exercises = day.get("exercises", [])
                if not isinstance(exercises, list):
                    raise CalculationError("Day 'exercises' must be a list")

                for exercise in exercises:
                    exercise_name = exercise.get("name")
                    if not isinstance(exercise_name, str) or not exercise_name.strip():
                        raise CalculationError(
                            f"Exercise must have a valid name: {exercise!r}"
                        )

                    raw_sets = exercise.get("sets", [])
                    if not isinstance(raw_sets, list):
                        raise CalculationError(
                            f"Exercise '{exercise_name}' has invalid 'sets': {raw_sets!r}"
                        )

                    expanded_sets = expand_sets(raw_sets)

                    for set_data in expanded_sets:
                        if "percentage_1rm" not in set_data:
                            continue

                        pct = validate_percentage(set_data["percentage_1rm"], exercise_name)
                        set_data["percentage_1rm"] = pct

                        mode = resolve_calculation_mode(
                            set_data=set_data,
                            exercise=exercise,
                            settings=settings,
                        )

                        reps_for_set = int(set_data["reps"])

                        if mode == "training_max":
                            set_data["weight"] = calculate_set_weight_training_max_mode(
                                exercise_name=exercise_name,
                                percentage_1rm=pct,
                                training_maxes=training_maxes,
                                rounding_cfg=rounding_cfg,
                            )

                        elif mode == "rep_table":
                            try:
                                weight, rep_table_percent = (
                                    calculate_set_weight_rep_table_mode(
                                        exercise_name=exercise_name,
                                        reps=reps_for_set,
                                        percentage_1rm=pct,
                                        rounding_cfg=rounding_cfg,
                                        rep_table_dir=rep_table_dir,
                                    )
                                )
                                set_data["weight"] = weight
                                set_data["rep_table_percent_for_reps"] = rep_table_percent
                            except RepTableError as exc:
                                # Fallback automatically when rep-table mode cannot handle
                                # the requested rep count or exercise format.
                                set_data["weight"] = calculate_set_weight_training_max_mode(
                                    exercise_name=exercise_name,
                                    percentage_1rm=pct,
                                    training_maxes=training_maxes,
                                    rounding_cfg=rounding_cfg,
                                )
                                set_data["rep_table_fallback"] = "training_max"
                                set_data["rep_table_error"] = str(exc)

                    exercise["sets"] = expanded_sets

    return calculated_program
