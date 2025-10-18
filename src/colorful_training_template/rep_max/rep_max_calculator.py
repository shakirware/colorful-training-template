import logging

import yaml

from colorful_training_template.rep_max.data_loader import (
    load_rep_max_excel,
    load_workout_data,
)

logger = logging.getLogger(__name__)

# ========== CONFIG ==========
# If False (default), weights are computed as: weight = TM * (%/100), rounded down.
# If True, will try rep-max table lookup first (not recommended for your block),
# then fall back to % of TM if needed.
USE_REP_MAX_TABLES = False

# Use Training Max anchors (True) vs raw table 1RMs (False)
USE_TRAINING_MAX = True

# Round down to this increment (microplates). 1.25kg = 0.25*5.
ROUNDING_INCREMENT_KG = 1.25
ROUNDING_MODE = "down"  # "down" | "nearest"

# Preferred: directamente set Training Maxes per lift (these are THE anchors)
TRAINING_MAX_OVERRIDE = {
    "Weighted Pull-Ups": 82.5,  # added load
    "Weighted Dips": 92.5,  # added load
    "Weighted Muscle-Ups": 28.0,  # added load
    "Squat": 195.0,  # barbell load
    "Close Grip Bench Press": 115.0,  # barbell load
    "Weighted Ring Dips": 57.5,  # added load (if used)
}

# If you don’t override, we can derive TM from your table 1RM via factors.
TM_FACTORS_DEFAULT = {
    "Weighted Pull-Ups": 0.88,  # ~75 from 85
    "Weighted Dips": 0.944,  # ~85 from 90
    "Weighted Muscle-Ups": 1.00,
    "Squat": 1.00,
    "Close Grip Bench Press": 0.92,  # ~115 from 125
    "Weighted Ring Dips": 0.944,
}
# ===========================


def find_weight_for_reps(df, target_1rm, reps):
    try:
        reps = int(reps)
    except ValueError:
        logger.error("Invalid reps value: %s. Skipping calculation.", reps)
        return None
    if reps not in df.columns:
        logger.warning(
            "Reps value '%s' not found in DataFrame columns: %s",
            reps,
            df.columns.tolist(),
        )
        return None
    closest_row_index = (df[reps] - target_1rm).abs().idxmin()
    return df.loc[closest_row_index, "Weight (kg)"]


def round_to_increment(value: float, inc: float, mode: str = "down") -> float:
    if inc <= 0:
        return value
    if mode == "down":
        # always bias a touch conservative
        steps = int(value // inc)
        return steps * inc
    # default to nearest
    return round(value / inc) * inc


def format_weight(weight):
    w = float(weight)
    return f"{int(w)}kg" if w.is_integer() else f"{w:.1f}kg"


def expand_sets(exercise):
    expanded_sets = []
    for set_data in exercise.get("sets", []):
        reps_val = set_data.get("reps")
        if isinstance(reps_val, str) and "x" in reps_val.lower():
            try:
                n, r = reps_val.lower().split("x")
                num_sets = int(n.strip())
                reps_per_set = int(r.strip())
                for _ in range(num_sets):
                    new_set = set_data.copy()
                    new_set["reps"] = reps_per_set
                    expanded_sets.append(new_set)
            except Exception as e:
                logger.error("Error parsing reps from '%s': %s", reps_val, e)
                expanded_sets.append(set_data)
        else:
            expanded_sets.append(set_data)
    exercise["sets"] = expanded_sets


def compute_effective_maxes(
    one_rep_max_values, use_training_max=True, tm_factors=None, tm_override=None
):
    tm_factors = tm_factors or TM_FACTORS_DEFAULT
    tm_override = tm_override or {}
    effective = {}
    for lift, table_1rm in one_rep_max_values.items():
        if use_training_max and lift in tm_override:
            effective[lift] = float(tm_override[lift])
            src = "OVERRIDE"
        elif use_training_max:
            factor = tm_factors.get(lift, 0.90)
            effective[lift] = table_1rm * factor
            src = f"FACTOR({factor:.3f})"
        else:
            effective[lift] = float(table_1rm)
            src = "RAW_1RM"
        logger.info("Anchor for %-25s = %7.2f kg  [%s]", lift, effective[lift], src)
    return effective


def process_exercise(exercise, rep_max_data, effective_max_values, exercise_aliases):
    exercise_name = exercise.get("name", "Unnamed Exercise")
    canonical_exercise = exercise_aliases.get(exercise_name, exercise_name)
    logger.info(
        "Processing exercise '%s' (mapped to '%s')", exercise_name, canonical_exercise
    )

    original_sets = exercise.get("sets", [])
    expanded_sets = []

    # expand NxM
    for set_data in original_sets:
        reps_val = set_data.get("reps")
        if isinstance(reps_val, str) and "x" in reps_val.lower():
            try:
                n, r = reps_val.lower().split("x")
                for _ in range(int(n.strip())):
                    new_set = set_data.copy()
                    new_set["reps"] = int(r.strip())
                    expanded_sets.append(new_set)
            except Exception as e:
                logger.error("Error parsing reps from '%s': %s", reps_val, e)
                expanded_sets.append(set_data)
        else:
            expanded_sets.append(set_data)

    # calculate
    df = rep_max_data.get(canonical_exercise)
    max_anchor = effective_max_values.get(canonical_exercise)

    if not max_anchor:
        logger.warning("No max value for exercise '%s'", canonical_exercise)
        exercise["sets"] = expanded_sets
        return

    for set_data in expanded_sets:
        pct = set_data.get("percentage_1rm")
        if pct is None:
            continue
        try:
            reps_int = int(set_data["reps"])
        except Exception:
            reps_int = None

        target = max_anchor * (float(pct) / 100.0)

        weight = None
        if (
            USE_REP_MAX_TABLES
            and df is not None
            and reps_int is not None
            and reps_int in df.columns
        ):
            # Table-based suggestion (optional path)
            weight = find_weight_for_reps(df, target, reps_int)

        if weight is None:
            # Pure % of TM, rounded down (recommended path)
            weight = round_to_increment(
                target, ROUNDING_INCREMENT_KG, mode=ROUNDING_MODE
            )

        set_data["weight"] = format_weight(weight)

    exercise["sets"] = expanded_sets


def update_workout_data_with_rep_max(
    use_training_max: bool = USE_TRAINING_MAX,
    tm_factors: dict | None = None,
    rounding_increment: float = ROUNDING_INCREMENT_KG,
    use_rep_max_tables: bool = USE_REP_MAX_TABLES,
    tm_override: dict | None = TRAINING_MAX_OVERRIDE,
):
    """
    Calculate set weights from % of Training Max (recommended) or optionally via rep-max tables.
    Writes 'data/calc_workout_data.yaml'.
    """
    global ROUNDING_INCREMENT_KG, USE_REP_MAX_TABLES
    ROUNDING_INCREMENT_KG = rounding_increment
    USE_REP_MAX_TABLES = use_rep_max_tables

    rep_max_files = {
        "Weighted Pull-Ups": "data/one_rep_max_data_pullups_90kg.xlsx",
        "Weighted Muscle-Ups": "data/one_rep_max_data_muscleups_90kg.xlsx",
        "Weighted Dips": "data/one_rep_max_data_dips_90kg.xlsx",
        "Squat": "data/one_rep_max_data_squat_90kg.xlsx",
        "Close Grip Bench Press": "data/one_rep_max_data_cgbp_90kg.xlsx",
    }
    # NOTE: Added load for WP/Dips/MU; barbell loads for Squat/Bench.
    one_rep_max_values = {
        "Weighted Pull-Ups": 85,
        "Weighted Muscle-Ups": 28,
        "Weighted Dips": 90,
        "Squat": 195,
        "Close Grip Bench Press": 125,
        "Weighted Ring Dips": 60,
    }
    exercise_aliases = {
        "Back Squats": "Squat",
        "Paused Squats": "Squat",
        "Dynamic Effort Squats": "Squat",
        "Weighted (Ring) Dips": "Weighted Dips",
    }

    # Load tables only if we plan to use them
    rep_max_data = {}
    if use_rep_max_tables:
        for exercise, file_path in rep_max_files.items():
            try:
                rep_max_data[exercise] = load_rep_max_excel(file_path)
                logger.info("Loaded rep max data for %s", exercise)
            except Exception as e:
                logger.error("Error loading rep max data for %s: %s", exercise, e)

    # Load program YAML
    try:
        workout_data = load_workout_data("data/workout_data.yaml")
        logger.info("Loaded workout YAML data successfully.")
    except Exception as e:
        logger.error("Error loading workout YAML data: %s", e)
        return None

    if not isinstance(workout_data, list):
        logger.error("Workout data is not in the expected list format.")
        return None

    # Compute anchors (TM or raw 1RM, with optional override)
    effective_max_values = compute_effective_maxes(
        one_rep_max_values,
        use_training_max=use_training_max,
        tm_factors=tm_factors,
        tm_override=tm_override,
    )

    WEEKDAY_ORDER = {
        "Monday": 1,
        "Tuesday": 2,
        "Wednesday": 3,
        "Thursday": 4,
        "Friday": 5,
        "Saturday": 6,
        "Sunday": 7,
    }

    # Process each week/day
    for week in workout_data:
        if not isinstance(week, dict):
            logger.warning(
                "Each week should be a dict mapping week names to days. Skipping invalid week: %s",
                week,
            )
            continue
        for week_name, day_list in week.items():
            logger.info("Processing %s", week_name)
            if not isinstance(day_list, list):
                logger.warning("Days for %s should be a list. Skipping.", week_name)
                continue

            sorted_day_list = sorted(
                day_list, key=lambda d: WEEKDAY_ORDER.get(d.get("weekday", ""), 99)
            )
            week[week_name] = sorted_day_list

            for day_data in sorted_day_list:
                day_label = day_data.get("weekday", "Unknown")
                logger.info("Processing day %s", day_label)
                exercises = day_data.get("exercises")
                if not exercises:
                    logger.warning("No exercises found for day %s", day_label)
                    continue
                for exercise in exercises:
                    process_exercise(
                        exercise, rep_max_data, effective_max_values, exercise_aliases
                    )

    # Save
    try:
        with open("data/calc_workout_data.yaml", "w") as file:
            yaml.dump(workout_data, file, default_flow_style=False)
        logger.info("Updated training program saved to 'data/calc_workout_data.yaml'")
    except Exception as e:
        logger.error("Error saving updated workout data: %s", e)

    return workout_data


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    updated_data = update_workout_data_with_rep_max(
        use_training_max=USE_TRAINING_MAX,
        tm_factors=TM_FACTORS_DEFAULT,
        rounding_increment=ROUNDING_INCREMENT_KG,
        use_rep_max_tables=USE_REP_MAX_TABLES,
        tm_override=TRAINING_MAX_OVERRIDE,
    )
    if updated_data is not None:
        logger.info("Workout data updated successfully.")
