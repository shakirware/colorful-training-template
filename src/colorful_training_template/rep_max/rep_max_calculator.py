import logging

import yaml

from colorful_training_template.rep_max.data_loader import (
    load_rep_max_excel,
    load_workout_data,
)

logger = logging.getLogger(__name__)


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


def round_to_nearest_1_25(value):
    return round(value * 4) / 4


def format_weight(weight):
    return f"{int(weight)}kg" if weight.is_integer() else f"{weight:.1f}kg"


def expand_sets(exercise):
    """
    Expands any set specification in an exercise where the "reps" field
    is provided in the format "NxM" (e.g. "3x8" means 3 sets of 8 reps).
    Each such set is replaced by N separate set dictionaries with reps=M.
    """
    expanded_sets = []
    for set_data in exercise.get("sets", []):
        reps_val = set_data.get("reps")
        if isinstance(reps_val, str) and "x" in reps_val:
            try:
                parts = reps_val.lower().split("x")
                num_sets = int(parts[0].strip())
                reps_per_set = int(parts[1].strip())
            except Exception as e:
                logger.error("Error parsing reps from '%s': %s", reps_val, e)
                continue
            logger.info(
                "Expanding set '%s' into %d sets of %d reps",
                reps_val,
                num_sets,
                reps_per_set,
            )
            for _ in range(num_sets):
                new_set = set_data.copy()
                new_set["reps"] = reps_per_set
                expanded_sets.append(new_set)
        else:
            expanded_sets.append(set_data)
    exercise["sets"] = expanded_sets


def process_exercise(exercise, rep_max_data, one_rep_max_values, exercise_aliases):
    exercise_name = exercise.get("name", "Unnamed Exercise")
    canonical_exercise = exercise_aliases.get(exercise_name, exercise_name)
    logger.info(
        "Processing exercise '%s' (mapped to '%s')", exercise_name, canonical_exercise
    )

    # Expand sets from "NxM" notation.
    expand_sets(exercise)

    # Attempt to get the rep max table; if none exists, we'll use the fallback.
    df = rep_max_data.get(canonical_exercise)

    # Get the one rep max value (you can set this to 50kg for weighted ring dips).
    if canonical_exercise in one_rep_max_values:
        max_1rm = one_rep_max_values[canonical_exercise]
    else:
        logger.warning(
            "No one rep max value available for exercise '%s'", exercise_name
        )
        return

    for set_data in exercise.get("sets", []):
        percentage = set_data.get("percentage_1rm")
        if percentage is None:
            continue
        try:
            target_1rm = max_1rm * (float(percentage) / 100)
        except Exception as e:
            logger.error("Error computing target 1RM for %s: %s", exercise_name, e)
            continue

        reps = set_data.get("reps")
        if reps is None:
            logger.warning(
                "No reps value found for a set in %s. Skipping set.", exercise_name
            )
            continue

        try:
            reps_int = int(reps)
        except Exception as e:
            logger.error(
                "Error converting reps '%s' to int for %s: %s", reps, exercise_name, e
            )
            continue

        # If a rep max table exists and contains the rep column, use it; otherwise, fallback.
        if df is not None and reps_int in df.columns:
            weight = find_weight_for_reps(df, target_1rm, reps_int)
            if weight is not None:
                set_data["weight"] = format_weight(weight)
                logger.debug("Set weight for %s: %s", exercise_name, set_data["weight"])
        else:
            fallback_weight = round_to_nearest_1_25(target_1rm)
            set_data["weight"] = format_weight(fallback_weight)
            logger.debug(
                "Fallback set weight for %s: %s", exercise_name, set_data["weight"]
            )


def update_workout_data_with_rep_max():
    """
    Loads the workout data, calculates weights based on rep max files,
    sorts each week’s days based on the 'weekday' key, updates the workout data in memory,
    optionally writes it to a file, and returns the updated workout data.
    """
    rep_max_files = {
        "Weighted Pull-Ups": "data/one_rep_max_data_pullups_90kg.xlsx",
        "Weighted Muscle-Ups": "data/one_rep_max_data_muscleups_90kg.xlsx",
        "Weighted Dips": "data/one_rep_max_data_dips_90kg.xlsx",
        "Squat": "data/one_rep_max_data_squat_90kg.xlsx",
        "Close Grip Bench Press": "data/one_rep_max_data_cgbp_90kg.xlsx",
    }
    one_rep_max_values = {
        "Weighted Pull-Ups": 100,
        "Weighted Muscle-Ups": 33.75,
        "Weighted Dips": 90,
        "Squat": 210,
        "Close Grip Bench Press": 125,
        "Weighted Ring Dips": 60,
    }
    exercise_aliases = {
        "Back Squats": "Squat",
        "Paused Squats": "Squat",
        "Dynamic Effort Squats": "Squat",
        "Weighted (Ring) Dips": "Weighted Dips",
    }

    rep_max_data = {}
    for exercise, file_path in rep_max_files.items():
        try:
            rep_max_data[exercise] = load_rep_max_excel(file_path)
            logger.info("Loaded rep max data for %s", exercise)
        except Exception as e:
            logger.error("Error loading rep max data for %s: %s", exercise, e)

    try:
        workout_data = load_workout_data("data/workout_data.yaml")
        logger.info("Loaded workout YAML data successfully.")
    except Exception as e:
        logger.error("Error loading workout YAML data: %s", e)
        return None

    if not isinstance(workout_data, list):
        logger.error("Workout data is not in the expected list format.")
        return None

    # Define weekday order mapping (Monday=1, Tuesday=2, ... Sunday=7)
    WEEKDAY_ORDER = {
        "Monday": 1,
        "Tuesday": 2,
        "Wednesday": 3,
        "Thursday": 4,
        "Friday": 5,
        "Saturday": 6,
        "Sunday": 7,
    }

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

            # Sort the list of days based on the weekday order.
            sorted_day_list = sorted(
                day_list, key=lambda d: WEEKDAY_ORDER.get(d.get("weekday", ""), 99)
            )
            week[week_name] = sorted_day_list
            logger.info("Sorted days for %s using weekday order.", week_name)

            for day_data in sorted_day_list:
                day_label = day_data.get("weekday", "Unknown")
                logger.info("Processing day %s", day_label)
                exercises = day_data.get("exercises")
                if not exercises:
                    logger.warning("No exercises found for day %s", day_label)
                    continue
                for exercise in exercises:
                    process_exercise(
                        exercise, rep_max_data, one_rep_max_values, exercise_aliases
                    )

    try:
        with open("data/calc_workout_data.yaml", "w") as file:
            yaml.dump(workout_data, file, default_flow_style=False)
        logger.info("Updated training program saved to 'data/calc_workout_data.yaml'")
    except Exception as e:
        logger.error("Error saving updated workout data: %s", e)

    return workout_data


if __name__ == "__main__":
    import logging

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    updated_data = update_workout_data_with_rep_max()
    if updated_data is not None:
        logger.info("Workout data updated successfully.")
