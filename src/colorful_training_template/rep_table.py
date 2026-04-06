from __future__ import annotations

from pathlib import Path

import pandas as pd


class RepTableError(Exception):
    """Raised when rep table loading/calculation fails."""


REP_TABLE_FILES = {
    "Weighted Pull-Ups": "one_rep_max_data_pullups_90kg.xlsx",
    "Weighted Dips": "one_rep_max_data_dips_90kg.xlsx",
    "Weighted Muscle-Ups": "one_rep_max_data_muscleups_90kg.xlsx",
    "Squat": "one_rep_max_data_squat_90kg.xlsx",
    "Close Grip Bench Press": "one_rep_max_data_cgbp_90kg.xlsx",
}


def get_rep_table_path(exercise_name: str, rep_table_dir: str | Path) -> Path:
    if exercise_name not in REP_TABLE_FILES:
        raise RepTableError(f"No rep table file mapping found for exercise: {exercise_name}")
    return Path(rep_table_dir) / REP_TABLE_FILES[exercise_name]


def _parse_kg(value) -> float:
    if pd.isna(value):
        raise RepTableError("Encountered empty value in rep table")

    text = str(value).strip().lower().replace("kg", "").strip()
    try:
        return float(text)
    except ValueError as exc:
        raise RepTableError(f"Could not parse kg value from {value!r}") from exc


def load_rep_table_matrix(exercise_name: str, rep_table_dir: str | Path) -> pd.DataFrame:
    path = get_rep_table_path(exercise_name, rep_table_dir)

    if not path.exists():
        raise RepTableError(f"Rep table file not found: {path}")

    df = pd.read_excel(path, sheet_name=0)

    if df.empty or len(df) < 2:
        raise RepTableError(f"Rep table file is empty or malformed: {path}")

    header_row = df.iloc[0].tolist()
    data_rows = df.iloc[1:].copy()

    first_col_name = df.columns[0]
    data_rows = data_rows.rename(columns={first_col_name: "base_weight"})

    rename_map = {}
    for i, col_name in enumerate(df.columns[1:], start=1):
        header_val = header_row[i]
        try:
            rep_num = int(float(str(header_val).strip()))
        except Exception as exc:
            raise RepTableError(
                f"Could not parse rep header from first row value {header_val!r}"
            ) from exc
        rename_map[col_name] = rep_num

    data_rows = data_rows.rename(columns=rename_map)

    valid_cols = ["base_weight"] + sorted(rename_map.values())
    data_rows = data_rows[valid_cols].copy()

    data_rows["base_weight"] = data_rows["base_weight"].apply(_parse_kg)

    for rep_col in sorted(rename_map.values()):
        data_rows[rep_col] = data_rows[rep_col].apply(_parse_kg)

    data_rows = data_rows.sort_values("base_weight").reset_index(drop=True)
    return data_rows


def get_supported_reps(exercise_name: str, rep_table_dir: str | Path) -> list[int]:
    df = load_rep_table_matrix(exercise_name, rep_table_dir)
    return [col for col in df.columns if isinstance(col, int)]


def get_rep_table_percent_for_reps(
    exercise_name: str,
    reps: int,
    rep_table_dir: str | Path,
) -> float:
    supported = get_supported_reps(exercise_name, rep_table_dir)
    if reps not in supported:
        raise RepTableError(
            f"Reps={reps} not supported in rep table for '{exercise_name}'. "
            f"Supported reps: {supported}"
        )
    return float(reps)


def calculate_weight_from_rep_table(
    exercise_name: str,
    reps: int,
    target_weight: float,
    rep_table_dir: str | Path,
) -> float:
    """
    The rep tables are estimated-1RM tables:
    - base_weight row = lifted weight
    - reps column = reps performed
    - cell value = estimated 1RM

    So to get the planned working weight for a target estimated 1RM:
    - find the requested reps column
    - find the row whose estimated-1RM cell is closest to target_weight
    - return that row's base_weight
    """
    df = load_rep_table_matrix(exercise_name, rep_table_dir)

    if reps not in df.columns:
        supported = [col for col in df.columns if isinstance(col, int)]
        raise RepTableError(
            f"Reps={reps} not supported in rep table for '{exercise_name}'. "
            f"Supported reps: {supported}"
        )

    target_estimated_1rm = float(target_weight)

    idx = (df[reps] - target_estimated_1rm).abs().idxmin()
    return float(df.loc[idx, "base_weight"])
