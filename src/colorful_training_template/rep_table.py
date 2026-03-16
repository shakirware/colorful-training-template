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

    # Expected format:
    # row 0: ["Weight (kg)", 2, 3, 4, 5, ...]
    # row 1+: [10, "12.9kg", "15.9kg", ...]
    header_row = df.iloc[0].tolist()
    data_rows = df.iloc[1:].copy()

    # First column = base weight
    first_col_name = df.columns[0]
    data_rows = data_rows.rename(columns={first_col_name: "base_weight"})

    # Rename rep columns using the values from the first row
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

    # Keep only base_weight + rep columns
    valid_cols = ["base_weight"] + sorted(rename_map.values())
    data_rows = data_rows[valid_cols].copy()

    # Parse base weights
    data_rows["base_weight"] = data_rows["base_weight"].apply(_parse_kg)

    # Parse all rep-result cells
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
    """
    Compatibility helper for existing calculator code.

    Since your current tables are NOT percent tables, this returns the reps column
    requested as a simple float marker. It's kept only so existing metadata logic
    does not break.

    If you don't care about this metadata, you can later remove it from calculator.py.
    """
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
    percentage_1rm: float,
    rep_table_dir: str | Path,
) -> float:
    """
    Use the rep table directly.

    Logic:
    - percentage_1rm chooses a row by base weight
    - reps chooses a column
    - returned cell is the planned weight

    Example:
    - percentage_1rm = 82.5
    - nearest base_weight row = 82.5
    - reps = 3
    - result = value in column 3 for row 82.5
    """
    df = load_rep_table_matrix(exercise_name, rep_table_dir)

    if reps not in df.columns:
        supported = [col for col in df.columns if isinstance(col, int)]
        raise RepTableError(
            f"Reps={reps} not supported in rep table for '{exercise_name}'. "
            f"Supported reps: {supported}"
        )

    target_base_weight = float(percentage_1rm)

    # Nearest row match on base_weight
    idx = (df["base_weight"] - target_base_weight).abs().idxmin()
    return float(df.loc[idx, reps])
