import pandas as pd
import yaml


def load_rep_max_excel(file_path):
    try:
        df = pd.read_excel(file_path, skiprows=1)
    except Exception as e:
        raise ValueError(f"Error reading Excel file '{file_path}': {e}")
    for column in df.columns:
        if column != "Weight (kg)":
            df[column] = df[column].apply(
                lambda x: float(str(x).replace("kg", "").strip()) if pd.notna(x) else x
            )
    try:
        df["Weight (kg)"] = df["Weight (kg)"].astype(float)
    except Exception as e:
        raise ValueError(
            f"Error converting 'Weight (kg)' column to float in '{file_path}': {e}"
        )
    return df


def load_workout_data(file_path):
    with open(file_path, "r") as file:
        data = yaml.safe_load(file)
    # If data is a dict with a "weeks" key, return that list.
    if isinstance(data, dict):
        if "weeks" in data:
            return data["weeks"]
        else:
            raise ValueError(
                "YAML data is a dictionary but is missing the 'weeks' key."
            )
    elif isinstance(data, list):
        return data
    else:
        raise ValueError(
            "YAML data is not in the expected format (dict with 'weeks' key or list)."
        )
