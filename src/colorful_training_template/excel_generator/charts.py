import logging

import yaml
from openpyxl.chart import LineChart, Reference

logger = logging.getLogger(__name__)


def parse_reps(reps_value):
    """
    Parse the reps value.
    If it's a string in the form "NxM" (e.g., "5x8"), return (N, M).
    Otherwise, if it's numeric, assume (1, reps_value).
    """
    if isinstance(reps_value, str) and "x" in reps_value.lower():
        try:
            parts = reps_value.lower().split("x")
            count = int(parts[0].strip())
            reps_per_set = int(parts[1].strip())
            return count, reps_per_set
        except Exception as e:
            logger.error("Error parsing reps from '%s': %s", reps_value, e)
            return (1, 0)
    try:
        return (1, int(reps_value))
    except Exception:
        return (1, 0)


def get_chart_data_from_yaml(yaml_path):
    """
    Loads workout data from the YAML file and computes three metrics per week:
      - Stress Level: average %1RM across all sets (ignoring sets with null percentage)
      - Volume: Sum over all sets of (set_count × reps_per_set × weight)
      - Intensity: average %1RM (using the same value as stress here)

    Returns:
      (weeks, stress_levels, volumes, intensities)
    """
    with open(yaml_path, "r") as file:
        data = yaml.safe_load(file)
    # Handle both cases: data is a dict with a "weeks" key or data is a list
    if isinstance(data, dict):
        weeks_list = data.get("weeks", [])
    elif isinstance(data, list):
        weeks_list = data
    else:
        weeks_list = []

    week_numbers = []
    stress_levels = []
    volumes = []
    intensities = []
    week_index = 1

    # Iterate over each week.
    for week_dict in weeks_list:
        # week_dict is expected to be a dict with one key (e.g. "week_1")
        for week_name, days in week_dict.items():
            week_numbers.append(week_index)
            week_index += 1
            set_percentages = []
            total_volume = 0
            # Loop through each day in the week.
            for day in days:
                exercises = day.get("exercises", [])
                for exercise in exercises:
                    for set_entry in exercise.get("sets", []):
                        # Process percentage if available.
                        perc = set_entry.get("percentage_1rm")
                        if perc is not None:
                            try:
                                perc_val = float(perc)
                                set_percentages.append(perc_val)
                            except Exception:
                                pass
                        # Process volume data.
                        reps_val = set_entry.get("reps")
                        weight_val = set_entry.get("weight")
                        if reps_val is not None and weight_val is not None:
                            try:
                                # Extract numeric portion from weight (e.g., "47.5kg").
                                weight_num = float(
                                    "".join(
                                        c for c in weight_val if c.isdigit() or c == "."
                                    )
                                )
                            except Exception:
                                weight_num = 0
                            count, reps_per_set = parse_reps(reps_val)
                            total_volume += count * reps_per_set * weight_num
            # Compute averages. If no set percentages are available, use 0.
            if set_percentages:
                avg_stress = sum(set_percentages) / len(set_percentages)
            else:
                avg_stress = 0
            stress_levels.append(avg_stress)
            intensities.append(avg_stress)  # Can be adjusted for a different measure.
            volumes.append(total_volume)

    return week_numbers, stress_levels, volumes, intensities


def add_charts_to_workbook(wb, yaml_path="data/workout_data.yaml"):
    """
    Reads actual workout data from the YAML file, computes weekly metrics,
    and adds three charts to the workbook:
      1. Stress vs. Recovery (Sawtooth Pattern)
      2. Volume Over Time
      3. Intensity Progression (% of 1RM Over Time)
    """
    weeks, stress_levels, volumes, intensities = get_chart_data_from_yaml(yaml_path)

    if not weeks:
        logger.error("No week data found. Charts cannot be generated.")
        return

    # Create a new worksheet for chart data and charts.
    chart_sheet = wb.create_sheet("Charts")

    # Write header row.
    chart_sheet.append(["Week", "Stress Level", "Volume", "Intensity (% of 1RM)"])
    # Write weekly data.
    for w, stress, vol, inten in zip(weeks, stress_levels, volumes, intensities):
        chart_sheet.append([w, stress, vol, inten])

    # Create category reference (Weeks).
    cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=len(weeks) + 1)

    # Chart 1: Stress vs. Recovery
    stress_chart = LineChart()
    stress_chart.title = "Stress vs. Recovery"
    stress_chart.y_axis.title = "Stress Level"
    stress_chart.x_axis.title = "Week"
    data_stress = Reference(chart_sheet, min_col=2, min_row=1, max_row=len(weeks) + 1)
    stress_chart.add_data(data_stress, titles_from_data=True)
    stress_chart.set_categories(cats)
    chart_sheet.add_chart(stress_chart, "F2")

    # Chart 2: Volume Over Time
    volume_chart = LineChart()
    volume_chart.title = "Volume Over Time"
    volume_chart.y_axis.title = "Volume (Sets × Reps × Weight)"
    volume_chart.x_axis.title = "Week"
    data_volume = Reference(chart_sheet, min_col=3, min_row=1, max_row=len(weeks) + 1)
    volume_chart.add_data(data_volume, titles_from_data=True)
    volume_chart.set_categories(cats)
    chart_sheet.add_chart(volume_chart, "F20")

    # Chart 3: Intensity Progression (% of 1RM Over Time)
    intensity_chart = LineChart()
    intensity_chart.title = "Intensity Progression (% of 1RM Over Time)"
    intensity_chart.y_axis.title = "% of 1RM"
    intensity_chart.x_axis.title = "Week"
    data_intensity = Reference(
        chart_sheet, min_col=4, min_row=1, max_row=len(weeks) + 1
    )
    intensity_chart.add_data(data_intensity, titles_from_data=True)
    intensity_chart.set_categories(cats)
    chart_sheet.add_chart(intensity_chart, "F38")
