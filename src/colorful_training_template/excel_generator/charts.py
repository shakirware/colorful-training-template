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
    except (ValueError, TypeError):
        return (1, 0)


def get_chart_data_from_yaml(yaml_path):
    """
    Loads workout data from the YAML file and computes three metrics per week:
      - Stress Level: average %1RM across all sets
      - Volume: Sum over all sets of (sets × reps × weight)
      - Intensity: Same as stress level
    """
    with open(yaml_path, "r") as file:
        data = yaml.safe_load(file)

    # This logic correctly finds the list of weeks from your YAML
    weeks_list = []
    if isinstance(data, dict):
        weeks_list = data.get("custom_program", [])
    elif isinstance(data, list):
        weeks_list = data

    week_numbers = []
    stress_levels = []
    volumes = []
    intensities = []

    # Iterate over each week in the program
    for week_index, week_data in enumerate(weeks_list):
        week_name = list(week_data.keys())[0]  # e.g., "week_1"
        days = week_data[week_name]

        week_numbers.append(week_index + 1)

        set_percentages = []
        total_volume = 0

        # Loop through each day in the week
        for day in days:
            for exercise in day.get("exercises", []):
                # We need to parse each set entry individually now
                for set_entry in exercise.get("sets", []):
                    # Process percentage for stress/intensity
                    perc = set_entry.get("percentage_1rm")
                    if perc is not None:
                        try:
                            set_percentages.append(float(perc))
                        except (ValueError, TypeError):
                            pass

                    # Process volume data
                    reps_val = set_entry.get("reps")
                    weight_val = set_entry.get("weight")

                    if reps_val is not None and weight_val is not None:
                        try:
                            # Use your existing parse_reps for the "NxM" format
                            count, reps_per_set = parse_reps(reps_val)

                            weight_str = str(weight_val).replace("kg", "").strip()
                            weight_num = float(weight_str)

                            total_volume += count * reps_per_set * weight_num
                        except (ValueError, TypeError, IndexError) as e:
                            logger.warning(
                                "Could not parse volume for set (%s, %s): %s",
                                reps_val,
                                weight_val,
                                e,
                            )

        # Compute averages for the week
        avg_stress = (
            sum(set_percentages) / len(set_percentages) if set_percentages else 0
        )
        stress_levels.append(avg_stress)
        intensities.append(avg_stress)
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
