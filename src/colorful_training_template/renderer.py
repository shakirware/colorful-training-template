from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
import random

import yaml

from colorful_training_template.excel_generator.charts import add_charts_to_workbook
from colorful_training_template.excel_generator.companion_views import (
    add_companion_views_to_workbook,
)
from colorful_training_template.excel_generator.workout_template_generator import (
    WorkoutTemplateGenerator,
)
from colorful_training_template.utils.color_utils import generate_random_gradient


class RenderError(Exception):
    """Raised when rendering output fails."""


def write_yaml_output(data: list[dict[str, Any]], output_path: str | Path) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as f:
        yaml.safe_dump(
            data,
            f,
            sort_keys=False,
            allow_unicode=True,
            default_flow_style=False,
        )


def render_workbook(
    calculated_program: list[dict[str, Any]],
    settings: dict[str, Any],
) -> None:
    """
    Render the Excel workbook from calculated program data.

    Expected settings keys:
    - start_date
    - output_workbook
    - output_yaml
    """
    start_date_raw = settings.get("start_date")
    output_workbook = settings.get("output_workbook")
    output_yaml = settings.get("output_yaml")

    if not start_date_raw:
        raise RenderError("settings is missing 'start_date'")
    if not output_workbook:
        raise RenderError("settings is missing 'output_workbook'")
    if not output_yaml:
        raise RenderError("settings is missing 'output_yaml'")

    try:
        start_date = datetime.strptime(str(start_date_raw), "%Y-%m-%d")
    except ValueError as exc:
        raise RenderError(
            f"start_date must be in YYYY-MM-DD format, got {start_date_raw!r}"
        ) from exc

    output_workbook = Path(output_workbook)
    output_workbook.parent.mkdir(parents=True, exist_ok=True)

    generator = WorkoutTemplateGenerator(str(output_workbook), start_date)

    base_color = (random.randint(180, 240), 0.5, 0.7)
    gradient_colors = generate_random_gradient(
        base_color=base_color,
        lightness_variation=0.05,
        num_colors=6,
    )

    generator.create_consecutive_boxes(
        start_row=2,
        start_col=2,
        num_boxes=len(calculated_program),
        num_sets=7,
        num_exercises=8,
        space_between=3,
        set_width=4,
        headers=["Reps", "Weights", "%1RM", "Notes"],
        workout_data=calculated_program,
        fill_color=gradient_colors[0],
        week_fill_color=gradient_colors[1],
        set_fill_color=gradient_colors[2],
        day_fill_color=gradient_colors[3],
        exercise_fill_color=gradient_colors[4],
        grid_fill_color=gradient_colors[5],
    )

    try:
        add_charts_to_workbook(generator.workbook, yaml_path=output_yaml)
    except Exception as exc:
        raise RenderError(f"Failed to add charts: {exc}") from exc

    try:
        add_companion_views_to_workbook(
            generator.workbook,
            calculated_program=calculated_program,
            start_date=start_date,
        )
    except Exception as exc:
        raise RenderError(f"Failed to add week/today views: {exc}") from exc

    generator.save_workbook()
