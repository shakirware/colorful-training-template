import logging
import random
from datetime import datetime

from colorful_training_template.excel_generator.charts import add_charts_to_workbook
from colorful_training_template.excel_generator.workout_template_generator import (
    WorkoutTemplateGenerator,
)
from colorful_training_template.rep_max.rep_max_calculator import (
    update_workout_data_with_rep_max,
)
from colorful_training_template.utils.color_utils import generate_random_gradient

logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def generate_workout_workbook(updated_workout_data):
    start_date = datetime.strptime("2025-02-17", "%Y-%m-%d")
    generator = WorkoutTemplateGenerator("workout_plan_new.xlsx", start_date)
    base_color = (random.randint(180, 240), 0.5, 0.7)
    gradient_colors = generate_random_gradient(
        base_color, lightness_variation=0.05, num_colors=6
    )
    logger.info("Using gradient colors: %s", gradient_colors)
    generator.create_consecutive_boxes(
        start_row=2,
        start_col=2,
        num_boxes=len(updated_workout_data),
        num_sets=7,
        num_exercises=8,
        space_between=3,
        set_width=4,
        headers=["Reps", "Weights", "%1RM", "Notes"],
        workout_data=updated_workout_data,
        fill_color=gradient_colors[0],
        week_fill_color=gradient_colors[1],
        set_fill_color=gradient_colors[2],
        day_fill_color=gradient_colors[3],
        exercise_fill_color=gradient_colors[4],
        grid_fill_color=gradient_colors[5],
    )
    # Add charts based on actual YAML data.
    add_charts_to_workbook(generator.workbook, yaml_path="data/workout_data.yaml")
    generator.save_workbook()


def main():
    updated_workout_data = update_workout_data_with_rep_max()
    if updated_workout_data is not None:
        generate_workout_workbook(updated_workout_data)
    else:
        logger.error("Failed to update workout data.")


if __name__ == "__main__":
    main()
