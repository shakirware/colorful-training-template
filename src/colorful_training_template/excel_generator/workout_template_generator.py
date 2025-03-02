import logging
from datetime import timedelta

import openpyxl

from colorful_training_template.excel_generator.box_creator import BoxCreator
from colorful_training_template.excel_generator.data_populator import DataPopulator
from colorful_training_template.excel_generator.label_creator import LabelCreator

logger = logging.getLogger(__name__)


class WorkoutTemplateGenerator:
    def __init__(self, workbook_name, start_date):
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.workbook_name = workbook_name
        self.start_date = start_date
        self.box_creator = BoxCreator(self.sheet)
        self.label_creator = LabelCreator(self.sheet)
        self.data_populator = DataPopulator(self.sheet)
        logger.info(
            "Initialized WorkoutTemplateGenerator with workbook '%s' and start_date %s",
            workbook_name,
            start_date,
        )

    def create_consecutive_boxes(
        self,
        start_row,
        start_col,
        num_boxes,
        num_sets,
        num_exercises,
        space_between,
        fill_color,
        week_fill_color,
        set_fill_color,
        day_fill_color,
        exercise_fill_color,
        grid_fill_color,
        set_width,
        headers,
        workout_data,
    ):
        logger.info(
            "Creating %d consecutive boxes with %d sets and %d exercises each.",
            num_boxes,
            num_sets,
            num_exercises,
        )
        box_height, box_width = self.calculate_box_dimensions(
            set_width, num_sets, num_exercises
        )
        logger.debug(
            "Calculated box dimensions: height=%d, width=%d", box_height, box_width
        )
        for i in range(num_boxes):
            current_start_col = start_col + i * (box_width + space_between)
            logger.info(
                "Creating box %d at row %d, col %d", i + 1, start_row, current_start_col
            )
            # Create the outer box with a fill color
            self.box_creator.create_box(
                start_row,
                current_start_col,
                start_row + box_height - 1,
                current_start_col + box_width - 1,
                fill_color,
            )
            current_week_start_date = self.start_date + timedelta(weeks=i)
            logger.info(
                "Labeling week %d (commencing %s)",
                i + 1,
                current_week_start_date.strftime("%Y-%m-%d"),
            )
            # Label the week at the top of the box
            self.label_creator.label_weeks(
                start_row,
                current_start_col,
                box_width,
                i + 1,
                current_week_start_date,
                week_fill_color,
            )
            # Label the sets (starting a few rows down)
            set_start_row = start_row + 3
            logger.info("Labeling sets starting at row %d", set_start_row)
            self.label_creator.label_sets(
                set_start_row,
                current_start_col + 1,
                num_sets,
                set_fill_color,
                set_width,
                headers,
            )
            # Label days and fill in the exercise grid
            day_start_row = set_start_row + 4
            logger.info("Labeling days starting at row %d", day_start_row)
            self.label_creator.label_days(
                day_start_row,
                current_start_col,
                num_exercises,
                day_fill_color,
                exercise_fill_color,
                num_sets,
                set_width,
                grid_fill_color,
            )
            logger.info("Populating workout data for week %d", i + 1)
            self.data_populator.populate_workout_data(
                day_start_row,
                current_start_col,
                workout_data[i][f"week_{i+1}"],
                num_sets,
                set_width,
                num_exercises,
            )

    def calculate_box_dimensions(self, set_width, num_sets, num_exercises):
        # Header: 2 (top) + 1 (week) + 2 (set titles) + 1 (set headers) = 6 rows
        header_height = 6
        # Each day block gets num_exercises rows plus 1 spacer row
        day_block_height = num_exercises + 1
        min_height = header_height + 7 * day_block_height + 1
        # Width calculation (adjustable as needed)
        min_width = 1 + 3 + num_sets * (set_width + 1) - 1 + 2 + 1
        logger.debug(
            "Box dimensions calculated: header_height=%d, day_block_height=%d, min_height=%d, min_width=%d",
            header_height,
            day_block_height,
            min_height,
            min_width,
        )
        return min_height, min_width

    def save_workbook(self):
        self.workbook.save(self.workbook_name)
        logger.info("Workbook saved as %s", self.workbook_name)
