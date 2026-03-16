from colorful_training_template.config import (
    load_program,
    load_settings,
    load_training_maxes,
)
from colorful_training_template.calculator import calculate_program
from colorful_training_template.renderer import (
    render_workbook,
    write_yaml_output,
)


def build() -> int:
    """
    Main build path:
    1. Load inputs
    2. Calculate weights from training maxes and percentages
    3. Write calculated YAML
    4. Render workbook
    """
    training_maxes = load_training_maxes()
    settings = load_settings()
    program = load_program()

    calculated_program = calculate_program(
        program=program,
        training_maxes=training_maxes,
        settings=settings,
    )

    output_yaml = settings["output_yaml"]
    output_workbook = settings["output_workbook"]

    write_yaml_output(calculated_program, output_yaml)
    render_workbook(calculated_program, settings)

    print("Built workout plan successfully:")
    print(f"- {output_yaml}")
    print(f"- {output_workbook}")
    return 0


def main() -> int:
    return build()


if __name__ == "__main__":
    raise SystemExit(main())
