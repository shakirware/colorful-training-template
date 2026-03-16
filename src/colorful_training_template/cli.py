import argparse
import sys

from colorful_training_template.main import build
from colorful_training_template.config import (
    load_program,
    load_settings,
    load_training_maxes,
)


def validate() -> int:
    """
    Validate that the required config files can be loaded.
    Keep this intentionally simple for now.
    """
    try:
        load_training_maxes()
        load_settings()
        load_program()
    except Exception as exc:
        print(f"Validation failed: {exc}", file=sys.stderr)
        return 1

    print("Validation OK")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="training-plan",
        description="Build and validate a training plan from YAML config files.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser(
        "build",
        help="Build calculated workout data and the Excel workbook.",
    )
    subparsers.add_parser(
        "validate",
        help="Validate config files only.",
    )

    args = parser.parse_args()

    if args.command == "build":
        return build()

    if args.command == "validate":
        return validate()

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
