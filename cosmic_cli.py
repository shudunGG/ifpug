"""Command line interface for COSMIC Function Point measurement."""
from __future__ import annotations

import argparse
from pathlib import Path

from cosmic import ExcelExporter, load_measurement


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="COSMIC Function Point calculator")
    parser.add_argument("config", type=Path, help="Path to the measurement definition (YAML or JSON)")
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("cosmic_measurement.xlsx"),
        help="Path for the generated Excel workbook",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    measurement = load_measurement(args.config)
    exporter = ExcelExporter(measurement)
    output_path = exporter.export(args.output)
    print(f"Excel report generated at: {output_path.resolve()}")


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main()
