import unittest
from pathlib import Path

def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent


class MeasurementWorkflowTest(unittest.TestCase):
    def test_measurement_summary_and_excel(self) -> None:
        from cosmic import CosmicCalculator, ExcelExporter, load_measurement

        measurement = load_measurement(_project_root() / "example_measurement.yaml")
        calculator = CosmicCalculator(measurement)

        summaries = calculator.summarize()
        submit_order = summaries["Submit Order"]
        cancel_order = summaries["Cancel Order"]

        self.assertEqual(submit_order.entry_count, 1)
        self.assertEqual(submit_order.exit_count, 1)
        self.assertEqual(submit_order.read_count, 1)
        self.assertEqual(submit_order.write_count, 1)
        self.assertEqual(submit_order.total_cfp, 4)

        self.assertEqual(cancel_order.total_cfp, 4)
        self.assertEqual(calculator.total_cfp(), 8)

        exporter = ExcelExporter(measurement)
        output_file = exporter.export(_project_root() / "tests" / "report.xlsx")
        self.assertTrue(output_file.exists())


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
