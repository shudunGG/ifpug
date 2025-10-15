"""Excel export functionality without third-party dependencies."""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

from .calculator import CosmicCalculator, FunctionalProcessSummary
from .models import SystemMeasurement


@dataclass
class _Sheet:
    name: str
    rows: List[List[object]]


def _column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _cell_xml(value: object, row_index: int, column_index: int) -> str:
    cell_ref = f"{_column_letter(column_index)}{row_index}"
    if value is None:
        return f"<c r=\"{cell_ref}\"/>"
    if isinstance(value, (int, float)):
        return f"<c r=\"{cell_ref}\"><v>{value}</v></c>"
    if isinstance(value, _dt.date):
        # Excel stores dates as numbers with the epoch 1899-12-30
        excel_epoch = _dt.date(1899, 12, 30)
        delta = value - excel_epoch
        return f"<c r=\"{cell_ref}\" s=\"1\"><v>{delta.days}</v></c>"
    text = escape(str(value))
    return (
        f"<c r=\"{cell_ref}\" t=\"inlineStr\">"
        f"<is><t xml:space=\"preserve\">{text}</t></is>"
        f"</c>"
    )


def _sheet_xml(sheet: _Sheet) -> str:
    rows_xml = []
    for row_idx, row in enumerate(sheet.rows, start=1):
        cells_xml = "".join(_cell_xml(value, row_idx, col_idx) for col_idx, value in enumerate(row, start=1))
        rows_xml.append(f"<row r=\"{row_idx}\">{cells_xml}</row>")
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetData>"
        + "".join(rows_xml)
        + "</sheetData>"
        "</worksheet>"
    )


def _workbook_xml(sheets: Iterable[_Sheet]) -> str:
    sheet_entries = []
    for index, sheet in enumerate(sheets, start=1):
        sheet_entries.append(
            f"<sheet name=\"{escape(sheet.name)}\" sheetId=\"{index}\" r:id=\"rId{index}\"/>"
        )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets>"
        + "".join(sheet_entries)
        + "</sheets>"
        "</workbook>"
    )


def _workbook_rels_xml(sheet_count: int) -> str:
    relationships = []
    for index in range(1, sheet_count + 1):
        relationships.append(
            "<Relationship id=\"rId{0}\" type=\"http://schemas.openxmlformats.org/officeDocument/"
            "2006/relationships/worksheet\" target=\"worksheets/sheet{0}.xml\"/>".format(index)
        )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "".join(relationships)
        + "</Relationships>"
    )


def _content_types_xml(sheet_count: int) -> str:
    overrides = []
    for index in range(1, sheet_count + 1):
        overrides.append(
            "<Override PartName=\"/xl/worksheets/sheet{0}.xml\" ContentType=\"application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.worksheet+xml\"/>".format(index)
        )
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument."
        "spreadsheetml.styles+xml\"/>"
        + "".join(overrides)
        + "</Types>"
    )


def _styles_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font></fonts>"
        "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
        "<borders count=\"1\"><border/></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"2\">"
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
        "</cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "</styleSheet>"
    )


def _rels_xml() -> str:
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/"
        "2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )


def _build_functional_process_rows(summaries: Iterable[FunctionalProcessSummary]) -> List[List[object]]:
    header = [
        "Functional Process",
        "Trigger",
        "Object of Interest",
        "Description",
        "Entry (E)",
        "Exit (X)",
        "Read (R)",
        "Write (W)",
        "Total CFP",
    ]
    rows = [header]
    for summary in summaries:
        rows.append(
            [
                summary.name,
                summary.trigger,
                summary.object_of_interest,
                summary.description,
                summary.entry_count,
                summary.exit_count,
                summary.read_count,
                summary.write_count,
                summary.total_cfp,
            ]
        )
    return rows


def _build_data_movement_rows(measurement: SystemMeasurement) -> List[List[object]]:
    header = [
        "Functional Process",
        "Sequence",
        "Movement Type",
        "Description",
        "Object of Interest",
        "Trigger",
        "Code Reference",
        "Notes",
    ]
    rows = [header]
    for process in measurement.functional_processes:
        for index, movement in enumerate(process.data_movements, start=1):
            rows.append(
                [
                    process.name,
                    index,
                    movement.movement_type.value,
                    movement.description,
                    movement.object_of_interest or process.object_of_interest,
                    movement.trigger or process.trigger,
                    movement.code_reference,
                    movement.additional_notes,
                ]
            )
    return rows


def _build_summary_rows(calculator: CosmicCalculator) -> List[List[object]]:
    return [["Metric", "Value"], ["Total CFP", calculator.total_cfp()]]


def _create_workbook(sheets: List[_Sheet], output_path: Path) -> Path:
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types_xml(len(sheets)))
        archive.writestr("_rels/.rels", _rels_xml())
        archive.writestr("xl/workbook.xml", _workbook_xml(sheets))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels_xml(len(sheets)))
        archive.writestr("xl/styles.xml", _styles_xml())

        for index, sheet in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _sheet_xml(sheet))
    return output_path


def export_to_excel(measurement: SystemMeasurement, path: Path | str) -> Path:
    """Export the COSMIC measurement to an Excel workbook."""
    output_path = Path(path)
    calculator = CosmicCalculator(measurement)
    summaries = calculator.summarize().values()

    sheets = [
        _Sheet("Summary", _build_summary_rows(calculator)),
        _Sheet("Functional Processes", _build_functional_process_rows(summaries)),
        _Sheet("Data Movements", _build_data_movement_rows(measurement)),
    ]
    return _create_workbook(sheets, output_path)


class ExcelExporter:
    """High-level interface for exporting to Excel."""

    def __init__(self, measurement: SystemMeasurement) -> None:
        self.measurement = measurement

    def export(self, path: Path | str) -> Path:
        return export_to_excel(self.measurement, path)
