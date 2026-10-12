"""Create schema-consistent copies of the quarterly source Excel workbooks."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from convert_pdfs import OUTPUT_COLUMNS, PROJECT_DIR, _write_filterable_workbook


SOURCE_DIR = PROJECT_DIR / "Junctoin Blacksite list (2004-2024)"
OUTPUT_DIR = PROJECT_DIR / "standardized_xlsx"
REGIONS = {
    "hong kong": "Hong Kong",
    "kowloon": "Kowloon",
    "new territories": "New Territories",
}
PERIOD_RE = re.compile(r"yearly data ending at q([1-4])[- ](\d{4})", re.IGNORECASE)
FILE_PERIOD_RE = re.compile(r"Q([1-4])[- ](20\d{2})", re.IGNORECASE)
SEVERITY_COLUMNS = ("fatal", "serious", "slight", "total")


def _text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def _period_from_file(source_file: Path) -> tuple[int, int]:
    match = FILE_PERIOD_RE.search(source_file.name)
    if match is None:
        raise ValueError(f"Cannot determine year and quarter from {source_file.name}")
    return int(match.group(2)), int(match.group(1))


def _find_current_period_blocks(
    worksheet, header_row: int, year: int, quarter: int
) -> tuple[dict[str, dict[str, int]], int]:
    top_headers = [
        (cell.column, match)
        for cell in worksheet[header_row]
        if (match := PERIOD_RE.search(_text(cell.value))) is not None
    ]
    matches = [
        (column, match)
        for column, match in top_headers
        if int(match.group(1)) == quarter and int(match.group(2)) == year
    ]
    if not matches:
        return {}, header_row
    if len(matches) > 1:
        raise ValueError(
            f"Multiple current-period headers found in worksheet {worksheet.title!r}"
        )

    period_column, _ = matches[0]
    later_period_columns = [
        column for column, _ in top_headers if column > period_column
    ]
    block_end = min(later_period_columns, default=worksheet.max_column + 1)
    group_row = header_row + 1
    severity_row = header_row + 2
    groups: dict[str, dict[str, int]] = {}
    for cell in worksheet[group_row]:
        if not period_column <= cell.column < block_end:
            continue
        label = _text(cell.value).casefold()
        if label == "number of accidents":
            group_name = "accidents"
        elif label == "number of accidents involving pedestrians":
            group_name = "pedestrians"
        else:
            continue

        category_columns: dict[str, int] = {}
        for offset in range(len(SEVERITY_COLUMNS)):
            column = cell.column + offset
            category = _text(worksheet.cell(severity_row, column).value).casefold()
            if category in SEVERITY_COLUMNS:
                category_columns[category] = column
        if set(category_columns) != set(SEVERITY_COLUMNS):
            raise ValueError(
                f"Incomplete {group_name} metric headers in "
                f"{worksheet.title!r} at row {severity_row}"
            )
        groups[group_name] = category_columns

    return groups, severity_row


def _header_columns(worksheet) -> tuple[int, int, int, int] | None:
    for row in worksheet.iter_rows():
        values = {_text(cell.value).casefold(): cell.column for cell in row if cell.value}
        site_id_column = values.get("site id")
        junction_column = values.get("district & junction identification")
        if site_id_column is not None and junction_column is not None:
            return row[0].row, site_id_column, junction_column, junction_column + 1
    return None


def _is_five_year_only_table(worksheet, header_row: int) -> bool:
    period_header = any(
        "past five" in _text(cell.value).casefold()
        for cell in worksheet[header_row]
    )
    fatal_header = any(
        _text(cell.value).casefold() == "fatal"
        for row in worksheet.iter_rows(
            min_row=header_row + 1, max_row=header_row + 2
        )
        for cell in row
    )
    return period_header and fatal_header


def _metric_value(worksheet, row: int, column: int | None) -> int | None:
    if column is None:
        return None
    value = worksheet.cell(row, column).value
    if value is None or _text(value) == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(
            f"Expected a numeric metric at {worksheet.title}!{worksheet.cell(row, column).coordinate}, "
            f"found {value!r}"
        )
    if not float(value).is_integer():
        raise ValueError(
            f"Expected a whole-number metric at "
            f"{worksheet.title}!{worksheet.cell(row, column).coordinate}, found {value!r}"
        )
    return int(value)


def _split_junction(value: object) -> tuple[str, str]:
    parts = _text(value).split(",", maxsplit=1)
    return parts[0].strip(), parts[1].strip() if len(parts) == 2 else ""


def _extract_rows(source_file: Path) -> list[dict[str, object]]:
    year, quarter = _period_from_file(source_file)
    rows: list[dict[str, object]] = []
    workbook = load_workbook(source_file, read_only=False, data_only=True)
    try:
        for worksheet in workbook.worksheets:
            header = _header_columns(worksheet)
            if header is None:
                continue
            header_row, site_id_column, district_column, junction_column = header
            groups, last_header_row = _find_current_period_blocks(
                worksheet, header_row, year, quarter
            )
            five_year_only = not groups and _is_five_year_only_table(
                worksheet, header_row
            )
            if not groups and not five_year_only:
                continue

            region = ""
            district = ""
            first_data_row = (
                last_header_row + 1 if groups else header_row + 1
            )
            for row_number in range(first_data_row, worksheet.max_row + 1):
                site_id = worksheet.cell(row_number, site_id_column).value
                district_or_group = worksheet.cell(row_number, district_column).value
                junction = worksheet.cell(row_number, junction_column).value

                if site_id is None and junction is None and district_or_group is not None:
                    heading = _text(district_or_group)
                    region = REGIONS.get(heading.casefold(), region)
                    if heading.casefold() not in REGIONS:
                        district = heading.title()
                    continue

                if junction is None or _text(junction) == "":
                    continue

                street_1, street_2 = _split_junction(junction)
                record: dict[str, object] = {
                    "year": year,
                    "quarter": quarter,
                    "region": region or None,
                    "district": district or None,
                    "site_id": site_id,
                    "street_1": street_1,
                    "street_2": street_2 or None,
                    "accidents_total": None,
                    "accidents_pedestrians": None,
                    "accidents_fatal_serious": None,
                    "accidents_fatal": None,
                    "accidents_serious": None,
                    "accidents_slight": None,
                    "casualties_fatal": None,
                    "casualties_serious": None,
                    "casualties_slight": None,
                    "source_file": source_file.relative_to(PROJECT_DIR).as_posix(),
                    "source_page": None,
                }

                accident_metrics = groups.get("accidents")
                if accident_metrics is not None:
                    for severity in SEVERITY_COLUMNS:
                        value = _metric_value(
                            worksheet, row_number, accident_metrics[severity]
                        )
                        if severity == "total":
                            record["accidents_total"] = value
                        else:
                            record[f"accidents_{severity}"] = value
                    fatal = record["accidents_fatal"]
                    serious = record["accidents_serious"]
                    if fatal is not None and serious is not None:
                        record["accidents_fatal_serious"] = fatal + serious

                pedestrian_metrics = groups.get("pedestrians")
                if pedestrian_metrics is not None:
                    record["accidents_pedestrians"] = _metric_value(
                        worksheet, row_number, pedestrian_metrics["total"]
                    )

                rows.append(record)
    finally:
        workbook.close()

    if not rows:
        raise ValueError(
            f"No current-quarter annual records were found in {source_file.name}"
        )
    return rows


def standardize_all_excels() -> list[Path]:
    source_files = sorted(SOURCE_DIR.glob("20[1-2][0-9]/*.xlsx"))
    if not source_files:
        raise FileNotFoundError(f"No quarterly Excel workbooks found under {SOURCE_DIR}")

    standardized_files = []
    for source_file in source_files:
        records = _extract_rows(source_file)
        data = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
        year, quarter = _period_from_file(source_file)
        destination = OUTPUT_DIR / source_file.relative_to(SOURCE_DIR)
        table_name = f"Standardized_{year}_Q{quarter}"
        _write_filterable_workbook(data, destination, table_name)
        standardized_files.append(destination)
        print(f"Standardized {source_file.name}: {len(data)} annual records")

    return standardized_files


if __name__ == "__main__":
    files = standardize_all_excels()
    print(f"\nCreated {len(files)} standardized copies under {OUTPUT_DIR}")
