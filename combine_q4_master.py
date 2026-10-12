"""Combine the standardized Q4 junction blacksite workbooks into one master."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from convert_pdfs import OUTPUT_COLUMNS, PROJECT_DIR, _write_filterable_workbook


SOURCE_DIR = PROJECT_DIR / "Junction Blacksite Q4"
OUTPUT_FILE = PROJECT_DIR / "Junction_Blacksite_Master_Q4_2004-2024.xlsx"
SOURCE_COLUMNS = OUTPUT_COLUMNS[:16]
EXPECTED_YEARS = set(range(2004, 2025))


def combine_q4_workbooks() -> Path:
    source_files = sorted(SOURCE_DIR.glob("*.xlsx"))
    if not source_files:
        raise FileNotFoundError(f"No Excel workbooks found under {SOURCE_DIR}")

    yearly_data = []
    year_counts: dict[int, int] = {}
    for source_file in source_files:
        data = pd.read_excel(source_file, sheet_name="data")
        if tuple(data.columns) != tuple(SOURCE_COLUMNS):
            raise ValueError(
                f"Unexpected columns in {source_file.name}: "
                f"{list(data.columns)!r}"
            )
        if data.empty:
            raise ValueError(f"No records found in {source_file.name}")

        years = set(data["year"].dropna().astype(int).unique())
        quarters = set(data["quarter"].dropna().astype(int).unique())
        if len(years) != 1 or quarters != {4}:
            raise ValueError(
                f"Expected one year and quarter 4 in {source_file.name}; "
                f"found years={sorted(years)}, quarters={sorted(quarters)}"
            )

        year = years.pop()
        year_counts[year] = year_counts.get(year, 0) + 1
        data["source_file"] = source_file.relative_to(PROJECT_DIR).as_posix()
        data["source_page"] = None
        yearly_data.append(data.loc[:, OUTPUT_COLUMNS])

    actual_years = set(year_counts)
    missing_years = sorted(EXPECTED_YEARS - actual_years)
    unexpected_years = sorted(actual_years - EXPECTED_YEARS)
    duplicate_years = sorted(year for year, count in year_counts.items() if count > 1)
    if missing_years or unexpected_years or duplicate_years:
        raise ValueError(
            "Expected exactly one Q4 workbook for each year from 2004 to 2024; "
            f"missing={missing_years}, unexpected={unexpected_years}, "
            f"duplicates={duplicate_years}"
        )

    master = pd.concat(yearly_data, ignore_index=True)
    _write_filterable_workbook(master, OUTPUT_FILE, "Blacksite_Q4_Master")
    print(f"Combined {len(source_files)} workbooks and {len(master)} records.")
    print(f"Master workbook: {OUTPUT_FILE}")
    return OUTPUT_FILE


if __name__ == "__main__":
    combine_q4_workbooks()
