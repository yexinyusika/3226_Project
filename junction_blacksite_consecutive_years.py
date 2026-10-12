"""Create one output row for each uninterrupted blacksite-list period."""

from pathlib import Path
import re

import pandas as pd

INPUT_FILE = Path("Junction_Blacksite_Master_Q4_2004-2024.xlsx")
OUTPUT_FILE = Path("Junction_Blacksite_Consecutive_Years_On_List.xlsx")
OUTPUT_COLUMNS = [
    "region",
    "district",
    "site_id",
    "street_1",
    "street_2",
    "number of years on the blacksite list",
    "starting year on the blacksite list",
    "ending year on the blacksite list",
    "accidents_total(accumulated)",
    "accidents_pedestrians(accumulated)",
]


def normalize_street(value: object) -> str:
    if pd.isna(value):
        return ""
    street = str(value).upper().strip().replace(".", "")
    street = re.sub(r"\bST\b", "STREET", street)
    street = re.sub(r"\bRD\b", "ROAD", street)
    return re.sub(r"\s+", " ", street).strip()


def build_group_key(row: pd.Series) -> tuple[object, ...]:
    return (
        row["region"],
        row["district"],
        *sorted((row["street_1"], row["street_2"])),
    )


def main() -> None:
    df = pd.read_excel(INPUT_FILE)
    required_columns = {
        "year",
        "region",
        "district",
        "site_id",
        "street_1",
        "street_2",
        "accidents_total",
        "accidents_pedestrians",
    }
    missing_columns = required_columns.difference(df.columns)
    if missing_columns:
        raise ValueError(f"Missing required columns: {sorted(missing_columns)}")

    df = df.copy()
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    if df["year"].isna().any():
        raise ValueError("The input contains rows without a valid year.")
    df["site_id"] = pd.to_numeric(df["site_id"], errors="coerce")
    for column in ["region", "district"]:
        df[column] = df[column].fillna("")
    for column in ["street_1", "street_2"]:
        df[column] = df[column].map(normalize_street)

    if df[["street_1", "street_2"]].eq("").any(axis=None):
        raise ValueError("The input contains junctions with a missing street name.")

    df["group_key"] = df.apply(build_group_key, axis=1)
    periods = []
    for _, junction in df.groupby("group_key", sort=False):
        junction = junction.sort_values("year", kind="mergesort").copy()
        junction["_period"] = junction["year"].diff().ne(1).cumsum()
        street_1, street_2 = sorted(
            (junction.iloc[0]["street_1"], junction.iloc[0]["street_2"])
        )

        for _, period in junction.groupby("_period", sort=False):
            first = period.iloc[0]
            site_ids = period["site_id"].dropna()
            start_year = int(period["year"].min())
            end_year = int(period["year"].max())
            periods.append(
                {
                    "region": first["region"],
                    "district": first["district"],
                    "site_id": site_ids.iloc[0] if not site_ids.empty else pd.NA,
                    "street_1": street_1,
                    "street_2": street_2,
                    "number of years on the blacksite list": period["year"].nunique(),
                    "starting year on the blacksite list": start_year,
                    "ending year on the blacksite list": end_year,
                    "accidents_total(accumulated)": period["accidents_total"].sum(),
                    "accidents_pedestrians(accumulated)": period[
                        "accidents_pedestrians"
                    ].sum(),
                }
            )

    result = pd.DataFrame(periods, columns=OUTPUT_COLUMNS)
    result = result.sort_values(
        by=[
            "number of years on the blacksite list",
            "starting year on the blacksite list",
        ],
        ascending=[False, True],
        kind="mergesort",
    )
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(result)} consecutive junction periods to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
