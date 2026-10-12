"""Create one output row for each uninterrupted blacksite-list period."""

from pathlib import Path

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


def build_group_key(row: pd.Series) -> tuple[object, ...]:
    site_id = row["site_id"]
    if pd.notna(site_id):
        return ("site", int(site_id))
    return (
        "road",
        row["region"],
        row["district"],
        row["street_1"],
        row["street_2"],
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
    for column in ["region", "district", "street_1", "street_2"]:
        df[column] = df[column].fillna("")

    df["group_key"] = df.apply(build_group_key, axis=1)
    periods = []
    for _, junction in df.groupby("group_key", sort=False):
        junction = junction.sort_values("year", kind="mergesort").copy()
        junction["_period"] = junction["year"].diff().ne(1).cumsum()

        for _, period in junction.groupby("_period", sort=False):
            first = period.iloc[0]
            start_year = int(period["year"].min())
            end_year = int(period["year"].max())
            periods.append(
                {
                    "region": first["region"],
                    "district": first["district"],
                    "site_id": first["site_id"],
                    "street_1": first["street_1"],
                    "street_2": first["street_2"],
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
