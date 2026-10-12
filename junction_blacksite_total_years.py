"""Combine consecutive blacksite periods into total years per junction."""

from pathlib import Path

import pandas as pd

INPUT_FILE = Path("Junction_Blacksite_Consecutive_Years_On_List.xlsx")
OUTPUT_FILE = Path("Junction_Blacksite_Total_Years_On_List.xlsx")
OUTPUT_COLUMNS = [
    "region",
    "district",
    "site_id",
    "street_1",
    "street_2",
    "number of years on the blacksite list",
    "accidents_total(accumulated)",
    "accidents_pedestrians(accumulated)",
]
GROUP_COLUMNS = [
    "_region_key",
    "_district_key",
    "_street_1_key",
    "_street_2_key",
]


def main() -> None:
    df = pd.read_excel(INPUT_FILE)
    required_columns = set(OUTPUT_COLUMNS) | {
        "starting year on the blacksite list",
        "ending year on the blacksite list",
    }
    missing_columns = required_columns.difference(df.columns)
    if missing_columns:
        raise ValueError(f"Missing required columns: {sorted(missing_columns)}")

    df = df.copy()
    for column in ["region", "district", "street_1", "street_2"]:
        df[column] = df[column].fillna("").astype(str).str.strip()
    df["site_id"] = pd.to_numeric(df["site_id"], errors="coerce")
    numeric_columns = [
        "number of years on the blacksite list",
        "accidents_total(accumulated)",
        "accidents_pedestrians(accumulated)",
    ]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df["_region_key"] = df["region"].str.upper()
    df["_district_key"] = df["district"].str.upper()
    street_pairs = df[["street_1", "street_2"]].apply(
        lambda row: sorted((row["street_1"].upper(), row["street_2"].upper())),
        axis=1,
        result_type="expand",
    )
    df["_street_1_key"] = street_pairs[0]
    df["_street_2_key"] = street_pairs[1]

    result = (
        df.groupby(GROUP_COLUMNS, as_index=False, sort=False, dropna=False)
        .agg(
            region=("region", "first"),
            district=("district", "first"),
            site_id=("site_id", "first"),
            street_1=("_street_1_key", "first"),
            street_2=("_street_2_key", "first"),
            **{
                "number of years on the blacksite list": (
                    "number of years on the blacksite list",
                    "sum",
                ),
                "accidents_total(accumulated)": (
                    "accidents_total(accumulated)",
                    "sum",
                ),
                "accidents_pedestrians(accumulated)": (
                    "accidents_pedestrians(accumulated)",
                    "sum",
                ),
            },
        )
        .sort_values(
            "number of years on the blacksite list",
            ascending=False,
            kind="mergesort",
        )
    )
    result = result.loc[:, OUTPUT_COLUMNS]
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(result)} junctions to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
