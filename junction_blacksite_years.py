from pathlib import Path

import pandas as pd

INPUT_FILE = Path("Junction_Blacksite_Master_Q4_2004-2024.xlsx")
OUTPUT_FILE = Path("Junction_Blacksite_Years_On_List.xlsx")


def build_group_key(row: pd.Series) -> tuple[str, object]:
    site_id = row.get("site_id")
    if pd.notna(site_id):
        return ("site", int(float(site_id)))
    return (
        "road",
        row.get("region", ""),
        row.get("district", ""),
        row.get("street_1", ""),
        row.get("street_2", ""),
    )


def main() -> None:
    df = pd.read_excel(INPUT_FILE)
    df = df.copy()
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df["site_id"] = pd.to_numeric(df["site_id"], errors="coerce")
    df["region"] = df["region"].fillna("")
    df["district"] = df["district"].fillna("")
    df["street_1"] = df["street_1"].fillna("")
    df["street_2"] = df["street_2"].fillna("")

    df["group_key"] = df.apply(build_group_key, axis=1)
    df = df.sort_values(["group_key", "year"], kind="mergesort")

    agg = (
        df.groupby("group_key", as_index=False)
        .agg(
            region=("region", "first"),
            district=("district", "first"),
            site_id=("site_id", "first"),
            street_1=("street_1", "first"),
            street_2=("street_2", "first"),
            number_of_years=("year", "nunique"),
            starting_year=("year", "min"),
            ending_year=("year", "max"),
            accidents_total_accumulated=("accidents_total", "sum"),
            accidents_pedestrians_accumulated=("accidents_pedestrians", "sum"),
        )
        .rename(
            columns={
                "number_of_years": "number of years on the blacksite list",
                "starting_year": "starting year on the blacksite list",
                "ending_year": "ending year on the blacksite list",
                "accidents_total_accumulated": "accidents_total(accumulated)",
                "accidents_pedestrians_accumulated": "accidents_pedestrians(accumulated)",
            }
        )
    )

    agg["site_id"] = pd.to_numeric(agg["site_id"], errors="coerce")
    agg = agg.sort_values(
        by=["number of years on the blacksite list", "starting year on the blacksite list"],
        ascending=[False, True],
        kind="mergesort",
    )

    output_columns = [
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
    agg = agg.loc[:, output_columns]
    agg.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(agg)} junction rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
