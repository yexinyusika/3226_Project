"""Collect annual ATC station AADT records matching consecutive blacksite junctions."""

from __future__ import annotations

import argparse
from copy import copy
from dataclasses import dataclass
from html.parser import HTMLParser
import re
import tempfile
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, unquote, urljoin, urlsplit, urlunsplit
from urllib.request import Request, urlopen

import pymupdf
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_DIR / "Junction_Blacksite_Consecutive_Years_On_List.xlsx"
DEFAULT_OUTPUT = PROJECT_DIR / "Junction_Blacksite_Consecutive_Years_AADT_2004-2024.xlsx"
DEFAULT_CACHE = Path.home() / ".cache" / "hk_atc_aadt"
ARCHIVE_URL = (
    "https://www.td.gov.hk/en/publications_and_press_releases/publications/"
    "free_publications/atci2/index.html"
)

TARGET_JUNCTIONS = (
    {
        "display": "TSUEN TSING INTERCHANGE / TSUEN WAN ROAD",
        "street_1": "TSUEN TSING INTERCHANGE",
        "street_2": "TSUEN WAN ROAD",
        "side_1_aliases": ("Tsuen Tsing INT",),
        "side_2_aliases": ("Tsuen Wan Rd",),
    },
    {
        "display": "MONG KOK ROAD / NATHAN ROAD",
        "street_1": "MONG KOK ROAD",
        "street_2": "NATHAN ROAD",
        "side_1_aliases": ("Mong Kok Rd",),
        "side_2_aliases": ("Nathan Rd",),
    },
    {
        "display": "AUSTIN ROAD, CHEONG WAN ROAD / CHATHAM ROAD SOUTH",
        "street_1": "AUSTIN ROAD, CHEONG WAN ROAD",
        "street_2": "CHATHAM ROAD SOUTH",
        "side_1_aliases": ("Austin Rd", "Cheong Wan Rd"),
        "side_2_aliases": ("Chatham Rd S",),
    },
    {
        "display": "NATHAN ROAD / WATERLOO ROAD",
        "street_1": "NATHAN ROAD",
        "street_2": "WATERLOO ROAD",
        "side_1_aliases": ("Nathan Rd",),
        "side_2_aliases": ("Waterloo Rd",),
    },
    {
        "display": "ARGYLE STREET / NATHAN ROAD",
        "street_1": "ARGYLE STREET",
        "street_2": "NATHAN ROAD",
        "side_1_aliases": ("Argyle St",),
        "side_2_aliases": ("Nathan Rd",),
    },
    {
        "display": "ARGYLE STREET, PRINCESS MARGARET ROAD / WATERLOO ROAD",
        "street_1": "ARGYLE STREET, PRINCESS MARGARET ROAD",
        "street_2": "WATERLOO ROAD",
        "side_1_aliases": ("Argyle St", "Princess Margaret Rd"),
        "side_2_aliases": ("Waterloo Rd",),
    },
    {
        "display": "TEXACO ROAD NORTH / TSUEN KAM INTERCHANGE",
        "street_1": "TEXACO ROAD NORTH",
        "street_2": "TSUEN KAM INTERCHANGE",
        "side_1_aliases": ("Texaco Rd North", "Texaco Rd N", "Texaco Rd"),
        "side_2_aliases": ("Tsuen Kam INT",),
    },
    {
        "display": "JORDAN ROAD / NATHAN ROAD",
        "street_1": "JORDAN ROAD",
        "street_2": "NATHAN ROAD",
        "side_1_aliases": ("Jordan Rd",),
        "side_2_aliases": ("Nathan Rd",),
    },
    {
        "display": "MA TAU CHUNG ROAD / SUNG WONG TOI ROAD, FU NING STREET",
        "street_1": "MA TAU CHUNG ROAD",
        "street_2": "SUNG WONG TOI ROAD, FU NING STREET",
        "side_1_aliases": ("Ma Tau Chung Rd",),
        "side_2_aliases": ("Sung Wong Toi", "Fu Ning St"),
    },
    {
        "display": "ARGYLE STREET / TONG MI ROAD, CHERRY STREET",
        "street_1": "ARGYLE STREET",
        "street_2": "TONG MI ROAD, CHERRY STREET",
        "side_1_aliases": ("Argyle St",),
        "side_2_aliases": ("Tong Mi Rd", "Tong Mei Rd", "Cherry St"),
    },
)

INPUT_COLUMNS = {
    "region",
    "district",
    "site_id",
    "street_1",
    "street_2",
    "number of years on the blacksite list",
    "starting year on the blacksite list",
    "ending year on the blacksite list",
}

AADT_COLUMNS = [
    "Junction",
    "Street 1 (blacksite list)",
    "Street 2 (blacksite list)",
    "Year",
    "AADT",
    "AADT estimated by growth factor",
    "ATC station number",
    "ATC road name",
    "ATC from",
    "ATC to",
    "Road type",
    "Station type",
    "ATC PDF page",
    "ATC report URL",
]

COVERAGE_COLUMNS = [
    "Junction",
    "Year",
    "Expected ATC station matches",
    "Status",
    "ATC report URL",
]


@dataclass(frozen=True)
class Link:
    href: str
    text: str


@dataclass(frozen=True)
class Junction:
    display: str
    street_1: str
    street_2: str
    start_year: int
    end_year: int
    side_1_aliases: tuple[str, ...]
    side_2_aliases: tuple[str, ...]


@dataclass(frozen=True)
class Report:
    year: int
    page_url: str
    pdf_url: str


@dataclass(frozen=True)
class StationRecord:
    station_number: str
    road_name: str
    from_name: str
    to_name: str
    road_type: str
    station_type: str
    aadt: int | None
    estimated: bool
    pdf_page: int


class AnchorParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[Link] = []
        self._current_href: str | None = None
        self._current_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "a":
            self._current_href = dict(attrs).get("href") or ""
            self._current_text = []

    def handle_data(self, data: str) -> None:
        if self._current_href is not None:
            self._current_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._current_href is not None:
            text = " ".join(" ".join(self._current_text).split())
            self.links.append(Link(self._current_href, text))
            self._current_href = None
            self._current_text = []


def fetch_url(url: str) -> tuple[bytes, str]:
    request = Request(
        url,
        headers={"User-Agent": "Junction-AADT-research/1.0 (public data collection)"},
    )
    with urlopen(request, timeout=90) as response:
        return response.read(), response.geturl()


def fetch_links(url: str) -> list[Link]:
    content, final_url = fetch_url(url)
    parser = AnchorParser()
    parser.feed(content.decode("utf-8", errors="replace"))
    links = []
    for link in parser.links:
        if not link.href:
            continue
        parts = urlsplit(urljoin(final_url, link.href))
        escaped_path = quote(parts.path, safe="/%:@!$&'()*+,;=-._~")
        links.append(
            Link(
                urlunsplit(
                    (parts.scheme, parts.netloc, escaped_path, parts.query, parts.fragment)
                ),
                link.text,
            )
        )
    return links


def normalize_name(value: object) -> str:
    text = str(value).upper()
    for old, new in (
        ("INTERCHANGE", "INT"),
        ("STREET", "ST"),
        ("ROAD", "RD"),
        ("NORTH", "N"),
        ("SOUTH", "S"),
    ):
        text = re.sub(rf"\b{old}\b", new, text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def has_alias(field: str, aliases: Iterable[str]) -> bool:
    normalized_field = f" {normalize_name(field)} "
    return any(
        f" {normalize_name(alias)} " in normalized_field
        for alias in aliases
    )


def read_junctions(input_file: Path) -> list[Junction]:
    workbook = load_workbook(input_file, read_only=True, data_only=True)
    worksheet = workbook.active
    header = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
    missing_columns = INPUT_COLUMNS.difference(str(name) for name in header)
    if missing_columns:
        raise ValueError(f"{input_file} is missing columns: {sorted(missing_columns)}")
    column_indexes = {str(name): index for index, name in enumerate(header)}
    by_pair: dict[tuple[str, str], list[dict[str, object]]] = {}
    for values in worksheet.iter_rows(min_row=2, values_only=True):
        row = dict(zip(column_indexes, values))
        pair = tuple(
            sorted((normalize_name(row["street_1"]), normalize_name(row["street_2"])))
        )
        by_pair.setdefault(pair, []).append(row)
    workbook.close()

    junctions: list[Junction] = []
    for target in TARGET_JUNCTIONS:
        pair = tuple(
            sorted(
                (
                    normalize_name(target["street_1"]),
                    normalize_name(target["street_2"]),
                )
            )
        )
        candidates = [
            row
            for row in by_pair.get(pair, [])
            if isinstance(row["number of years on the blacksite list"], (int, float))
            and row["number of years on the blacksite list"] >= 7
        ]
        if len(candidates) != 1:
            raise ValueError(
                f"Expected exactly one consecutive-list row for "
                f"{target['display']}; found {len(candidates)}."
            )
        row = candidates[0]
        start = row["starting year on the blacksite list"]
        end = row["ending year on the blacksite list"]
        if not isinstance(start, (int, float)) or not isinstance(end, (int, float)):
            raise ValueError(f"Invalid year range for {target['display']}.")
        if end - start + 1 < 7:
            raise ValueError(
                f"{target['display']} is on the consecutive list for fewer than 7 years."
            )
        listed_year_count = int(row["number of years on the blacksite list"])
        if listed_year_count != end - start + 1:
            raise ValueError(
                f"The consecutive year count does not match the listed range for "
                f"{target['display']}."
            )
        junctions.append(
            Junction(
                display=target["display"],
                street_1=str(row["street_1"]),
                street_2=str(row["street_2"]),
                start_year=int(start),
                end_year=int(end),
                side_1_aliases=target["side_1_aliases"],
                side_2_aliases=target["side_2_aliases"],
            )
        )
    return junctions


def discover_reports(archive_url: str, years: set[int]) -> dict[int, Report]:
    archive_links = fetch_links(archive_url)
    page_by_year: dict[int, str] = {}
    for link in archive_links:
        match = re.search(r"annual\s+traffic\s+census\D*(20\d{2})", link.text, re.I)
        if match:
            year = int(match.group(1))
            if year in years:
                page_by_year[year] = link.href

    missing_years = years.difference(page_by_year)
    if missing_years:
        raise ValueError(
            f"The ATC catalogue did not link to report pages for: {sorted(missing_years)}"
        )

    reports: dict[int, Report] = {}
    for year in sorted(years):
        page_url = page_by_year[year]
        candidates = []
        for link in fetch_links(page_url):
            path = unquote(urlsplit(link.href).path).lower()
            label = f"{path} {link.text}".lower()
            if not path.endswith(".pdf") or "amendment" in label or "revision" in label:
                continue
            if str(year) not in label or "traffic census" not in label:
                continue
            candidates.append(link.href)
        if not candidates:
            raise ValueError(
                f"No main Annual Traffic Census PDF was found on the {year} publication page "
                f"{page_url}."
            )
        if len(candidates) > 1:
            raise ValueError(
                f"Found multiple possible main Annual Traffic Census PDFs for {year} "
                f"on {page_url}: {candidates}"
            )
        reports[year] = Report(year, page_url, candidates[0])
    return reports


def download_report(report: Report, cache_dir: Path) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = cache_dir / f"ATC_{report.year}.pdf"
    if pdf_path.is_file() and pdf_path.stat().st_size > 4:
        with pdf_path.open("rb") as cached:
            if cached.read(4) == b"%PDF":
                return pdf_path
        raise ValueError(f"Cached report is not a PDF: {pdf_path}")

    content, _ = fetch_url(report.pdf_url)
    if not content.startswith(b"%PDF"):
        raise ValueError(
            f"The ATC report for {report.year} did not return a PDF: {report.pdf_url}"
        )
    with tempfile.NamedTemporaryFile(
        mode="wb", dir=cache_dir, prefix=f".ATC_{report.year}_", delete=False
    ) as temporary:
        temporary.write(content)
        temporary_path = Path(temporary.name)
    temporary_path.replace(pdf_path)
    return pdf_path


def appendix_pages(document: pymupdf.Document, year: int) -> list[int]:
    pages = []
    for page_number, page in enumerate(document):
        text = page.get_text()
        compact = re.sub(r"\s+", " ", text).upper()
        if "AADT OF COUNTING STATIONS" in compact and "APPENDIX C" in compact:
            pages.append(page_number)
    if not pages:
        raise ValueError(
            f"Could not locate Appendix C (AADT of Counting Stations) in the {year} report."
        )
    return pages


def _column_text(words: list[tuple], left: float, right: float) -> str:
    selected = sorted(
        (word for word in words if left <= word[0] < right),
        key=lambda word: (word[1], word[0]),
    )
    return " ".join(str(word[4]) for word in selected).strip()


def _parse_aadt(words: list[tuple]) -> tuple[int | None, bool]:
    current_year_words = [
        word for word in words if 400 <= word[0] < 455
    ]
    values = [
        word
        for word in current_year_words
        if re.fullmatch(r"\d[\d,]*\*?", str(word[4]))
    ]
    if not values:
        return None, False
    value_word = max(values, key=lambda word: word[0])
    text = str(value_word[4])
    estimated = text.endswith("*") or any(
        str(word[4]) == "*" and abs(word[1] - value_word[1]) <= 3
        for word in current_year_words
    )
    number = text.rstrip("*").replace(",", "")
    return int(number), estimated


def extract_station_records(
    document: pymupdf.Document, year: int
) -> list[StationRecord]:
    records: list[StationRecord] = []
    for page_index in appendix_pages(document, year):
        page = document[page_index]
        words = page.get_text("words", sort=True)
        station_words = sorted(
            (
                word
                for word in words
                if 280 <= word[0] < 312
                and re.fullmatch(r"\d{3,5}", str(word[4]))
            ),
            key=lambda word: word[1],
        )
        for index, station_word in enumerate(station_words):
            station_y = station_word[1]
            previous_y = station_words[index - 1][1] if index else station_y - 12
            next_y = (
                station_words[index + 1][1]
                if index + 1 < len(station_words)
                else station_y + 12
            )
            top = (previous_y + station_y) / 2
            bottom = (station_y + next_y) / 2
            row_words = [
                word
                for word in words
                if top <= word[1] < bottom
            ]
            aadt, estimated = _parse_aadt(row_words)
            road_name = _column_text(row_words, 48, 130)
            from_name = _column_text(row_words, 130, 205)
            to_name = _column_text(row_words, 205, 280)
            if not any((road_name, from_name, to_name)):
                continue
            records.append(
                StationRecord(
                    station_number=str(station_word[4]),
                    road_name=road_name,
                    from_name=from_name,
                    to_name=to_name,
                    road_type=_column_text(row_words, 312, 340),
                    station_type=_column_text(row_words, 340, 360),
                    aadt=aadt,
                    estimated=estimated,
                    pdf_page=page_index + 1,
                )
            )
    if not records:
        raise ValueError(f"No station AADT records were parsed from the {year} report.")
    return records


def matches_junction(record: StationRecord, junction: Junction) -> bool:
    boundary_roads = f"{record.from_name} {record.to_name}"
    return (
        has_alias(record.road_name, junction.side_1_aliases)
        and has_alias(boundary_roads, junction.side_2_aliases)
    ) or (
        has_alias(record.road_name, junction.side_2_aliases)
        and has_alias(boundary_roads, junction.side_1_aliases)
    )


def add_excel_table(worksheet, name: str, last_row: int, last_column: int) -> None:
    table = Table(
        displayName=name,
        ref=f"A1:{worksheet.cell(last_row, last_column).coordinate}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def format_worksheet(worksheet, headers: list[str]) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for column_index, header in enumerate(headers, start=1):
        width = max(14, min(44, len(header) + 2))
        worksheet.column_dimensions[worksheet.cell(1, column_index).column_letter].width = width
    worksheet.sheet_view.showGridLines = False


def write_workbook(
    output_file: Path,
    junctions: list[Junction],
    reports: dict[int, Report],
    matches: dict[tuple[str, int], list[StationRecord]],
) -> None:
    workbook = Workbook()
    aadt_sheet = workbook.active
    aadt_sheet.title = "Yearly AADT"
    coverage_sheet = workbook.create_sheet("Coverage")
    notes_sheet = workbook.create_sheet("Read Me")
    aadt_sheet.append(AADT_COLUMNS)
    coverage_sheet.append(COVERAGE_COLUMNS)

    for junction in junctions:
        for year in range(junction.start_year, junction.end_year + 1):
            station_matches = matches[(junction.display, year)]
            report_url = reports[year].pdf_url
            if not station_matches:
                coverage_sheet.append(
                    [junction.display, year, 0, "No exact station-segment match", report_url]
                )
                aadt_sheet.append(
                    [
                        junction.display,
                        junction.street_1,
                        junction.street_2,
                        year,
                        None,
                        None,
                        None,
                        None,
                        None,
                        None,
                        None,
                        None,
                        None,
                        report_url,
                    ]
                )
                continue

            unavailable_aadt = any(record.aadt is None for record in station_matches)
            if len(station_matches) == 1:
                status = (
                    "Station matched; AADT unavailable"
                    if unavailable_aadt
                    else "One station-segment match"
                )
            else:
                status = (
                    "Multiple station matches; some AADT unavailable"
                    if unavailable_aadt
                    else "Multiple station-segment matches"
                )
            coverage_sheet.append(
                [junction.display, year, len(station_matches), status, report_url]
            )
            for record in station_matches:
                aadt_sheet.append(
                    [
                        junction.display,
                        junction.street_1,
                        junction.street_2,
                        year,
                        record.aadt,
                        record.estimated,
                        record.station_number,
                        record.road_name,
                        record.from_name,
                        record.to_name,
                        record.road_type,
                        record.station_type,
                        record.pdf_page,
                        report_url,
                    ]
                )

    format_worksheet(aadt_sheet, AADT_COLUMNS)
    format_worksheet(coverage_sheet, COVERAGE_COLUMNS)
    for row in aadt_sheet.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "#,##0"
    add_excel_table(aadt_sheet, "YearlyAADT", aadt_sheet.max_row, len(AADT_COLUMNS))
    add_excel_table(
        coverage_sheet,
        "Coverage",
        coverage_sheet.max_row,
        len(COVERAGE_COLUMNS),
    )
    notes_sheet.append(["Topic", "Details"])
    notes_sheet.append(
        [
            "Source list",
            "Junction_Blacksite_Consecutive_Years_On_List.xlsx only; each listed "
            "period of at least 7 consecutive years is taken from that workbook.",
        ]
    )
    notes_sheet.append(
        [
            "ATC source",
            "The annual Transport Department ATC reports, Appendix C: "
            "AADT of Counting Stations - ordered by Road Names.",
        ]
    )
    notes_sheet.append(
        [
            "Meaning of a match",
            "A report row is included when a blacksite street matches the ATC road "
            "name and another blacksite street matches its From/To boundary. "
            "Common ATC abbreviations are normalized.",
        ]
    )
    notes_sheet.append(
        [
            "Multiple values",
            "ATC reports AADT by station and road segment, not as one value per "
            "junction. A junction-year can therefore have multiple rows in "
            "'Yearly AADT'; station and segment details are retained and values "
            "are not summed or averaged.",
        ]
    )
    notes_sheet.append(
        [
            "Missing match",
            "A blank AADT with 'No exact station-segment match' means the report "
            "did not yield a direct road-name match. It is not a zero-traffic value.",
        ]
    )
    format_worksheet(notes_sheet, ["Topic", "Details"])
    notes_sheet.column_dimensions["B"].width = 110
    for row in notes_sheet.iter_rows(min_row=2, min_col=2, max_col=2):
        alignment = copy(row[0].alignment)
        alignment.wrap_text = True
        alignment.vertical = "top"
        row[0].alignment = alignment
    notes_sheet.row_dimensions[5].height = 48
    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)


def collect(
    input_file: Path,
    output_file: Path,
    cache_dir: Path,
    archive_url: str = ARCHIVE_URL,
) -> None:
    junctions = read_junctions(input_file)
    years = {
        year
        for junction in junctions
        for year in range(junction.start_year, junction.end_year + 1)
    }
    reports = discover_reports(archive_url, years)
    matches = {
        (junction.display, year): []
        for junction in junctions
        for year in range(junction.start_year, junction.end_year + 1)
    }

    for year in sorted(years):
        report_path = download_report(reports[year], cache_dir)
        with pymupdf.open(report_path) as document:
            records = extract_station_records(document, year)
        for junction in junctions:
            if junction.start_year <= year <= junction.end_year:
                matches[(junction.display, year)] = [
                    record
                    for record in records
                    if matches_junction(record, junction)
                ]
        matched_junctions = sum(
            bool(matches[(junction.display, year)])
            for junction in junctions
            if junction.start_year <= year <= junction.end_year
        )
        print(
            f"{year}: parsed {len(records)} ATC station rows; "
            f"{matched_junctions} in-range junctions have matches."
        )

    write_workbook(output_file, junctions, reports, matches)
    expected_pairs = sum(
        junction.end_year - junction.start_year + 1 for junction in junctions
    )
    matched_pairs = sum(bool(records) for records in matches.values())
    print(
        f"Saved {expected_pairs} junction-year coverage rows "
        f"({matched_pairs} matched) to {output_file}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--cache-dir", type=Path, default=DEFAULT_CACHE)
    parser.add_argument("--archive-url", default=ARCHIVE_URL)
    args = parser.parse_args()
    collect(args.input, args.output, args.cache_dir, args.archive_url)


if __name__ == "__main__":
    main()
