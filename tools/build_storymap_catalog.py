#!/usr/bin/env python3
from __future__ import annotations

import csv
import html
import argparse
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT_DIR = ROOT / "arcgis_ready" / "sign_catalog"
DEFAULT_SOURCE_TABLE = DEFAULT_OUT_DIR / "sign_catalog_table.xlsm"
DEFAULT_OUT_CSV = DEFAULT_OUT_DIR / "sign_catalog_table.csv"

REQUIRED_FIELDS = {
    "Pole_id",
    "Sign_id",
    "Photoname",
    "Transit_category",
    "photo_url",
    "details",
}

CATEGORY_ORDER = ["greenline_related", "bus_related", "other"]
CATEGORY_LABELS = {
    "greenline_related": "Greenline",
    "bus_related": "Bus",
    "other": "Other",
}
CATEGORY_COLORS = {
    "greenline_related": "#1f7a4d",
    "bus_related": "#1f4e79",
    "other": "#59636e",
}

XML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def esc(value: object) -> str:
    return html.escape(str(value or ""))


def normalize_category(value: str) -> str:
    value = (value or "").strip().lower()
    return value if value in CATEGORY_ORDER else "other"


def normalize_detail(value: str) -> str:
    value = (value or "").strip()
    return value if value else "N/A"


def int_or_zero(value: str) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def validate_fields(fieldnames: list[str] | None, source_table: Path) -> None:
    actual = set(fieldnames or [])
    missing = sorted(REQUIRED_FIELDS - actual)
    if missing:
        missing_list = ", ".join(missing)
        raise ValueError(f"{source_table} is missing required columns: {missing_list}")


def load_csv_rows(source_csv: Path) -> list[dict[str, str]]:
    with source_csv.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        validate_fields(reader.fieldnames, source_csv)
        return list(reader)


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("main:si", XML_NS):
        strings.append("".join(text.text or "" for text in item.findall(".//main:t", XML_NS)))
    return strings


def first_worksheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    first_sheet = workbook.find("main:sheets/main:sheet", XML_NS)
    if first_sheet is None:
        raise ValueError("Workbook has no worksheets.")

    relation_id = first_sheet.attrib.get(f"{{{XML_NS['rel']}}}id")
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pkgrel:Relationship", XML_NS):
        if rel.attrib.get("Id") == relation_id:
            target = rel.attrib["Target"].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise ValueError(f"Could not find worksheet relationship {relation_id}.")


def column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    index = 0
    for char in match.group(1):
        index = index * 26 + ord(char) - ord("A") + 1
    return index - 1


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//main:t", XML_NS))

    value = cell.find("main:v", XML_NS)
    raw = "" if value is None or value.text is None else value.text
    if cell_type == "s" and raw:
        return shared_strings[int(raw)]
    return raw


def load_excel_rows(source_excel: Path) -> list[dict[str, str]]:
    with zipfile.ZipFile(source_excel) as archive:
        shared_strings = read_shared_strings(archive)
        sheet_path = first_worksheet_path(archive)
        sheet = ET.fromstring(archive.read(sheet_path))

    table: list[list[str]] = []
    for row in sheet.findall(".//main:sheetData/main:row", XML_NS):
        values: list[str] = []
        for cell in row.findall("main:c", XML_NS):
            index = column_index(cell.attrib["r"])
            while len(values) <= index:
                values.append("")
            values[index] = cell_value(cell, shared_strings)
        table.append(values)

    if not table:
        return []

    headers = [header.strip() for header in table[0]]
    validate_fields(headers, source_excel)
    rows = []
    for values in table[1:]:
        if not any(value.strip() for value in values):
            continue
        padded = values + [""] * (len(headers) - len(values))
        rows.append({header: padded[index].strip() for index, header in enumerate(headers) if header})
    return rows


def load_raw_rows(source_table: Path) -> list[dict[str, str]]:
    suffix = source_table.suffix.lower()
    if suffix == ".csv":
        return load_csv_rows(source_table)
    if suffix in {".xlsm", ".xlsx"}:
        return load_excel_rows(source_table)
    raise ValueError(f"Unsupported source table format: {source_table.suffix}")


def normalize_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    rows = [dict(row) for row in rows]

    for row in rows:
        row["Transit_category"] = normalize_category(row.get("Transit_category", ""))
        row["details"] = normalize_detail(row.get("details", ""))

    rows.sort(
        key=lambda row: (
            CATEGORY_ORDER.index(row["Transit_category"]),
            row["details"],
            int_or_zero(row.get("Pole_id", "")),
            int_or_zero(row.get("Sign_id", "")),
        )
    )
    return rows


def load_rows(source_table: Path) -> list[dict[str, str]]:
    return normalize_rows(load_raw_rows(source_table))


def grouped(rows: list[dict[str, str]]) -> dict[str, dict[str, list[dict[str, str]]]]:
    groups: dict[str, dict[str, list[dict[str, str]]]] = defaultdict(lambda: defaultdict(list))
    for row in rows:
        groups[row["Transit_category"]][row["details"]].append(row)
    return groups


def sample_for(rows: list[dict[str, str]]) -> dict[str, str]:
    row = rows[0]
    return {
        "photo_name": row.get("Photoname", ""),
        "photo_url": row.get("photo_url", ""),
        "pole_id": row.get("Pole_id", ""),
        "sign_id": row.get("Sign_id", ""),
        "status": row.get("status", ""),
    }


def write_summary(rows: list[dict[str, str]], out_summary: Path) -> None:
    groups = grouped(rows)
    with out_summary.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "Transit_category",
                "Transit_category_count",
                "details",
                "details_count",
                "sample_photo",
                "sample_photo_url",
            ],
        )
        writer.writeheader()
        for category in CATEGORY_ORDER:
            detail_groups = groups.get(category, {})
            category_count = sum(len(items) for items in detail_groups.values())
            for detail in sorted(detail_groups):
                items = detail_groups[detail]
                sample = sample_for(items)
                writer.writerow(
                    {
                        "Transit_category": category,
                        "Transit_category_count": category_count,
                        "details": detail,
                        "details_count": len(items),
                        "sample_photo": sample["photo_name"],
                        "sample_photo_url": sample["photo_url"],
                    }
                )


def write_csv(rows: list[dict[str, str]], out_csv: Path) -> None:
    fieldnames: list[str] = []
    for row in rows:
        for field in row:
            if field not in fieldnames:
                fieldnames.append(field)

    with out_csv.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_html(rows: list[dict[str, str]], out_html: Path) -> None:
    groups = grouped(rows)
    total = len(rows)
    category_counts = Counter(row["Transit_category"] for row in rows)

    parts = [
        """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Boston E Branch Sign Catalog</title>
<style>
  :root {
    font-family: Arial, Helvetica, sans-serif;
    color: #1f2933;
    background: #f7f8fa;
  }
  body { margin: 0; }
  header {
    background: #102a43;
    color: white;
    padding: 28px 34px;
  }
  h1 { margin: 0 0 8px; font-size: 30px; letter-spacing: 0; }
  .subtitle { margin: 0; color: #d9e2ec; font-size: 15px; line-height: 1.45; }
  main { max-width: 1260px; margin: 0 auto; padding: 26px 22px 44px; }
  .overview {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    margin-bottom: 24px;
  }
  .metric {
    background: white;
    border: 1px solid #d9e2ec;
    border-radius: 6px;
    padding: 15px 16px;
  }
  .metric .count { display: block; font-size: 31px; font-weight: 700; line-height: 1; }
  .metric .label { display: block; margin-top: 7px; color: #52606d; font-size: 13px; }
  section {
    background: white;
    border: 1px solid #d9e2ec;
    border-radius: 6px;
    overflow: hidden;
    margin-bottom: 24px;
  }
  .section-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    padding: 18px 20px;
    border-top: 6px solid var(--accent);
    border-bottom: 1px solid #e4e7eb;
  }
  h2 { margin: 0; font-size: 22px; }
  .badge {
    background: #eef2f7;
    border-radius: 999px;
    padding: 5px 10px;
    font-weight: 700;
    font-size: 13px;
    white-space: nowrap;
  }
  .catalog-table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    background: #dcefd6;
  }
  th, td {
    text-align: left;
    vertical-align: middle;
    padding: 9px 10px;
    border: 1px solid #b8cdb2;
    font-size: 15px;
  }
  th {
    background: #2f7520;
    color: #fff;
    font-weight: 700;
    font-size: 16px;
  }
  tbody tr:last-child td { border-bottom: 0; }
  .category-cell {
    font-weight: 700;
    color: #000;
    text-align: center;
    font-size: 19px;
    background: #dcefd6;
  }
  .count-cell { text-align: right; font-variant-numeric: tabular-nums; }
  .sample {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 7px;
    text-align: center;
  }
  .sample img {
    width: 160px;
    height: 130px;
    object-fit: contain;
    background: #f5f7fa;
    border: 1px solid #d9e2ec;
    border-radius: 4px;
    display: block;
  }
  .photo-name { font-weight: 700; font-size: 13px; overflow-wrap: anywhere; color: #1f2933; }
  .detail-name { font-weight: 700; color: #000; }
  @media (max-width: 720px) {
    main { padding-left: 14px; padding-right: 14px; }
    .catalog-table { table-layout: auto; }
    .sample img { width: 110px; height: 92px; }
    th, td { font-size: 13px; padding: 8px; }
    .category-cell { font-size: 15px; }
  }
</style>
</head>
<body>
<header>
  <h1>Boston E Branch Sign Catalog</h1>
  <p class="subtitle">Catalog grouped by transit category, with detail-level counts and one sample image for each sign type. Built from sign_catalog_table.xlsm.</p>
</header>
<main>
"""
    ]

    parts.append('<div class="overview">')
    parts.append(
        f'<div class="metric"><span class="count">{total}</span><span class="label">Total signs</span></div>'
    )
    for category in CATEGORY_ORDER:
        parts.append(
            f'<div class="metric"><span class="count">{category_counts.get(category, 0)}</span><span class="label">{esc(CATEGORY_LABELS[category])}</span></div>'
        )
    parts.append("</div>")

    for category in CATEGORY_ORDER:
        detail_groups = groups.get(category, {})
        if not detail_groups:
            continue
        category_count = sum(len(items) for items in detail_groups.values())
        accent = CATEGORY_COLORS[category]
        label = CATEGORY_LABELS[category]
        parts.append(f'<section style="--accent:{accent}">')
        parts.append(
            f'<div class="section-head"><h2>{esc(label)}</h2><span class="badge">{category_count} signs</span></div>'
        )
        parts.append(
            '<table class="catalog-table"><colgroup><col style="width:17%"><col style="width:9%"><col style="width:26%"><col style="width:8%"><col style="width:40%"></colgroup><thead><tr><th>Transit_category</th><th>Count</th><th>Details</th><th>Count</th><th>Sample</th></tr></thead><tbody>'
        )
        sorted_details = sorted(detail_groups)
        for index, detail in enumerate(sorted_details):
            items = detail_groups[detail]
            sample = sample_for(items)
            sample_img = sample["photo_url"] or ""
            category_cells = ""
            if index == 0:
                category_cells = (
                    f'<td class="category-cell" rowspan="{len(sorted_details)}">{esc(label)}</td>'
                    f'<td class="count-cell" rowspan="{len(sorted_details)}">{category_count}</td>'
                )
            parts.append("<tr>")
            parts.append(category_cells)
            parts.append(f'<td class="detail-name">{esc(detail)}</td>')
            parts.append(f'<td class="count-cell">{len(items)}</td>')
            parts.append(
                f'''<td>
  <div class="sample">
    <img src="{esc(sample_img)}" alt="{esc(sample["photo_name"])}">
    <div class="photo-name">{esc(sample["photo_name"])}</div>
  </div>
</td>'''
            )
            parts.append("</tr>")
        parts.append("</tbody></table>")
        parts.append("</section>")

    parts.append("</main></body></html>")
    out_html.write_text("\n".join(parts), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build the embeddable StoryMap sign catalog page from sign_catalog_table.xlsm/csv."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE_TABLE,
        help=f"Input sign catalog table (.xlsm, .xlsx, or .csv). Default: {DEFAULT_SOURCE_TABLE}",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=DEFAULT_OUT_DIR,
        help=f"Output folder for sign_catalog.html and storymap_catalog_summary.csv. Default: {DEFAULT_OUT_DIR}",
    )
    parser.add_argument(
        "--out-csv",
        type=Path,
        default=DEFAULT_OUT_CSV,
        help=f"CSV copy to write from the source table. Default: {DEFAULT_OUT_CSV}",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_table = args.source.resolve()
    out_dir = args.out_dir.resolve()
    out_csv = args.out_csv.resolve()
    out_html = out_dir / "sign_catalog.html"
    out_summary = out_dir / "storymap_catalog_summary.csv"

    out_dir.mkdir(parents=True, exist_ok=True)
    rows = load_rows(source_table)
    write_csv(rows, out_csv)
    write_summary(rows, out_summary)
    write_html(rows, out_html)
    print(f"Read {source_table}")
    print(f"Wrote {out_csv}")
    print(f"Wrote {out_html}")
    print(f"Wrote {out_summary}")
    print(f"Rows: {len(rows)}")
    for category in CATEGORY_ORDER:
        count = Counter(row["Transit_category"] for row in rows).get(category, 0)
        print(f"{category}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
