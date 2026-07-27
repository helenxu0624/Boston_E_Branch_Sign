# Boston E Branch Sign Catalog

This folder contains the data source and generated page for the embeddable sign catalog.

## Files

- `sign_catalog_table.xlsm`: source table. Edit this file when sign categories, details, URLs, or coordinates change.
- `sign_catalog_table.csv`: generated CSV copy of the Excel table.
- `sign_catalog.html`: generated catalog page for ArcGIS StoryMaps embed.
- `storymap_catalog_summary.csv`: generated summary grouped by `Transit_category` and `details`.

## Update The HTML After Editing The Excel Table

From the project root, run:

```bash
python3 tools/build_storymap_catalog.py
```

The script reads:

```text
arcgis_ready/sign_catalog/sign_catalog_table.xlsm
```

and rewrites:

```text
arcgis_ready/sign_catalog/sign_catalog_table.csv
arcgis_ready/sign_catalog/sign_catalog.html
arcgis_ready/sign_catalog/storymap_catalog_summary.csv
```

You can also pass a custom source table or output folder:

```bash
python3 tools/build_storymap_catalog.py --source arcgis_ready/sign_catalog/sign_catalog_table.xlsm --out-dir arcgis_ready/sign_catalog
```

Required CSV columns:

```text
Pole_id, Sign_id, Photoname, Transit_category, photo_url, details
```

Other columns can stay in the CSV and will be preserved as source data, but the StoryMap table uses `Transit_category`, `details`, `Photoname`, and `photo_url`.
