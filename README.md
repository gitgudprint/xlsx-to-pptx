# PPT Creator

Generates one HR Dashboard PPTX per region from a shared `template.pptx` and a set of
source `.xlsx` workbooks (attrition, headcount, training, fraud, etc.). Each region's
deck is stamped out from the same template, with charts, tables, highlight indicators,
and text filled in with that region's real data.

## Requirements

- Python 3.9+
- `pip install -r requirements.txt`

## Usage

```bash
python main.py                     # generate all 12 regions
python main.py "Jawa Tengah"       # generate one specific region
python main.py --dry-run           # load data only, skip PPTX generation
```

Output files are written to `output/`.

## Data

This repo does **not** include any source data. `main.py` expects an input folder
(configured via `INPUT_DIR` in [`src/config.py`](src/config.py)) containing:

- `template.pptx` — the master slide deck template
- `a.` through `g.` prefixed `.xlsx` files — one per data domain (see `XLSX_FILES`
  in `src/config.py` for the exact expected filenames)

These files are confidential and are git-ignored — set `INPUT_DIR` to point at
wherever they live locally before running.

## Project layout

| File | Role |
|---|---|
| [`main.py`](main.py) | Entry point — loads all data once, then generates a PPTX per region |
| [`src/config.py`](src/config.py) | File paths, region name mappings, template placeholder tokens |
| [`src/data_loader.py`](src/data_loader.py) | Reads every xlsx source, including a pivot-cache parser for sheets that are live Excel PivotTables |
| [`src/chart_data.py`](src/chart_data.py) | Shapes loaded data into `(categories, series)` per chart number |
| [`src/xml_updater.py`](src/xml_updater.py) | Low-level, byte-safe OOXML editing (chart caches, embedded workbooks, table cells, in-memory zip editing) |
| [`src/pptx_updater.py`](src/pptx_updater.py) | Per-region orchestration — fills every chart, table, and highlight indicator on each slide |

## Notes

- Several source sheets are live Excel PivotTables filtered to a single view; `data_loader.py`
  reads their underlying pivot **cache** instead, so every region's real data is computed
  correctly regardless of what the sheet happens to display when saved.
- Table-cell filling goes through `_rebuild_row_cells` in `pptx_updater.py`, which rebuilds
  rows by XML position rather than by text search-and-replace — needed because a cell's real
  value can coincidentally match its template placeholder (e.g. `0`), which breaks naive
  content-based replacement.
