# cme-lab1

Converts CME Group's daily volume/open-interest Excel reports into a clean, queryable dimensional model (star schema) as CSVs.

## Why

CME Group publishes daily "Volume and Open Interest by Product" reports as Excel files, each with inconsistent header positions, repeated/blank rows, and occasional missing columns. This script parses a folder of those reports into four normalized tables ready to load into a database or query directly.

## Output

- `dim_exchange.csv` — exchange reference table
- `dim_instrument.csv` — unique instrument definitions (asset class, exchange, commodity/future-option indicator, product)
- `dim_source_file.csv` — one row per ingested report, with its parsed trade date
- `fact_daily_volume.csv` — daily volume/open-interest figures, keyed to the dimension tables above

## How it works

- Locates the header row dynamically per file (handles files where it isn't always on the same row)
- Validates required columns are present; fills optional ones (`mtd_adv`, `open_interest`) with `None` when missing instead of failing the whole file
- Parses the trade date from the sheet content first, falling back to the filename pattern (`daily_volume_YYYYMMDD.xlsx`)
- Cleans numeric fields (strips commas, coerces safely) and normalizes column names
- Assigns stable surrogate keys to exchanges and instruments as it processes files, so the same instrument across multiple days maps to one `instrument_id`

## Run

```bash
python csv-converter.py
```

Reads `.xlsx` files from `data_raw/`, writes the four CSVs to `out_csv/`.

## Data source

Input files are CME Group's publicly available daily volume reports. (An earlier version of this pipeline used a Selenium script to pull the reports automatically; this repo currently expects the `.xlsx` files to already be in `data_raw/`.)

## Stack

Python, `pandas`, `openpyxl` (via pandas' Excel reader), standard library `csv`/`re`.
