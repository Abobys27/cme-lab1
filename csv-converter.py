import re
import csv
from pathlib import Path
from datetime import datetime
import pandas as pd

PRODUCT_SHEET = "CME Group Vol and OI by Product"

NUMERIC_COLS = [
    "cme_globex_volume",
    "pit_volume",
    "expit_volume",
    "otc_volume",
    "total_volume",
    "mtd_adv",
    "open_interest",
]

def clean_col(col: str) -> str:
    col = str(col).strip()
    col = col.replace("\n", " ")
    col = re.sub(r"[®*]+", "", col)          # remove CME® / ** / * noise
    col = re.sub(r"\s+", " ", col).strip()
    col = col.lower().replace(" ", "_")
    col = re.sub(r"[^a-z0-9_]", "", col)
    col = col.replace("cme_globex_volume", "cme_globex_volume")  # keep stable
    return col

def parse_trade_date_from_sheet(xlsx_path: Path) -> str | None:
    """
    Looks for 'Trade Date:  mm/dd/yyyy' in the first few rows of the product sheet.
    Returns ISO date string 'YYYY-MM-DD' or None.
    """
    try:
        raw = pd.read_excel(xlsx_path, sheet_name=PRODUCT_SHEET, header=None, nrows=6)
        # usually the "Trade Date:" is in row 1, col 0
        for cell in raw.iloc[:, 0].dropna().astype(str).tolist():
            m = re.search(r"Trade Date:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", cell)
            if m:
                dt = datetime.strptime(m.group(1), "%m/%d/%Y").date()
                return dt.isoformat()
    except Exception:
        pass
    return None

def parse_trade_date_from_filename(xlsx_path: Path) -> str | None:
    """
    Expected filename pattern: daily_volume_YYYYMMDD.xlsx
    """
    m = re.search(r"(\d{8})", xlsx_path.stem)
    if not m:
        return None
    try:
        dt = datetime.strptime(m.group(1), "%Y%m%d").date()
        return dt.isoformat()
    except ValueError:
        return None

def to_int_or_none(x):
    if pd.isna(x):
        return None
    # Some cells are floats that are really integers
    try:
        if isinstance(x, str):
            x = x.replace(",", "").strip()
            if x == "":
                return None
        v = float(x)
        return int(v)
    except Exception:
        return None

def to_float_or_none(x):
    if pd.isna(x):
        return None
    try:
        if isinstance(x, str):
            x = x.replace(",", "").strip()
            if x == "":
                return None
        return float(x)
    except Exception:
        return None

def read_product_sheet(xlsx_path: Path) -> pd.DataFrame:
    # try to locate row that contains the header names
    raw = pd.read_excel(xlsx_path, sheet_name=PRODUCT_SHEET, header=None)

    header_row_index = None
    for i in range(20):   # search first 20 rows
        row_vals = raw.iloc[i].astype(str).str.lower().tolist()
        if "product description" in row_vals or "product_description" in row_vals:
            header_row_index = i
            break

    if header_row_index is None:
        raise ValueError(f"{xlsx_path.name}: cannot find header row")

    df = pd.read_excel(xlsx_path, sheet_name=PRODUCT_SHEET, header=header_row_index)
    df.columns = [clean_col(c) for c in df.columns]

    needed = [
        "description",
        "exchange_name",
        "commodity_indicator",
        "product_description",
        "futureoption_indicator",
        "cme_globex_volume",
        "pit_volume",
        "expit_volume",
        "otc_volume",
        "total_volume",
        "mtd_adv",
        "open_interest",
    ]

    optional = {"mtd_adv", "open_interest"}

    missing = [c for c in needed if c not in df.columns]
    hard_missing = [c for c in missing if c not in optional]

    if hard_missing:
        raise ValueError(f"{xlsx_path.name}: missing REQUIRED columns: {hard_missing}")

    for col in missing:
        print(f"    Warning: missing optional column '{col}' in {xlsx_path.name}, filling with None")
        df[col] = None

    # Drop junk rows (sometimes Excel exports include repeated header lines or blanks)
    df = df.dropna(subset=["description", "exchange_name", "product_description"])

    # Clean strings
    for c in ["description", "exchange_name", "commodity_indicator", "product_description", "futureoption_indicator"]:
        df[c] = df[c].astype(str).str.strip()

    # Numeric cleanup
    for c in ["cme_globex_volume", "pit_volume", "expit_volume", "otc_volume", "total_volume", "open_interest"]:
        df[c] = df[c].map(to_int_or_none)
    df["mtd_adv"] = df["mtd_adv"].map(to_float_or_none)

    return df[needed]

def build_csvs(input_dir: str, out_dir: str):
    in_path = Path(input_dir)
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # Output files
    exchanges_csv = out_path / "dim_exchange.csv"
    instruments_csv = out_path / "dim_instrument.csv"
    sources_csv = out_path / "dim_source_file.csv"
    fact_csv = out_path / "fact_daily_volume.csv"

    # ID maps
    exchange_id = {}
    instrument_id = {}

    next_exchange_id = 1
    next_instrument_id = 1
    next_source_id = 1

    # Prepare writers
    fact_f = open(fact_csv, "w", newline="", encoding="utf-8")
    fact_w = csv.writer(fact_f)
    fact_w.writerow([
        "trade_date",
        "instrument_id",
        "source_id",
        "cme_globex_volume",
        "pit_volume",
        "expit_volume",
        "otc_volume",
        "total_volume",
        "mtd_adv",
        "open_interest",
    ])

    src_f = open(sources_csv, "w", newline="", encoding="utf-8")
    src_w = csv.writer(src_f)
    src_w.writerow(["source_id", "source_file", "trade_date"])

    # Process files
    files = sorted(in_path.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in {in_path}")

    for xlsx in files:
        trade_date = parse_trade_date_from_sheet(xlsx) or parse_trade_date_from_filename(xlsx)
        if not trade_date:
            print(f"Skipping {xlsx.name}: cannot parse trade_date")
            continue

        source_id = next_source_id
        next_source_id += 1
        src_w.writerow([source_id, xlsx.name, trade_date])

        try:
            df = read_product_sheet(xlsx)
        except Exception as e:
            print(f"Skipping {xlsx.name}: {e}")
            continue

        # Write fact rows with dim mapping
        for row in df.itertuples(index=False):
            asset_class = row.description
            exch = row.exchange_name
            comm_ind = row.commodity_indicator
            prod = row.product_description
            fo = row.futureoption_indicator

            if exch not in exchange_id:
                exchange_id[exch] = next_exchange_id
                next_exchange_id += 1

            exch_id = exchange_id[exch]

            # instrument key = “what defines a unique traded line”
            key = (asset_class, exch_id, comm_ind, prod, fo)
            if key not in instrument_id:
                instrument_id[key] = next_instrument_id
                next_instrument_id += 1

            inst_id = instrument_id[key]

            fact_w.writerow([
                trade_date,
                inst_id,
                source_id,
                row.cme_globex_volume,
                row.pit_volume,
                row.expit_volume,
                row.otc_volume,
                row.total_volume,
                row.mtd_adv,
                row.open_interest,
            ])

        print(f"Ingested {xlsx.name}: {len(df)} rows")

    fact_f.close()
    src_f.close()

    # Write dimension tables at the end (from maps)
    with open(exchanges_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["exchange_id", "exchange_name"])
        for name, eid in sorted(exchange_id.items(), key=lambda x: x[1]):
            w.writerow([eid, name])

    with open(instruments_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "instrument_id",
            "asset_class",
            "exchange_id",
            "commodity_indicator",
            "product_description",
            "futureoption_indicator",
        ])
        # invert map for stable ordering
        inv = sorted(((iid, k) for k, iid in instrument_id.items()), key=lambda x: x[0])
        for iid, (asset_class, exch_id, comm_ind, prod, fo) in inv:
            w.writerow([iid, asset_class, exch_id, comm_ind, prod, fo])

    print("\nDone. CSV outputs:")
    print(exchanges_csv)
    print(instruments_csv)
    print(sources_csv)
    print(fact_csv)

if __name__ == "__main__":
    # Example:
    # build_csvs(input_dir="./cme_xlsx", out_dir="./out_csv")
    build_csvs(input_dir="./data_raw", out_dir="./out_csv")