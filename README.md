# historical-mf

Builds a history of Indian mutual fund portfolio holdings from two sources:

- **API pipeline (`api/`)**: pulls monthly portfolios from the AdvisorKhoj API, cleans them and adds symbol and market cap from the AMFI list. It writes one JSON file per category.
- **Excel pipeline (`excel/`)**: parses AMC monthly disclosure spreadsheets into a standard format (holdings, scheme snapshot and AMC-reported totals).
- **Neo4j loader (`ingestion-pipeline/`)**: loads the API pipeline's JSON output into a Neo4j graph.

```
historical-mf/
├── .env                          # API + Neo4j credentials (not committed)
├── .env.example                  # template for .env
├── requirements.txt              # all Python dependencies
├── venv/                         # project virtual environment
├── api/
│   ├── parser.py                 # API pipeline entry point
│   ├── api_client.py             # AdvisorKhoj HTTP client
│   ├── input/
│   │   ├── AMFI_data_JAN_JUL_2026.xlsx   # ISIN -> symbol / market cap (sheet "FINAL")
│   │   └── data.json             # sample raw API response (reference only)
│   └── output/                   # <Category>_<Mon>_<YYYY>_<Mon>_<YYYY>.json
├── excel/
│   ├── monthly_disclosures/<YEAR>/<MON>/<AMC>/<Fund>.xlsx|.xls|.xlsb
│   ├── market_cap_mapping/
│   │   ├── <YEAR>/market_cap_{June|December}_<YEAR>.xlsx
│   │   ├── isin_changes_*.csv        # old <-> new ISIN pairs (corporate actions)
│   │   └── manual_overrides.csv      # equity no AMFI list covers yet
│   ├── scheme_master_mapping/
│   │   ├── apply_scheme_mapping.py
│   │   ├── scheme_master.xlsx
│   │   └── scheme_name_mapping.csv
│   └── standardized_format/
│       ├── mf_portfolio_parser.py    # Excel pipeline entry point
│       └── output/
└── ingestion-pipeline/
    └── neo4j_data.py             # loads api/output/*.json into Neo4j
```

## Setup

You need Python 3.11. Run all commands from the repo root.

```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Copy `.env.example` to `.env` in the repo root and fill in the values. `load_dotenv()` searches upward, so both `api/` and `ingestion-pipeline/` pick it up.

```powershell
copy .env.example .env
```

```ini
# AdvisorKhoj API
API_KEY=...
API_URL=https://mfapi.advisorkhoj.com/getMutualfundHistoricalPortfolio   # optional, this is the default

# Neo4j (only needed for ingestion-pipeline/neo4j_data.py)
NEO4J_URI=neo4j+s://<instance>.databases.neo4j.io
NEO4J_USERNAME=...
NEO4J_PASSWORD=...
NEO4J_DATABASE=neo4j
```

### What is not in the repo

`.gitignore` keeps credentials and data out of version control, so a fresh clone does not include:

| Path | Why | How to get it |
|---|---|---|
| `.env` | Credentials | Copy from `.env.example` |
| `api/input/` | Licensed AMFI list and sample API data | Obtain separately, place as shown above |
| `api/output/` | Licensed API data | Run the API pipeline |
| `excel/monthly_disclosures/` | Large raw AMC files | Download from AMC websites into `<YEAR>/<MON>/<AMC>/` |
| `excel/standardized_format/output/` | Generated | Run the Excel pipeline |
| `venv/` | Local environment | See setup above |

## API pipeline

```powershell
# All categories, every month from the start month up to the cutoff
python api\parser.py

# A single category (use the exact name the API returns)
python api\parser.py --category "Equity: Large Cap"

# Only the last 3 months, first 5 funds per category (handy for testing)
python api\parser.py --category "Equity: Large Cap" --months 3 --limit 5
```

| Option | Meaning |
|---|---|
| `--category NAME` | Process only this category. Without it, every category from `getAllSchemeCategories` is processed. |
| `--months N` | Keep only the last N months of the resolved range. |
| `--limit N` | Process only the first N funds in each category. |

The date range is set by constants at the top of [api/parser.py](api/parser.py):

- `DATA_START_YEAR`, `DATA_START_MONTH`: the first month to fetch (currently July 2025).
- `END_MONTH_LAG`: how many months to stay behind the current month. `2` means that in September the last month fetched is July.

**What it does, per fund and month:**

1. Calls `getMutualfundHistoricalPortfolio`. Months that fail are logged and skipped.
2. Removes scheme-level and raw fields (`value`, `rating`, `portfolio_date`, …) and cleans instrument names.
3. Classifies each equity row as Domestic, International or Derivatives, and moves REITs out of `Others`.
4. Looks up the ISIN, symbol and market cap in `api/input/AMFI_data_JAN_JUL_2026.xlsx`. The lookup tries the exact ISIN first, then the normalized company name, then the parent company name.
5. Gives special share classes a suffix on their ISIN: `-PS` for preference shares, `-PP` for partly paid, `-L` for locked-in and `-R` for rights.

The output is `api/output/<Category>_<StartMon>_<StartYear>_<EndMon>_<EndYear>.json`, shaped as `{year: {category: [{fund_name, scheme_amfi_code, <month>: [records]}]}}`.

> Update the `EXCEL_FILE` path in `api/parser.py` whenever you replace the AMFI list with a newer file.

## Neo4j loader

```powershell
python ingestion-pipeline\neo4j_data.py
```

| Option | Meaning |
|---|---|
| `--input PATH [PATH ...]` | JSON file(s) or folder(s) to load. Default `api/output/`. The Excel pipeline's `holdings.json` and `categories/<Category>.json` use the same schema, so they can be loaded directly. |
| `--category NAME [NAME ...]` | Only load these categories (case-insensitive), e.g. `--category Value`. |

```powershell
# Load only the Value category from the Excel pipeline
python ingestion-pipeline\neo4j_data.py --input excel\standardized_format\output\categories\Value.json
```

By default it reads every `*.json` in `api/output/` except `clean_data.json`, creates the constraints, and merges these nodes: `FUND`, `MONTHLY_SNAPSHOT`, `ASSET_CLASS`, `EQUITY_INSTRUMENT` and `INSTRUMENT`. Snapshots are linked with `LATEST_MONTHLY_SNAPSHOT` and `PREVIOUS_MONTHLY_SNAPSHOT`. Because every write is a `MERGE`, you can safely run it again.

## Excel pipeline

```powershell
# 1. Parse all disclosures -> holdings.csv / holdings.json / scheme_snapshot.csv / amc_reported_totals.csv
#    (defaults: excel\monthly_disclosures -> excel\standardized_format\output; works from any directory)
python excel\standardized_format\mf_portfolio_parser.py

#    only the Value category -> holdings_Value.json / .csv, scheme_snapshot_Value.csv, ...
python excel\standardized_format\mf_portfolio_parser.py --category Value

#    or with explicit paths:
python excel\standardized_format\mf_portfolio_parser.py <disclosures_root> <out_dir>

# 2. Apply canonical scheme names + AMC scheme codes -> holdings_mapped.csv
python excel\scheme_master_mapping\apply_scheme_mapping.py excel\standardized_format\output\holdings.csv
```

- Optional filters: `--category NAME` (as returned by `get_fund_category`, e.g. `Value`), `--year 2024`, `--month APR`, `--amc <text in AMC folder name>`. With `--category`, outputs are written as `holdings_<Category>.json` / `.csv` etc. so the full-run files are not overwritten.
- Every run also writes `output/categories/<Category>.json` (e.g. `Value.json`, `Large_Cap.json`), one file per category in the same schema, for loading into Neo4j category by category. To build these from an existing `holdings.json` without re-parsing: `python excel\standardized_format\mf_portfolio_parser.py --split-json excel\standardized_format\output\holdings.json`.
- The parser's arguments are `[disclosures_root] [out_dir]`. They default to `excel/monthly_disclosures` and `excel/standardized_format/output`, based on where the script is, not the current directory. It picks up files at `<root>/<YEAR>/<MON>/<AMC>/*.xls*`, where `<MON>` is the three-letter month, for example `APR`.
- It chooses the Excel reader from each file's contents, not its extension. Real legacy `.xls` files use `xlrd` and `.xlsb` files use `pyxlsb`. Some files named `.xls` are actually xlsx inside.
- Market cap comes from `excel/market_cap_mapping/<YEAR>/`. Disclosures for January to June use `market_cap_June_<YEAR>.xlsx`, and July to December use `market_cap_December_<YEAR>.xlsx`. If the file is missing, that disclosure is skipped with an error. For example, July to December 2026 needs `market_cap_December_2026.xlsx`.
- ISIN changes from corporate actions (face-value splits and similar) are read from every `excel/market_cap_mapping/isin_changes_*.csv` (columns `old`, `new`, …). If an ISIN isn't in the AMFI list for the period, its aliases are tried, so a holding reported under an old or new ISIN still gets a market cap.
- `excel/market_cap_mapping/manual_overrides.csv` (`company_name, isin, symbol, market_cap_category, note`) covers equity that is printed without an ISIN and isn't in any AMFI list yet, such as a newly demerged company. Rows are matched on the normalized company name. If you fill in only the `isin`, the symbol and market cap come from the AMFI list.
- Special equity instruments get an ISIN suffix based on their `asset_sub_type`, the same way the API pipeline does: `-PS` for Preference Shares, `-PP` for Partly Paid Shares and `-R` for Rights Entitlements. Market cap is still looked up with the original ISIN.
- Files that fail to parse are listed as `SKIPPED FILE` in the log. The rest of the run carries on.
- ISINs from any country are accepted (`^[A-Z]{2}[A-Z0-9]{9}[0-9]$`; non-`IN` codes must pass the ISO check digit). Overseas holdings (US…, GB…, IE…, KY…) are classified as Equity / Foreign Equity (`equityType` = International Equity) and get no market cap.
- There is no `UNCLASSIFIED` asset class. Rows without an ISIN whose value equals the sum of the instruments directly below them (ICICI `Equity shares`, `Foreign Securities/Overseas ETFs`) or above them (unlabelled subtotals) go to `amc_reported_totals.csv` as `(aggregate row: …)`, not to holdings. Anything left that matches no rule is dropped and listed in `unclassified_rows.csv` for review.
- Excel lock files (`~$*.xlsx`) in the month folders are ignored.
- `apply_scheme_mapping.py <holdings.csv> [out.csv]` writes `<name>_mapped.csv` by default. It also prints any scheme titles it couldn't map.

## Adding new data

| Task | What to do |
|---|---|
| New AMC disclosure month | Add files under `excel/monthly_disclosures/<YEAR>/<MON>/<AMC>/` and rerun the Excel pipeline. |
| New AMFI half-yearly market cap list | Add `market_cap_{June\|December}_<YEAR>.xlsx` to `excel/market_cap_mapping/<YEAR>/`. For the API pipeline, also update `EXCEL_FILE`. |
| Extend the API date range | Change `DATA_START_*` or `END_MONTH_LAG` in `api/parser.py`. |
| New scheme name variant | Add a row to `excel/scheme_master_mapping/scheme_name_mapping.csv`. |
| A company's ISIN changed | Add an `old,new` row to an `excel/market_cap_mapping/isin_changes_*.csv` file. |
| Holding with no ISIN and no AMFI entry | Add a row to `excel/market_cap_mapping/manual_overrides.csv`. Fill in the ISIN once it is known. |
