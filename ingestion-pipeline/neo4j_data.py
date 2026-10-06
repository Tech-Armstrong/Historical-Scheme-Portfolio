import argparse
import json
import os
from datetime import datetime
from collections import defaultdict

from dotenv import load_dotenv
from neo4j import GraphDatabase, RoutingControl

load_dotenv()

URI = os.getenv("NEO4J_URI")
USERNAME = os.getenv("NEO4J_USERNAME")
PASSWORD = os.getenv("NEO4J_PASSWORD")
DATABASE = os.getenv("NEO4J_DATABASE", "neo4j")

DEFAULT_INPUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "api", "output"
)

arg_parser = argparse.ArgumentParser(
    description="Load API-schema portfolio JSON ({year: {category: [funds]}}) into Neo4j."
)
arg_parser.add_argument(
    "--input", nargs="+", default=[DEFAULT_INPUT],
    help="JSON file(s) or folder(s) of JSON files (default: api/output)",
)
arg_parser.add_argument(
    "--category", nargs="+",
    help='Only load these categories, e.g. --category Value',
)
args = arg_parser.parse_args()

json_files = []

for input_path in args.input:
    if os.path.isdir(input_path):
        json_files.extend(
            os.path.join(input_path, file_name)
            for file_name in os.listdir(input_path)
            if (
                file_name.lower().endswith(".json")
                and file_name.lower() != "clean_data.json"
            )
        )
    elif os.path.isfile(input_path):
        json_files.append(input_path)
    else:
        print("Input not found:", input_path)

if not json_files:
    print("No JSON files found in:", ", ".join(args.input))
    exit()

wanted_categories = (
    {name.strip().lower() for name in args.category}
    if args.category else None
)

all_data = []

for json_file in sorted(json_files):
    print("Loading:", os.path.basename(json_file))

    with open(json_file, "r", encoding="utf-8") as file:
        file_data = json.load(file)

    all_data.append(file_data)

print(f"Loaded {len(all_data)} JSON files successfully.")

driver = GraphDatabase.driver(URI, auth=(USERNAME, PASSWORD))
driver.verify_connectivity()

print("Connected to Neo4j successfully!")

constraints = [
    "CREATE CONSTRAINT fund_amfi_code_unique IF NOT EXISTS FOR (f:FUND) REQUIRE f.amfiCode IS UNIQUE",
    "CREATE CONSTRAINT snapshot_fund_year_month_unique IF NOT EXISTS FOR (s:MONTHLY_SNAPSHOT) REQUIRE (s.amfiCode, s.year, s.month) IS UNIQUE",
    "CREATE CONSTRAINT equity_instrument_isin_unique IF NOT EXISTS FOR (i:EQUITY_INSTRUMENT) REQUIRE i.ISIN IS UNIQUE",
    "CREATE CONSTRAINT asset_class_id_unique IF NOT EXISTS FOR (a:ASSET_CLASS) REQUIRE a.assetClassId IS UNIQUE",
]

for statement in constraints:
    driver.execute_query(statement, database_=DATABASE)

print("Constraints and indexes verified.")

def first_value(dictoionary, *keys):
    for key in keys:
        value = dictoionary.get(key)
        if value is not None and value != "":
            return value
    return None
def clean_number(value):
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        value = value.replace(",", "").replace("₹", "").strip()
        try:
            number = float(value)
            return int(number) if number.is_integer() else number
        except ValueError:
            return 0
    return 0


def parse_date_safe(date_string):
    if not date_string:
        return None
    date_string = str(date_string).strip()
    formats = ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    print(f"  WARNING: could not parse date '{date_string}', skipping this record")
    return None


def get_asset_class(item):
    asset_class = first_value(
        item, "asset_class", "assetClass", "asset_class_name", "assetClassName"
    )
    if not asset_class:
        return None
    return str(asset_class).strip()


def get_total_value(item):
    value = first_value(
        item,
        "market_value", "marketValue", "value",
        "total_value", "totalValue", "valuation", "amount",
    )
    return clean_number(value)

ALL_ASSET_CLASS_NAMES = {
    "equity": "Equity",
    "international": "Equity",
    "derivatives": "Equity",
    "debt": "Debt",
    "reit": "REITs & InvITs",
    "reits & invits": "REITs & InvITs",
    "commodity": "Commodity",
    "cash & cash equivalents": "Cash and Cash Equivalents",
    "cash and cash equivalents": "Cash and Cash Equivalents",
    "net receivables / (payables)": "Net Receivable/Payable",
    "mf": "MF and ETF",
    "fund units": "MF and ETF",
    "mutual funds": "MF and ETF",
    "mf and etf": "MF and ETF",
}


monthly_data = defaultdict(list)

for file_data in all_data:
    for year, categories in file_data.items():
        for category_name, funds in categories.items():
            if (
                wanted_categories
                and category_name.strip().lower() not in wanted_categories
            ):
                continue

            for fund in funds:

                fund_name = fund.get("fund_name")
                amfi_code = fund.get("scheme_amfi_code")

                if not amfi_code:
                    print(
                        f"WARNING: Missing AMFI code for fund: {fund_name}"
                    )
                    continue

                for month, month_records in fund.items():
                    if month in ["fund_name", "scheme_amfi_code"]:
                        continue
                    if not isinstance(month_records, list):
                        continue
                    try:
                        month_number = datetime.strptime(
                            month,
                            "%B"
                        ).month
                    except ValueError:
                        print(
                            f"WARNING: Unknown month '{month}' "
                            f"for fund '{fund_name}'"
                        )
                        continue

                    date_object = datetime(
                        int(year),
                        month_number,
                        1
                    )

                    month_key = (
                        str(amfi_code),
                        date_object.strftime("%Y-%m")
                    )

                    for item in month_records:

                        monthly_data[month_key].append(
                            {
                                "fund_name": fund_name,
                                "amfi_code": str(amfi_code),
                                "category": category_name,
                                "date": date_object,
                                "item": item
                            }
                        )

sorted_months = sorted(monthly_data.keys())

sorted_months = sorted(monthly_data.keys())

print()
print("Fund/month snapshots found:")

for amfi_code, month_key in sorted_months:
    print(
        " ",
        amfi_code,
        "->",
        month_key,
        "->",
        len(monthly_data[(amfi_code, month_key)]),
        "portfolio records"
    )

query = """

MERGE (fund:FUND {
    amfiCode: $amfiCode
})
SET fund.fundName = $fundName,
    fund.category = $category

MERGE (snapshot:MONTHLY_SNAPSHOT {
    amfiCode: $amfiCode,
    year: $year,
    month: $month
})
SET snapshot.amfiCode = $amfiCode,
    snapshot.reportedDate = date($reportedDate),
    snapshot.fundManager = $fundManager,
    snapshot.AUM = $AUM,
    snapshot.marketCapBreakdown = $marketCapBreakdown,
    snapshot.assetClassBreakdown = $assetClassBreakdown,
    snapshot.sectorBreakdown = $sectorBreakdown
 
WITH fund, snapshot

OPTIONAL MATCH (fund)-[latestRel:LATEST_MONTHLY_SNAPSHOT]->(currentLatest)

FOREACH (_ IN CASE
    WHEN currentLatest IS NOT NULL
         AND currentLatest <> snapshot
         AND snapshot.reportedDate > currentLatest.reportedDate
    THEN [1]
    ELSE []
END |
    MERGE (snapshot)-[:PREVIOUS_MONTHLY_SNAPSHOT]->(currentLatest)
    DELETE latestRel
    MERGE (fund)-[:LATEST_MONTHLY_SNAPSHOT]->(snapshot)
)

FOREACH (_ IN CASE
    WHEN currentLatest IS NULL
    THEN [1]
    ELSE []
END |
    MERGE (fund)-[:LATEST_MONTHLY_SNAPSHOT]->(snapshot)
)

WITH snapshot, $instruments AS instruments

UNWIND instruments AS item

MERGE (assetClass:ASSET_CLASS {
    assetClassId:
        snapshot.amfiCode + "-" +
        toString(snapshot.year) + "-" +
        snapshot.month + "-" +
        item.assetClass
})

SET assetClass.name = item.assetClass,
    assetClass.snapshotYear = snapshot.year,
    assetClass.snapshotMonth = snapshot.month

MERGE (snapshot)-[:HAS_ASSET_CLASS]->(assetClass)

FOREACH (_ IN CASE
    WHEN item.assetClass = "Equity"
    THEN [1]
    ELSE []
END |
    MERGE (instrument:EQUITY_INSTRUMENT {
        ISIN: item.isin
    })
    SET instrument.name = item.instrument,
        instrument.equityType = item.equityType,
        instrument.sector = item.sector,
        instrument.industry = item.industry,
        instrument.symbol = item.symbol,
        instrument.marketCap = item.marketCap

    MERGE (assetClass)-[h:HAS_INSTRUMENT]->(instrument)


    SET h.holdings = item.holdings
)

FOREACH (_ IN CASE
    WHEN item.assetClass = "Cash and Cash Equivalents"
    THEN [1]
    ELSE []
END |
    MERGE (instrument:INSTRUMENT {
        ISIN: coalesce(item.isin, item.instrument)
    })
    SET instrument.name = item.instrument

    MERGE (assetClass)-[h:HAS_INSTRUMENT]->(instrument)

    SET h.holdings = item.holdings
)

FOREACH (_ IN CASE
    WHEN item.assetClass = "MF and ETF"
    THEN [1]
    ELSE []
END |
    MERGE (instrument:INSTRUMENT {
        ISIN: coalesce(item.isin, item.instrument)
    })
    SET instrument.name = item.instrument

    MERGE (assetClass)-[h:HAS_INSTRUMENT]->(instrument)
    SET h.holdings = item.holdings
)

FOREACH (_ IN CASE
    WHEN item.assetClass IN ["Debt", "REITs & InvITs"]
    THEN [1]
    ELSE []
END |
    MERGE (instrument:INSTRUMENT {
        ISIN: coalesce(item.isin, item.instrument)
    })
    SET instrument.name = item.instrument,
        instrument.sector = item.sector,
        instrument.industry = item.industry,
        instrument.symbol = item.symbol,
        instrument.marketCap = item.marketCap

    MERGE (assetClass)-[h:HAS_INSTRUMENT]->(instrument)

    SET h.holdings = item.holdings
)

RETURN snapshot.reportedDate AS processedDate

"""

for amfi_code, month_key in sorted_months:

    entries = monthly_data[(amfi_code, month_key)]

    print()
    print("-" * 60)
    print("Processing:", month_key)
    print("Records:", len(entries))
    print("-" * 60)

    dates_in_group = [entry["date"] for entry in entries]

    canonical_date = max(dates_in_group)

    reported_date_str = canonical_date.strftime("%Y-%m-%d")

    year = canonical_date.year
    month_num = canonical_date.month
    month_name = canonical_date.strftime("%B")

    fund_name = entries[0]["fund_name"]
    category = entries[0]["category"]

    fund_manager = None
    AUM = None
    market_cap_breakdown = None
    asset_class_breakdown = None
    sector_breakdown = None

    instruments = []

    for entry in entries:
        item = entry["item"]
        raw_asset_class = get_asset_class(item)
        normalized = raw_asset_class.lower().strip() if raw_asset_class else ""
        display_name = ALL_ASSET_CLASS_NAMES.get(normalized)

        if display_name is None:
            print("WARNING: unrecognized asset_class, skipping record:", raw_asset_class)
            continue

        value = get_total_value(item)
        holdings_pct = clean_number(first_value(item, "holdings"))

        if display_name not in ["Equity", "Debt", "REITs & InvITs", "Cash and Cash Equivalents", "MF and ETF"]:
            continue

        raw_instrument = first_value(item, "instrument", "name")

        instruments.append({
        "assetClass": display_name,
        "equityType": first_value(item, "equityType", "equity_type"),
        "isin": first_value(item, "isin", "ISIN"),
        "instrument": raw_instrument,
        "sector": first_value(item, "sector"),
        "industry": first_value(item, "industry"),
        "symbol": first_value(item, "symbol"),
        "marketCap": first_value(item, "marketCap", "market_cap"),
        "holdings": holdings_pct,
    })

    unique_instruments = {}

    for instrument in instruments:

        isin = instrument.get("isin")
        key = (
            isin
            or instrument.get("symbol")
            or instrument.get("instrument")
        )

        if key in unique_instruments:
            unique_instruments[key]["holdings"] += instrument["holdings"]
        elif key:
            unique_instruments[key] = instrument
        else:
            print(
                "  WARNING: dropping instrument with no "
                "ISIN/symbol/name:",
                instrument
            )

    instruments = list(unique_instruments.values())
    instruments = [
        instrument
        for instrument in instruments
        if not (
            instrument.get("assetClass") == "Equity"
            and not instrument.get("isin")
        )
    ]



    class_totals = [
    {
        "name": name,
    }
    for name in dict.fromkeys(ALL_ASSET_CLASS_NAMES.values())
]

    driver.execute_query(
        query,
        amfiCode=amfi_code,
        fundName=fund_name,
        category=category,

        year=year,
        month=month_name,
        reportedDate=reported_date_str,

        fundManager=fund_manager,
        AUM=clean_number(AUM),
        marketCapBreakdown=market_cap_breakdown,
        assetClassBreakdown=asset_class_breakdown,
        sectorBreakdown=sector_breakdown,

        classTotals=class_totals,
        instruments=instruments,

        database_=DATABASE,
        routing_=RoutingControl.WRITE,
    )

    print("Snapshot processed:", month_name, year)
    print("Equity instruments:", len(instruments))

driver.close()

print()
print("-" * 60)
print("GRAPH CREATED SUCCESSFULLY!")
print("-" * 60)

