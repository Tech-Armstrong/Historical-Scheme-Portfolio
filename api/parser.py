import argparse
import json
import re
from pathlib import Path

import openpyxl

from api_client import (
    get_all_funds,
    get_categories,
    get_funds,
    get_portfolio,
    resolve_year_months,
    MONTH_NAME_MAP,
)

DATA_START_YEAR = 2025
DATA_START_MONTH = 7
END_MONTH_LAG = 2

API_DIR = Path(__file__).resolve().parent

EXCEL_FILE = API_DIR / "input" / "AMFI_data_JAN_JUL_2026.xlsx"
OUTPUT_DIR = API_DIR / "output"

SHEET_NAME = "FINAL"

COL_ISIN = "ISIN Code"
COL_BSE = "BSE Symbol"
COL_NSE = "NSE Symbol"
COL_MCAP = "Market Cap"
COL_NAME = "Company Name"

EXCLUDE_KEYS = {
    "scheme_amfi_common",
    "scheme_amfi_code",
    "scheme_name",
    "value",
    "rating",
    "rating_eq",
    "portfolio_date",
}

SCHEME_KEYS = {
    "scheme_amfi_common",
    "scheme_amfi_code",
    "scheme_name",
}

# RENAME_KEYS = {
#     "portfolio_date": "reported_date",
# }

NAME_NOISE_RE = re.compile(
    r"\b(LIMITED|LTD\.?|PVT\.?|PRIVATE|CO\.?|COMPANY|NEW|THE|INC\.?)\b"
    r"|[^A-Z0-9]",
    flags=re.IGNORECASE,
)

FILENAME_SAFE_RE = re.compile(r"[^A-Za-z0-9_-]+")


def normalize_name(name):
    if not isinstance(name, str):
        return ""

    name = name.upper().strip()

    # Normalize common company-name variations
    name = re.sub(r"\bLIMITED\b", "LTD", name)
    name = re.sub(r"\bCOMPANY\b", "CO", name)

    # Remove punctuation / spaces / common legal suffixes
    name = NAME_NOISE_RE.sub("", name)

    return name.strip()

def normalize_parent_name(name):

    if not isinstance(name, str):
        return ""

    text = name.upper().strip()

    # brackets

    text = re.sub(
        r"\([^)]*\)",
        " ",
        text
    )

    #number removal
    text = re.sub(
        r"\b\d+(?:\.\d+)?\s*%.*$",
        "",
        text
    )

    #preference share

    text = re.sub(
        r"\s*[-]?\s*"
        r"(?:PREF(?:ERENCE)?(?:\s+SHARE)?|NCRPS)\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    #rights

    text = re.sub(
        r"\s*[-]?\s*RIGHTS?\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    #pp

    text = re.sub(
        r"\s*[-]?\s*PP\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    #locked-in and suffix

    text = re.sub(
        r"\s*[-]?\s*LOCKED\s*[-]?\s*IN\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    #company and limited

    text = re.sub(
        r"\bCOMPANY\b",
        "CO",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\bLIMITED\b",
        "LTD",
        text,
        flags=re.IGNORECASE
    )

    #punctuation

    text = re.sub(
        r"[^A-Z0-9]",
        "",
        text
    )

    return text

def clean_instrument(instrument):

    if not isinstance(instrument, str):
        return instrument

    raw = instrument.strip()

    if not raw:
        return raw

    # =========================================================
    # DETECT SPECIAL INSTRUMENT TYPES
    # =========================================================

    # Preference shares:
    # PREF
    # PREFERENCE
    # PREF SHARE
    # NCRPS
    is_pref = bool(
        re.search(
            r"\bPREF(?:ERENCE)?\b"
            r"|\bPREF(?:ERENCE)?\s+SHARE\b"
            r"|\bNCRPS\b",
            raw,
            flags=re.IGNORECASE
        )
    )

    # Rights
    is_rights = bool(
        re.search(
            r"\bRIGHTS?\b",
            raw,
            flags=re.IGNORECASE
        )
    )

    # PP
    is_pp = bool(
        re.search(
            r"\bPP\b",
            raw,
            flags=re.IGNORECASE
        )
    )

    # Locked-in
    is_locked_in = bool(
        re.search(
            r"\bLOCKED\s*[-]?\s*IN\b",
            raw,
            flags=re.IGNORECASE
        )
    )

    text = raw

    # =========================================================
    # REMOVE LEADING "EQ"
    #
    # Example:
    # EQ - MANIPAL HEALTH ENTERPRISES LTD
    # -> MANIPAL HEALTH ENTERPRISES LTD
    #
    # Also handles:
    # EQ MANIPAL HEALTH ENTERPRISES LTD
    # =========================================================
    text = re.sub(
        r"^\s*EQ\b\s*[-–—]?\s*",
        "",
        text,
        flags=re.IGNORECASE
    )

    # =========================================================
    # REMOVE TRAILING / SUFFIX "EQ"
    #
    # Example:
    # MCX INDIA LIMITED EQ NEW RS. 10/-
    # -> MCX INDIA LIMITED
    #
    # ASTER DM HEALTHCARE LIMITED EQ
    # -> ASTER DM HEALTHCARE LIMITED
    # =========================================================
    text = re.sub(
        r"\s+\bEQ\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    # =========================================================
    # REMOVE BRACKET CONTENT
    #
    # Example:
    # ABC LTD (01-Sep-2026)
    # -> ABC LTD
    # =========================================================
    text = re.sub(
        r"\([^)]*\)",
        " ",
        text
    )

    # =========================================================
    # REMOVE LEADING PERCENTAGE
    #
    # IMPORTANT:
    #
    # OLD:
    # 6% TVS MOTOR CO LTD NCRPS
    #
    # became EMPTY because we removed:
    # 6% + everything after it.
    #
    # NEW:
    # 6% TVS MOTOR CO LTD NCRPS
    # -> TVS MOTOR CO LTD NCRPS
    # =========================================================
    text = re.sub(
        r"^\s*\d+(?:\.\d+)?\s*%\s*",
        "",
        text
    )

    # =========================================================
    # REMOVE TRAILING PERCENTAGE INFORMATION
    #
    # Example:
    # ABC LTD 60%
    # ABC LTD 60% something
    # -> ABC LTD
    # =========================================================
    text = re.sub(
        r"\s+\d+(?:\.\d+)?\s*%.*$",
        "",
        text
    )

    # =========================================================
    # CONVERT LIMITED -> LTD
    # =========================================================
    text = re.sub(
        r"\bLIMITED\b",
        "LTD",
        text,
        flags=re.IGNORECASE
    )

    # =========================================================
    # NORMALIZE SPACES
    # =========================================================
    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip()

    # =========================================================
    # FIND LTD
    #
    # If the company has LTD, everything after LTD is treated
    # as security-type information.
    # =========================================================
    ltd_match = re.search(
        r"\bLTD\.?\b",
        text,
        flags=re.IGNORECASE
    )

    if ltd_match:

        base_name = text[
            :ltd_match.end()
        ].strip(" .-")

        # Preference Share
        if is_pref:
            return f"{base_name} - Pref Share"

        # Rights
        if is_rights:
            return f"{base_name} - Rights"

        # PP
        if is_pp:
            return f"{base_name} - PP"

        # Locked-in
        if is_locked_in:
            return base_name

        return base_name

    # =========================================================
    # NO LTD FOUND
    #
    # Remove special-security suffixes from the remaining text.
    # =========================================================

    text = re.sub(
        r"\bLOCKED\s*[-]?\s*IN\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\bRIGHTS?\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\b(?:PREF(?:ERENCE)?(?:\s+SHARE)?|NCRPS)\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\bPP\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip(" -.")

    return text

def get_special_instrument_suffix(instrument):

    if not isinstance(instrument, str):
        return None

    if re.search(
        r"\bPREF(?:ERENCE)?\b"
        r"|\bPREF(?:ERENCE)?\s+SHARE\b"
        r"|\bNCRPS\b",
        instrument,
        flags=re.IGNORECASE
    ):
        return "-PS"

    if re.search(
        r"\bPP\b",
        instrument,
        flags=re.IGNORECASE
    ):
        return "-PP"

    if re.search(
        r"\bLOCKED\s*[-]?\s*IN\b",
        instrument,
        flags=re.IGNORECASE
    ):
        return "-L"

    if re.search(
        r"\bRIGHTS?\b",
        instrument,
        flags=re.IGNORECASE
    ):
        return "-R"

    return None

def clean_reit_instrument(instrument):

    if not isinstance(instrument, str):
        return instrument

    text = instrument.strip()

    text = re.sub(
        r"\s*[-–—]?\s*\bREITs?\b\s*[-–—]?\s*",
        " ",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(r"\s+", " ", text).strip()

    text = text.strip(" -–—")

    return text

def clean_symbol(value):
    if value is None:
        return None

    text = str(value).strip()
    if text == "" or text == "-":
        return None

    return text

def build_category_data(category_name, fund_results):
    short_category_name = clean_category_name(category_name)

    category_data = {}

    for fund in fund_results:
        if not fund:
            continue

        fund_name = fund["fund_name"]
        scheme_amfi_code = fund["scheme_amfi_code"]
        records_by_year = fund["records"]

        for year, months in records_by_year.items():

            if year not in category_data:
                category_data[year] = {}

            if short_category_name not in category_data[year]:
                category_data[year][short_category_name] = []

            fund_entry = {
                "fund_name": fund_name,
                "scheme_amfi_code": scheme_amfi_code,
            }

            for month, month_records in months.items():
                fund_entry[month] = month_records

            category_data[year][short_category_name].append(
                fund_entry
            )

    return category_data

def get_category_filename(category_name, year_months):

    if ":" in category_name:
        category_name = category_name.split(":", 1)[1].strip()

    safe_category = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        category_name
    ).strip("_")

    if not year_months:
        return f"{safe_category}.json"

    start_year, start_month = year_months[0]
    end_year, end_month = year_months[-1]

    filename = (
        f"{safe_category}_"
        f"{start_month.capitalize()}_{start_year}_"
        f"{end_month.capitalize()}_{end_year}.json"
    )

    return filename


def load_isin_symbol_map(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[SHEET_NAME]

    rows = ws.iter_rows(values_only=True)
    headers = next(rows)
    col_index = {name: i for i, name in enumerate(headers)}

    idx_isin = col_index[COL_ISIN]
    idx_bse = col_index[COL_BSE]
    idx_nse = col_index[COL_NSE]
    idx_mcap = col_index[COL_MCAP]
    idx_name = col_index.get(COL_NAME)

    if idx_name is None:
        print(
            f"  WARNING: column '{COL_NAME}' not found in sheet headers "
            f"{list(headers)}. Name-based fallback matching will be "
            f"disabled - update COL_NAME to the correct header."
        )

    isin_map = {}
    name_map = {}
    parent_name_map = {}
    duplicate_isins = []

    for row in rows:
        isin = row[idx_isin]
        company_name = row[idx_name] if idx_name is not None else None

        bse_symbol = clean_symbol(row[idx_bse])
        nse_symbol = clean_symbol(row[idx_nse])
        mcap = row[idx_mcap]
        symbol = nse_symbol if nse_symbol else bse_symbol

        isin_key = str(isin).strip() if isin else None

        record = {
            "symbol": symbol,
            "market_cap": mcap if mcap else None,
            "isin": isin_key,
            "company_name": company_name,
        }

        if isin_key:
            if isin_key in isin_map:
                duplicate_isins.append(isin_key)
                existing = isin_map[isin_key]
                if not existing.get("symbol") and symbol:
                    existing["symbol"] = symbol
                if not existing.get("market_cap") and mcap:
                    existing["market_cap"] = mcap
            else:
                isin_map[isin_key] = record

        if company_name:
            name_key = normalize_name(company_name)

            if name_key:
                name_map[name_key] = record

            parent_key = normalize_parent_name(company_name)

            if parent_key and parent_key not in parent_name_map:
                parent_name_map[parent_key] = record

    if duplicate_isins:
        print(
            f"  WARNING: {len(duplicate_isins)} duplicate ISIN rows found "
            f"in Excel (merged, not overwritten): {duplicate_isins[:10]}"
        )

    return isin_map, name_map, parent_name_map


def extract_scheme_info(portfolio_list):
    if not portfolio_list:
        return {}

    first_row = portfolio_list[0]
    return {key: first_row.get(key) for key in SCHEME_KEYS if key in first_row}

def clean_category_name(category):
    if ":" in category:
        return category.split(":", 1)[1].strip()

    return category.strip()

def clean_data(data, isin_map, name_map, parent_name_map):

    if isinstance(data, dict):

        cleaned = {}

        # =========================================================
        # GET ORIGINAL INSTRUMENT NAME
        # =========================================================
        raw_instrument_name = (
            data.get("instrument")
            or data.get("name")
        )

        # Detect special instrument type from the RAW API name.
        #
        # Pref Share -> -PS
        # PP         -> -PP
        # Locked-in  -> -L
        # Rights     -> -R
        special_suffix = get_special_instrument_suffix(
            raw_instrument_name
        )

        # =========================================================
        # BASIC CLEANING
        # =========================================================
        for key, value in data.items():

            if key.lower() in EXCLUDE_KEYS:
                continue

            # Clean instrument before doing company-name matching.
            if (
                key.lower() == "instrument"
                and isinstance(value, str)
            ):
                value = clean_instrument(value)

            cleaned[key] = clean_data(
                value,
                isin_map,
                name_map,
                parent_name_map
            )

        # =========================================================
        # EQUITY TYPE CLASSIFICATION
        # =========================================================

        if cleaned.get("asset_class") == "International":
            cleaned["asset_class"] = "Equity"
            cleaned["equityType"] = "International"

        elif (
            cleaned.get("asset_class") == "Equity"
            and cleaned.get("equityType") != "Derivatives"
        ):
            cleaned["equityType"] = "Domestic"

        # =========================================================
        # REIT HANDLING
        # =========================================================
        if (
            cleaned.get("asset_class") == "Others"
            and isinstance(cleaned.get("instrument"), str)
            and re.search(
                r"\bREITs?\b",
                cleaned["instrument"],
                flags=re.IGNORECASE
            )
        ):
            cleaned["asset_class"] = "REIT"

            cleaned["instrument"] = clean_reit_instrument(
                cleaned["instrument"]
            )

        # =========================================================
        # DERIVATIVES
        # =========================================================
        if (
            cleaned.get("equityType") == "Derivatives"
            and cleaned.get("isin") == "derivative"
        ):
            cleaned["symbol"] = None
            cleaned["market_cap"] = None
            cleaned["isin"] = None

            return cleaned

        # =========================================================
        # EQUITY
        # =========================================================
        if (
            cleaned.get("asset_class") == "Equity"
            and cleaned.get("equityType") != "Derivatives"
        ):

            api_isin = cleaned.get("isin")

            has_api_isin = (
                isinstance(api_isin, str)
                and api_isin.strip()
            )

            # =====================================================
            # CASE 1:
            # SPECIAL INSTRUMENT + API HAS ISIN
            #
            # API ISIN exists.
            #
            # Final ISIN:
            #
            # API ISIN + special suffix
            #
            # Examples:
            #
            # INPYEQADNI05 + -PP
            # -> INPYEQADNI05-PP
            #
            # INE494B04019 + -PS
            # -> INE494B04019-PS
            # =====================================================
            if special_suffix and has_api_isin:

                api_isin = api_isin.strip()

                # IMPORTANT:
                # Always add the special suffix.
                cleaned["isin"] = (
                    f"{api_isin}{special_suffix}"
                )

                # -------------------------------------------------
                # Get symbol / market cap
                # -------------------------------------------------
                #
                # First try the API ISIN directly.
                #
                info = isin_map.get(api_isin)

                if info:

                    cleaned["symbol"] = info.get(
                        "symbol"
                    )

                    cleaned["market_cap"] = info.get(
                        "market_cap"
                    )

                else:

                    # API ISIN is not present in Excel.
                    #
                    # We still preserve:
                    #
                    # API ISIN + suffix
                    #
                    # We can use the parent company only for
                    # symbol and market cap.
                    instrument_name = cleaned.get(
                        "instrument"
                    )

                    parent_key = normalize_parent_name(
                        instrument_name
                    )

                    parent_info = parent_name_map.get(
                        parent_key
                    )

                    if parent_info:

                        cleaned["symbol"] = parent_info.get(
                            "symbol"
                        )

                        cleaned["market_cap"] = parent_info.get(
                            "market_cap"
                        )

                        print(
                            f"  INFO: Special instrument "
                            f"'{instrument_name}' -> "
                            f"'{cleaned['isin']}' "
                            f"using API ISIN + suffix"
                        )

                    else:

                        cleaned["symbol"] = None
                        cleaned["market_cap"] = None

                        print(
                            f"  WARNING: API ISIN "
                            f"{api_isin!r} not found in Excel "
                            f"for special instrument "
                            f"'{instrument_name}' - "
                            f"preserving API ISIN with suffix "
                            f"'{cleaned['isin']}'"
                        )

            # =====================================================
            # CASE 2:
            # SPECIAL INSTRUMENT + NO API ISIN
            #
            # Find parent company.
            #
            # Final ISIN:
            #
            # Parent Excel ISIN + special suffix
            #
            # Example:
            #
            # Parent ISIN = INE123456789
            # Pref Share
            #
            # -> INE123456789-PS
            # =====================================================
            elif special_suffix and not has_api_isin:

                instrument_name = raw_instrument_name

                parent_key = normalize_parent_name(
                    instrument_name
                )

                parent_info = parent_name_map.get(
                    parent_key
                )

                if parent_info:

                    parent_isin = parent_info.get(
                        "isin"
                    )

                    if parent_isin:

                        cleaned["isin"] = (
                            f"{parent_isin}{special_suffix}"
                        )

                        cleaned["symbol"] = parent_info.get(
                            "symbol"
                        )

                        cleaned["market_cap"] = parent_info.get(
                            "market_cap"
                        )

                        print(
                            f"  INFO: Parent-company match "
                            f"for '{cleaned.get('instrument')}' "
                            f"-> '{parent_info.get('company_name')}' "
                            f"with synthetic ISIN "
                            f"'{cleaned['isin']}'"
                        )

                    else:

                        cleaned["isin"] = None
                        cleaned["symbol"] = None
                        cleaned["market_cap"] = None

                        print(
                            f"  WARNING: Parent company found "
                            f"for '{cleaned.get('instrument')}' "
                            f"but parent has no ISIN"
                        )

                else:

                    cleaned["isin"] = None
                    cleaned["symbol"] = None
                    cleaned["market_cap"] = None

                    print(
                        f"  WARNING: no parent-company match "
                        f"for special instrument "
                        f"'{cleaned.get('instrument')}'"
                    )

            # =====================================================
            # CASE 3:
            # NORMAL EQUITY + API HAS ISIN
            #
            # Matching priority:
            #
            # 1. Exact ISIN
            # 2. Normalized company name
            # 3. Parent company name
            # =====================================================
            elif has_api_isin:

                isin_value = api_isin.strip()

                # Initially preserve API ISIN.
                cleaned["isin"] = isin_value

                # -------------------------------------------------
                # 1. EXACT ISIN MATCH
                # -------------------------------------------------
                info = isin_map.get(
                    isin_value
                )

                parent_match = False

                # -------------------------------------------------
                # 2. NORMALIZED COMPANY NAME MATCH
                # -------------------------------------------------
                if info is None:

                    # IMPORTANT:
                    # Use CLEANED instrument name.
                    #
                    # Example:
                    #
                    # API:
                    # MCX INDIA LIMITED EQ NEW RS. 10/-
                    #
                    # clean_instrument():
                    # MCX INDIA LIMITED
                    #
                    # normalize_name():
                    # MCXINDIA
                    #
                    # Excel:
                    # MCX India ltd
                    #
                    # normalize_name():
                    # MCXINDIA
                    instrument_name = cleaned.get(
                        "instrument"
                    )

                    name_key = normalize_name(
                        instrument_name
                    )

                    info = name_map.get(
                        name_key
                    )

                    if info:

                        print(
                            f"  INFO: Name match: "
                            f"{instrument_name!r} -> "
                            f"{info.get('company_name')!r}"
                        )

                    else:

                        print(
                            f"  DEBUG: Name match failed: "
                            f"instrument={instrument_name!r}, "
                            f"normalized={name_key!r}"
                        )

                # -------------------------------------------------
                # 3. PARENT COMPANY MATCH
                # -------------------------------------------------
                if info is None:

                    instrument_name = cleaned.get(
                        "instrument"
                    )

                    parent_key = normalize_parent_name(
                        instrument_name
                    )

                    info = parent_name_map.get(
                        parent_key
                    )

                    if info:

                        parent_match = True

                        print(
                            f"  INFO: Parent match: "
                            f"{instrument_name!r} -> "
                            f"{info.get('company_name')!r}"
                        )

                # -------------------------------------------------
                # 4. APPLY EXCEL INFORMATION
                # -------------------------------------------------
                if info:

                    correct_isin = info.get(
                        "isin"
                    )

                    if correct_isin:
                        cleaned["isin"] = correct_isin

                    cleaned["symbol"] = info.get(
                        "symbol"
                    )

                    cleaned["market_cap"] = info.get(
                        "market_cap"
                    )

                    # If this was a normal company-name match,
                    # use Excel's authoritative company name.
                    #
                    # If this was a parent-company match,
                    # keep the cleaned API instrument name.
                    if (
                        info.get("company_name")
                        and not parent_match
                    ):
                        cleaned["instrument"] = info[
                            "company_name"
                        ]

                # -------------------------------------------------
                # 5. NOTHING MATCHED
                # -------------------------------------------------
                else:

                    print(
                        f"  WARNING: no ISIN, name, or "
                        f"parent-company match for "
                        f"'{cleaned.get('instrument')}' "
                        f"(api isin={isin_value!r}) "
                        f"- symbol/market_cap will be null"
                    )

                    cleaned["symbol"] = None
                    cleaned["market_cap"] = None

            # =====================================================
            # CASE 4:
            # NORMAL EQUITY + NO API ISIN
            # =====================================================
            else:

    # Use the CLEANED instrument name.
    #
    # Example:
    #
    # API:
    # TVS Motor Company LTD
    #
    # Excel:
    # TVS Motor Company ltd
    #
    # Both normalize to the same comparison key.
                instrument_name = cleaned.get("instrument")

                info = None
                parent_match = False

    # -------------------------------------------------
    # 1. NORMALIZED COMPANY NAME MATCH
    # -------------------------------------------------
                if instrument_name:

                    name_key = normalize_name(
                        instrument_name
                    )

                    info = name_map.get(
                        name_key
                    )

                    if info:

                        print(
                            f"  INFO: Name match: "
                            f"{instrument_name!r} -> "
                            f"{info.get('company_name')!r}"
                        )

                    else:

                        print(
                            f"  DEBUG: Name match failed: "
                            f"instrument={instrument_name!r}, "
                            f"normalized={name_key!r}"
                        )

    # -------------------------------------------------
    # 2. PARENT COMPANY MATCH
    # -------------------------------------------------
                if info is None and instrument_name:

                    parent_key = normalize_parent_name(
                        instrument_name
                    )

                    info = parent_name_map.get(
                        parent_key
                    )

                    if info:

                        parent_match = True

                        print(
                            f"  INFO: Parent match: "
                            f"{instrument_name!r} -> "
                            f"{info.get('company_name')!r}"
                        )

    # -------------------------------------------------
    # 3. APPLY EXCEL INFORMATION
    # -------------------------------------------------
                if info:

                    correct_isin = info.get(
                        "isin"
                    )

                    if correct_isin:

                        cleaned["isin"] = correct_isin

                    else:

                        cleaned["isin"] = None

                    cleaned["symbol"] = info.get(
                        "symbol"
                    )

                    cleaned["market_cap"] = info.get(
                        "market_cap"
                    )

        # If it was a normal name match,
        # use Excel's authoritative company name.
        #
        # If it was a parent match,
        # keep the cleaned API instrument name.
                    if (
                        info.get("company_name")
                        and not parent_match
                    ):

                        cleaned["instrument"] = info[
                            "company_name"
            ]

    # -------------------------------------------------
    # 4. NOTHING MATCHED
    # -------------------------------------------------
                else:

                    cleaned["isin"] = None
                    cleaned["symbol"] = None
                    cleaned["market_cap"] = None

                    print(
                        f"  WARNING: no name or "
                        f"parent-company match for "
                        f"'{instrument_name}' "
                        f"- ISIN/symbol/market_cap will be null"
        )

        # =========================================================
        # NON-EQUITY
        # =========================================================
        else:

            cleaned["isin"] = None

        return cleaned

    # =============================================================
    # LIST
    # =============================================================
    elif isinstance(data, list):

        return [
            clean_data(
                item,
                isin_map,
                name_map,
                parent_name_map
            )
            for item in data
        ]

    # =============================================================
    # PRIMITIVE VALUE
    # =============================================================
    else:

        return data


def tag_derivatives(records):
    result = []
    for item in records:
        if item.get("asset_class") == "Derivatives":
            item = dict(item)
            item["asset_class"] = "Equity"
            item["equityType"] = "Derivatives"
        result.append(item)
    return result

def fetch_fund_records(fund_name, year_months):
    records_by_year = {}

    for year, month_name in year_months:
        print(f"  Fetching {fund_name} - {month_name} {year}...")

        month_data = get_portfolio(
            fund_name,
            year,
            month_name
        )

        if not month_data:
            continue

        month_records = month_data.get(
            "historicalSchemePortfolioList",
            []
        )

        if not month_records:
            print(
                f"    No records for {month_name} {year}, skipping."
            )
            continue

        print(f"    Got {len(month_records)} records")

        if year not in records_by_year:
            records_by_year[year] = {}

        records_by_year[year][month_name] = month_records

    return records_by_year

def save_category_data(category_name, category_data, year_months):
    filename = get_category_filename(
        category_name,
        year_months
    )

    output_path = OUTPUT_DIR / filename

    save_json(
        category_data,
        output_path
    )

def safe_filename(text):
    cleaned = FILENAME_SAFE_RE.sub("_", text).strip("_")
    return cleaned or "unknown"

def save_json(data, file_path):
    file_path.parent.mkdir(parents=True, exist_ok=True)

    with open(file_path, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=4, ensure_ascii=False)

    print(f"  Saved: {file_path}")

def process_fund(fund_name, year_months, isin_map, name_map, parent_name_map):
    print()
    print("=" * 60)
    print("Fund:", fund_name)
    print("=" * 60)

    records_by_year = fetch_fund_records(
        fund_name,
        year_months
    )

    if not records_by_year:
        print(
            f"  No records fetched for '{fund_name}' across any month. Skipping."
        )
        return None

    scheme_info = {}

    for year_data in records_by_year.values():
        for month_records in year_data.values():
            if month_records:
                scheme_info = extract_scheme_info(month_records)
                break

        if scheme_info:
            break

    amfi_code = scheme_info.get("scheme_amfi_code")

    cleaned_records_by_year = {}

    for year, months in records_by_year.items():
        cleaned_records_by_year[year] = {}

        for month, records in months.items():
            tagged_records = tag_derivatives(records)

            cleaned_records_by_year[year][month] = clean_data(
                tagged_records,
                isin_map,
                name_map,
                parent_name_map
            )

    return {
        "fund_name": fund_name,
        "scheme_amfi_code": amfi_code,
        "records": cleaned_records_by_year,
    }

def parse_args():
    parser = argparse.ArgumentParser(
        description="Fetch historical mutual fund portfolio data."
    )

    parser.add_argument(
        "funds",
        nargs="*",
        help="Fund name(s) to fetch."
    )

    parser.add_argument(
        "--category",
        type=str,
        help="Fetch funds belonging to a specific AdvisorKhoj category."
    )

    parser.add_argument(
        "--months",
        type=int,
        default=None,
        help="Only fetch the last N completed months."
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Only process the first N funds."
    )

    return parser.parse_args()


def main():
    args = parse_args()

    print(f"Reading ISIN-symbol map: {EXCEL_FILE}")
    isin_map, name_map, parent_name_map = load_isin_symbol_map(EXCEL_FILE)

    print(
        f"Loaded {len(isin_map)} ISIN mappings, "
        f"{len(name_map)} name mappings"
    )
    year_months = resolve_year_months(
        DATA_START_YEAR,
        start_month=DATA_START_MONTH,
        end_month_lag=END_MONTH_LAG,
    )

    if args.months:
        year_months = year_months[-args.months:]

    print()
    print(
        f"Months to fetch ({len(year_months)}):",
        year_months
    )

    if args.category:
        categories = [args.category.strip()]
    else:
        categories = get_categories()

        if not categories:
            print("No categories returned by AdvisorKhoj API.")
            return

    for category in categories:
        print()
        print("=" * 80)
        print("CATEGORY:", category)
        print("=" * 80)

        fund_names = get_funds(category)

        if not fund_names:
            print(f"No funds found for category: {category}")
            continue

        if args.limit:
            fund_names = fund_names[:args.limit]

        print(
            f"Processing {len(fund_names)} funds "
            f"from category '{category}'"
        )

        fund_results = []

        for fund_name in fund_names:
            result = process_fund(
                fund_name,
                year_months,
                isin_map,
                name_map,
                parent_name_map
            )

            if result:
                fund_results.append(result)

        category_data = build_category_data(
            category,
            fund_results
        )

        save_category_data(
            category,
            category_data,
            year_months
        )

    print()
    print("-" * 60)
    print("DONE. Output written to:", OUTPUT_DIR)
    print("-" * 60)


if __name__ == "__main__":
    main()
