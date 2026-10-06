import os
from datetime import datetime
from urllib.parse import urljoin
import requests
from dotenv import load_dotenv

load_dotenv()

BASE_HOST = "https://mfapi.advisorkhoj.com/"
API_KEY = os.getenv("API_KEY")

PORTFOLIO_ENDPOINT = os.getenv("API_URL", urljoin(BASE_HOST, "getMutualfundHistoricalPortfolio"))

CATEGORIES_ENDPOINT = urljoin(BASE_HOST, "getAllSchemeCategories")
FUNDS_ENDPOINT = urljoin(BASE_HOST, "getAllSchemesCommonNamesbyCategory")
FUNDS_CATEGORY_PARAM = "category"

MONTH_NAME_MAP = {
    1: "january", 2: "february", 3: "march", 4: "april",
    5: "may", 6: "june", 7: "july", 8: "august",
    9: "september", 10: "october", 11: "november", 12: "december",
}

REQUEST_TIMEOUT = 60

def _get(url, params=None):
    params = dict(params or {})
    params["key"] = API_KEY
    response = requests.get(url, params=params, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.json()

def get_categories():
   
    data = _get(CATEGORIES_ENDPOINT)
    return data.get("list", [])

def get_funds(category):
  
    data = _get(FUNDS_ENDPOINT, params={FUNDS_CATEGORY_PARAM: category})
    return data.get("list", [])

def get_all_funds():
    categories = get_categories()

    all_funds = []

    for category in categories:
        print(f"Fetching funds for category: {category}")

        funds = get_funds(category)

        for fund in funds:
            if fund not in all_funds:
                all_funds.append(fund)

    return all_funds

def get_portfolio(scheme_amfi_common, year, month):
    
    if isinstance(month, int):
        month_name = MONTH_NAME_MAP[month]
    else:
        month_name = str(month).lower()

    try:
        return _get(
            PORTFOLIO_ENDPOINT,
            params={
                "scheme_amfi_common": scheme_amfi_common,
                "year": year,
                "month": month_name,
            },
        )
    except requests.exceptions.RequestException as error:
        print(f"  WARNING: failed to fetch {scheme_amfi_common} {month_name} {year}: {error}")
        return None


def resolve_year_months(
    start_year,
    start_month=1,
    end_month_lag=1,
    months_filter=None,
):
    now = datetime.now()

    current_year = now.year
    current_month = now.month

    total_current_months = current_year * 12 + current_month
    cutoff_total = total_current_months - end_month_lag

    cutoff_year = (cutoff_total - 1) // 12
    cutoff_month = ((cutoff_total - 1) % 12) + 1

    months_to_check = (
        sorted(months_filter)
        if months_filter
        else sorted(MONTH_NAME_MAP.keys())
    )

    result = []

    for year in range(start_year, cutoff_year + 1):
        for month_num in months_to_check:

            if year == start_year and month_num < start_month:
                continue

            if (
                year > cutoff_year
                or (
                    year == cutoff_year
                    and month_num > cutoff_month
                )
            ):
                continue

            result.append(
                (year, MONTH_NAME_MAP[month_num])
            )

    return result

if __name__ == "__main__":
    print("Resolved months:", resolve_year_months(2025))