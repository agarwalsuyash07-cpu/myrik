"""
Citywise delivery cost report generator.

Reads the "Orders_Export" style xlsx, filters orders by the date they were
marked "Delivered At", and prints a citywise report: total cost, number of
rikshaws (unique delivery partners assigned that day), and order count —
in Indian numbering format (e.g. 5,25,772 or 5,25,772.50).

Orders with no delivery partner assigned ("-" or blank) are not counted
towards the rikshaw count for that city, but are still counted as orders.

Usage:
    python citywise_deliveries.py <path_to_xlsx> <DD-MM-YYYY> [--no-rikshaws]

Example:
    python citywise_deliveries.py Orders_Export_18-07-2026_13-09-24.xlsx 16-07-2026
    python citywise_deliveries.py Orders_Export_18-07-2026_13-09-24.xlsx 16-07-2026 --no-rikshaws

Output format (default, with rikshaws):
    16 July Deliveries
    Fatehpur: 5,25,772rs, 9 rikshaws, 35 Orders
    Banda: 2,47,207rs, 0 rikshaws, 26 Orders
    ...
    Total: 14,93,099rs, 155 orders

Output format (--no-rikshaws):
    16 July Deliveries
    Fatehpur: 5,25,772rs, 35 Orders
    Banda: 2,47,207rs, 26 Orders
    ...
    Total: 14,93,099rs, 155 orders
"""

import sys
from collections import defaultdict

import openpyxl

# Column indices in the export (0-based), based on the header row:
# Customer Phone, Delivery Address, City, Created At, Order Placed At,
# Order Created By, Delivery Partner, Shipped At, Delivered At,
# Delivery TAT, Full Delivery TAT, Cancelled At, Total Price, Payment Mode,
# Order Status, Shipping Status, First Action Time (mins), Order Id,
# Internal Stock Transfer Order
COL_CITY = 2
COL_DELIVERY_PARTNER = 6
COL_DELIVERED_AT = 8
COL_TOTAL_PRICE = 12

HEADER_ROW_INDEX = 6  # 0-based index of the row containing column names

MONTH_NAMES = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December",
}

NO_PARTNER_VALUES = {"", "-"}


def indian_format(n: float) -> str:
    """Format a number in the Indian numbering system (e.g. 12,34,567.50).
    Drops the decimal part entirely if it's zero (e.g. 12,34,567)."""
    n = round(n, 2)
    is_whole = float(n).is_integer()

    if is_whole:
        int_part = str(int(n))
        dec_part = ""
    else:
        s = f"{n:.2f}"
        int_part, dec_part = s.split(".")

    neg = int_part.startswith("-")
    if neg:
        int_part = int_part[1:]

    if len(int_part) <= 3:
        grouped = int_part
    else:
        last3 = int_part[-3:]
        rest = int_part[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        grouped = ",".join(parts) + "," + last3

    if neg:
        grouped = "-" + grouped

    return grouped if not dec_part else f"{grouped}.{dec_part}"


def build_report(
    xlsx_path: str,
    target_date: str,
    sheet_name: str = None,
    include_rikshaws: bool = True,
) -> str:
    """target_date must be in DD-MM-YYYY format, matching the 'Delivered At' column."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    rows = list(ws.iter_rows(values_only=True))
    data_rows = rows[HEADER_ROW_INDEX + 1:]

    city_totals = defaultdict(float)
    city_counts = defaultdict(int)
    city_partners = defaultdict(set)

    for r in data_rows:
        delivered_at = r[COL_DELIVERED_AT]
        if not delivered_at or not str(delivered_at).startswith(target_date):
            continue

        city = r[COL_CITY] or "Unknown"
        try:
            price = float(r[COL_TOTAL_PRICE])
        except (TypeError, ValueError):
            price = 0.0

        city_totals[city] += price
        city_counts[city] += 1

        partner = r[COL_DELIVERY_PARTNER]
        partner_clean = str(partner).strip() if partner is not None else ""
        if partner_clean not in NO_PARTNER_VALUES:
            city_partners[city].add(partner_clean)

    total_amount = sum(city_totals.values())
    total_orders = sum(city_counts.values())

    day, month, year = target_date.split("-")
    month_name = MONTH_NAMES.get(month, month)
    lines = [f"{int(day)} {month_name} Deliveries"]

    for city, amount in sorted(city_totals.items(), key=lambda x: -x[1]):
        if include_rikshaws:
            rikshaw_count = len(city_partners[city])
            lines.append(
                f"{city}: {indian_format(amount)}rs, {rikshaw_count} rikshaws, "
                f"{city_counts[city]} Orders"
            )
        else:
            lines.append(
                f"{city}: {indian_format(amount)}rs, {city_counts[city]} Orders"
            )

    lines.append(f"Total: {indian_format(total_amount)}rs, {total_orders} orders")

    return "\n".join(lines)


if __name__ == "__main__":
    args = sys.argv[1:]
    no_rikshaws = "--no-rikshaws" in args
    if no_rikshaws:
        args.remove("--no-rikshaws")

    if len(args) != 2:
        print("Usage: python citywise_deliveries.py <path_to_xlsx> <DD-MM-YYYY> [--no-rikshaws]")
        sys.exit(1)

    xlsx_path, target_date = args
    print(build_report(xlsx_path, target_date, include_rikshaws=not no_rikshaws))
