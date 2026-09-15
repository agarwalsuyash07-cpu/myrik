#!/usr/bin/env python3
"""
Cancelled / Marked-for-Cancellation Order Report Generator

Takes ONE Orders_Export .xlsx file (status = 'Cancelled' or 'Marked for Cancellation')
and produces a per-city breakdown report + anomaly flags.

Usage:
  python3 cancelled_report.py <file.xlsx>

Report format (per city):
  City: N order(s): price1rs, price2rs, ...
  Subtotal: Xrs

Ends with:
  TOTAL: N cancellations, Xrs

Also flags (printed after the report, not part of the report body):
  - 0rs orders
  - duplicate amounts within the same city (possible duplicate/batch orders)
  - outlier amounts (>= 5x the day's median cancelled order value)
"""
import sys
import openpyxl
from collections import defaultdict, Counter
from statistics import median


VALID_STATUSES = ('Cancelled', 'Marked for Cancellation')


def indian_fmt(n):
    n = int(round(n))
    sign = '-' if n < 0 else ''
    n = abs(n)
    s = str(n)
    if len(s) <= 3:
        return sign + s
    last3 = s[-3:]
    rest = s[:-3]
    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        parts.insert(0, rest)
    return sign + ','.join(parts) + ',' + last3


def load_export(path):
    """Returns (meta, header, data_rows)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    meta = {}
    header_idx = None
    for i, r in enumerate(rows):
        if r[0] in ('Generated On', 'Date Range', 'Order Status Filter', 'Total Orders'):
            meta[r[0]] = r[1]
        if 'Customer Phone' in r:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError(f"Could not find header row in {path}")

    header = rows[header_idx]
    city_idx = header.index('City')
    data = [r for r in rows[header_idx + 1:] if r[city_idx] is not None]
    return meta, header, data


def build_cancelled_report(data, header, date_label):
    # Column positions shift as the export gains columns over time, so
    # they're looked up by header name instead of hardcoded here.
    city_idx = header.index('City')
    price_idx = header.index('Total Price')
    canc = defaultdict(list)
    for r in data:
        city = r[city_idx]
        try:
            price = float(r[price_idx])
        except (TypeError, ValueError):
            price = 0.0
        canc[city].append(price)

    lines = [f"Cancelled Orders Report — {date_label}", "=" * 40]
    grand = 0.0
    total_orders = 0
    for city in sorted(canc):
        prices = canc[city]
        total_orders += len(prices)
        subtotal = sum(prices)
        grand += subtotal
        price_str = ', '.join(indian_fmt(p) + 'rs' for p in prices)
        lines.append(f"{city}: {len(prices)} order(s): {price_str}")
        lines.append(f"Subtotal: {indian_fmt(subtotal)}rs")
    lines.append(f"TOTAL: {total_orders} cancellations, {indian_fmt(grand)}rs")
    return "\n".join(lines), canc


def build_flags(canc):
    all_prices = [p for prices in canc.values() for p in prices]
    flags = []

    # Zero-value orders
    zero_orders = [(city, p) for city, prices in canc.items() for p in prices if p == 0]
    if zero_orders:
        cities = ', '.join(sorted(set(c for c, _ in zero_orders)))
        flags.append(f"{len(zero_orders)} order(s) priced at 0rs — cities: {cities}")

    # Duplicate amounts within the same city
    for city, prices in canc.items():
        counts = Counter(p for p in prices if p > 0)
        dups = {p: n for p, n in counts.items() if n > 1}
        if dups:
            dup_str = ', '.join(f"{indian_fmt(p)}rs x{n}" for p, n in sorted(dups.items()))
            flags.append(f"{city}: repeated amounts — {dup_str} (possible duplicate/batch orders)")

    # Outliers: >= 5x day's median (using non-zero prices)
    nonzero = [p for p in all_prices if p > 0]
    if len(nonzero) >= 3:
        med = median(nonzero)
        if med > 0:
            outliers = [(city, p) for city, prices in canc.items() for p in prices if p >= 5 * med]
            if outliers:
                out_str = ', '.join(f"{city}: {indian_fmt(p)}rs" for city, p in outliers)
                flags.append(f"Outlier amount(s) (>=5x day's median of {indian_fmt(med)}rs): {out_str}")

    return flags


def main():
    if len(sys.argv) != 2:
        print("Usage: python3 cancelled_report.py <file.xlsx>")
        sys.exit(1)

    path = sys.argv[1]
    meta, header, data = load_export(path)
    status = meta.get('Order Status Filter', '')
    date_range = meta.get('Date Range', '')
    date_label = date_range.split(' to ')[0].split(' ')[0] if date_range else "Unknown Date"

    if status not in VALID_STATUSES:
        print(f"WARNING: Order Status Filter is {status!r}, expected 'Cancelled' or 'Marked for Cancellation'.")
        print("Proceeding anyway — verify this is the right file.\n")

    report, canc = build_cancelled_report(data, header, date_label)
    print(report)

    flags = build_flags(canc)
    if flags:
        print()
        print("Flags:")
        for f in flags:
            print(f"- {f}")


if __name__ == "__main__":
    main()
