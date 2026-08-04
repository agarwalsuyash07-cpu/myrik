#!/usr/bin/env python3
"""
pending_orders_report.py

Reads an "Order Confirmed" / pending-status Orders_Export .xlsx file and
prints how many orders are still pending delivery, grouped by city and
by order date (1st July, 2nd July, ... style).

Optionally cross-checks against a separate "Delivered" filter export so
that any order that has since been delivered is excluded from the
pending count (dedup is done by Order Id).

USAGE
-----
    python3 pending_orders_report.py PENDING_FILE.xlsx
    python3 pending_orders_report.py PENDING_FILE.xlsx --delivered DELIVERED_FILE.xlsx
    python3 pending_orders_report.py PENDING_FILE.xlsx --auto-delivered
    python3 pending_orders_report.py PENDING_FILE.xlsx --json out.json

If --delivered is omitted, the script will *not* cross-check unless
--auto-delivered is passed, in which case it scans the same folder as
PENDING_FILE for the most recent export whose "Order Status Filter"
metadata cell is "Delivered" and uses that automatically.

Expected column layout (as exported by the source system), header row
detected automatically by locating the "Customer Phone" cell:

    Customer Phone | Delivery Address | City | Created At | Order Placed At |
    Order Created By | Delivery Partner | Shipped At | Delivered At |
    Delivery TAT | Full Delivery TAT | Cancelled At | Total Price |
    Payment Mode | Order Status | Shipping Status |
    First Action Time (mins) | Order Id | Internal Stock Transfer Order
"""

import argparse
import glob
import json
import os
import sys
from collections import defaultdict
from datetime import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("This script requires openpyxl. Install it with: pip install openpyxl")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def parse_date(date_str):
    """'02-08-2026' (or an ISO date, if openpyxl typed the cell) -> datetime.
    Returns None for anything unparseable, e.g. the 'Unknown' bucket."""
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(date_str), fmt)
        except (ValueError, TypeError):
            continue
    return None


def ordinal(day: int) -> str:
    """Return e.g. '1st', '2nd', '3rd', '11th', '21st' for a day-of-month int."""
    if 11 <= day <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def load_workbook_rows(path):
    """
    Load an Orders_Export .xlsx file.

    Returns (metadata: dict, header: list, rows: list of tuples).
    Locates the header row by finding the row whose first cell is
    'Customer Phone', rather than assuming a fixed row number, so the
    script keeps working even if the metadata block above the table
    changes size.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    metadata = {}
    header_row_idx = None
    header = None

    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and row[0] == "Customer Phone":
            header_row_idx = i
            header = list(row)
            break
        # Collect simple key/value metadata rows (Generated On, Date Range, etc.)
        if row and row[0] and row[1] is not None and row[0] not in (None,):
            metadata[str(row[0])] = row[1]

    if header_row_idx is None:
        raise ValueError(
            f"Could not find the 'Customer Phone' header row in {path!r}. "
            "Is this an Orders_Export file?"
        )

    rows = list(ws.iter_rows(min_row=header_row_idx + 1, values_only=True))
    # Drop fully-empty trailing rows
    rows = [r for r in rows if any(c is not None for c in r)]

    return metadata, header, rows


def column_index(header, name):
    try:
        return header.index(name)
    except ValueError:
        raise ValueError(f"Expected column {name!r} not found in header: {header}")


def find_auto_delivered_file(pending_path):
    """
    Scan the same directory as pending_path for the most recent export
    whose 'Order Status Filter' metadata is 'Delivered'. Returns the path
    or None if none found.
    """
    directory = os.path.dirname(os.path.abspath(pending_path)) or "."
    candidates = []
    for fp in glob.glob(os.path.join(directory, "Orders_Export_*.xlsx")):
        if os.path.abspath(fp) == os.path.abspath(pending_path):
            continue
        try:
            meta, _, _ = load_workbook_rows(fp)
        except Exception:
            continue
        if str(meta.get("Order Status Filter", "")).strip().lower() == "delivered":
            candidates.append((meta.get("Generated On", ""), fp))

    if not candidates:
        return None
    # Pick the most recently generated one (string sort works for the
    # dd-mm-yyyy hh-mm-ss naming convention used in these exports because
    # of consistent zero-padding — fall back to filename sort otherwise).
    candidates.sort(key=lambda t: t[1])
    return candidates[-1][1]


# --------------------------------------------------------------------------
# Core logic
# --------------------------------------------------------------------------

def build_pending_report(pending_path, delivered_path=None):
    meta, header, rows = load_workbook_rows(pending_path)

    idx_city = column_index(header, "City")
    idx_created = column_index(header, "Created At")
    idx_delivered_at = column_index(header, "Delivered At")
    idx_order_id = column_index(header, "Order Id")

    delivered_ids = set()
    delivered_meta = None
    if delivered_path:
        d_meta, d_header, d_rows = load_workbook_rows(delivered_path)
        d_idx_order_id = column_index(d_header, "Order Id")
        delivered_ids = {r[d_idx_order_id] for r in d_rows}
        delivered_meta = d_meta

    # Dedup by Order Id, drop anything already marked delivered in the
    # reference file, and also drop anything that already shows a
    # Delivered At value in the pending file itself (defensive — in
    # practice these exports show '-' for every row here).
    pending = {}
    skipped_already_delivered = 0
    for r in rows:
        oid = r[idx_order_id]
        delivered_at = r[idx_delivered_at]
        if oid in delivered_ids:
            skipped_already_delivered += 1
            continue
        if delivered_at not in (None, "-", ""):
            skipped_already_delivered += 1
            continue
        pending[oid] = r

    data = defaultdict(lambda: defaultdict(int))
    for r in pending.values():
        city = r[idx_city] or "Unknown"
        created = r[idx_created]
        date_str = str(created).split(" ")[0] if created else "Unknown"
        data[city][date_str] += 1

    return {
        "pending_meta": meta,
        "delivered_meta": delivered_meta,
        "total_rows_in_pending_file": len(rows),
        "skipped_already_delivered": skipped_already_delivered,
        "grand_total_pending": len(pending),
        "by_city": data,
    }


# --------------------------------------------------------------------------
# Output formatting
# --------------------------------------------------------------------------

def print_report(report):
    data = report["by_city"]

    def date_sort_key(date_str):
        dt = parse_date(date_str)
        return (dt is None, dt or datetime.max)  # push unparseable dates to the end

    for city in sorted(data):
        print(f"{city}")
        for date_str in sorted(data[city], key=date_sort_key):
            dt = parse_date(date_str)
            label = f"{ordinal(dt.day)} {dt.strftime('%B')}" if dt else date_str
            print(f"{label}: {data[city][date_str]} orders pending")
        print()

    print(f"Grand total pending: {report['grand_total_pending']}")

    if report["skipped_already_delivered"]:
        print(
            f"({report['skipped_already_delivered']} rows excluded — "
            f"already marked delivered in the reference file)"
        )

    if report["delivered_meta"]:
        gen_on = report["delivered_meta"].get("Generated On", "unknown")
        print(f"Cross-checked against Delivered export generated on: {gen_on}")
    else:
        print("No Delivered-status reference file supplied — counts are unchecked.")


def report_to_json(report):
    # Convert defaultdicts to plain dicts for JSON serialization
    serializable = dict(report)
    serializable["by_city"] = {
        city: dict(dates) for city, dates in report["by_city"].items()
    }
    return json.dumps(serializable, indent=2, default=str)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Report pending (undelivered) orders grouped by city and date."
    )
    parser.add_argument("pending_file", help="Path to the pending Orders_Export .xlsx file")
    parser.add_argument(
        "--delivered",
        help="Path to a Delivered-status Orders_Export .xlsx file to cross-check against",
    )
    parser.add_argument(
        "--auto-delivered",
        action="store_true",
        help="Auto-detect the most recent Delivered export in the same folder",
    )
    parser.add_argument(
        "--json",
        metavar="OUT_FILE",
        help="Also write the report as JSON to this path",
    )
    args = parser.parse_args()

    delivered_path = args.delivered
    if not delivered_path and args.auto_delivered:
        delivered_path = find_auto_delivered_file(args.pending_file)
        if delivered_path:
            print(f"[auto-detected delivered reference: {delivered_path}]\n")
        else:
            print("[--auto-delivered: no Delivered export found in folder, skipping cross-check]\n")

    report = build_pending_report(args.pending_file, delivered_path)
    print_report(report)

    if args.json:
        with open(args.json, "w") as f:
            f.write(report_to_json(report))
        print(f"\nJSON report written to {args.json}")


if __name__ == "__main__":
    main()
