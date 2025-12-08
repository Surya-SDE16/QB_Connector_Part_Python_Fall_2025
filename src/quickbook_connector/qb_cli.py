"""
Command-line interface for the QuickBooks Item/Parts comparer.

This module provides the entry point for running the item comparison tool
from the command line. It runs the comparison against QuickBooks Desktop,
and prints where the JSON reports were written.
"""

from __future__ import annotations

import sys
from pathlib import Path

from .qb_comparer import (
    fetch_items_from_quickbooks_linked,
    read_items_from_excel,
    compare_item_lists,
    print_report,
    write_conflicts_to_json,
    push_new_items_to_quickbooks,
)


def main() -> int:
    # Hardcoded Excel path
    excel_path = Path(
        "C:/Users/PotharajuS/QB_Connector_Part_Python_Fall_2025/company_data.xlsx"
    )
    if not excel_path.exists():
        print(f"ERROR: Excel file not found at: {excel_path}")
        return 1

    # Output paths
    json_out_path = Path("conflicts_output.json")

    # Load data
    qb_items = fetch_items_from_quickbooks_linked()
    excel_items = read_items_from_excel(excel_path)

    # Compare results
    report = compare_item_lists(qb_items, excel_items)
    print_report(report)

    # Export outputs
    write_conflicts_to_json(report, json_out_path)

    # Push Excel-only items into QuickBooks
    push_new_items_to_quickbooks(report)

    print(f"JSON report written to: {json_out_path}")

    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
