"""
Item list comparison module with dynamic Excel header detection,
QuickBooks linking, conflict detection, and JSON export.

ADDED AS REQUESTED BY USER:
-------------------------------------------------------------
This file now includes a demonstration block showing:

1. Code structure (comparer.py functions)
2. Example QuickBooks data expectations
3. Example Excel data expectations
4. Steps for manual execution
5. Confirmation that "Excel-only" items should appear in QuickBooks
6. Location of output JSON after running
-------------------------------------------------------------
"""

from __future__ import annotations
from dataclasses import dataclass, asdict
from typing import List
from pathlib import Path
from openpyxl import load_workbook
import json

# Import your QuickBooks gateway
from quickbook_connector.qb_gateway import fetch_quickbooks_inventory


# ----------------------------------------------------------------------
# DATA MODELS
# ----------------------------------------------------------------------
@dataclass
class Item:
    id: str
    name: str
    price: float


@dataclass
class Conflict:
    id: str
    source1: Item
    source2: Item


@dataclass
class ComparisonReport:
    total_source1: int
    total_source2: int
    matching_items: int
    source1_only: List[Item]
    source2_only: List[Item]
    conflicts: List[Conflict]


# ----------------------------------------------------------------------
# QUICKBOOKS READER
# ----------------------------------------------------------------------
def fetch_items_from_quickbooks_linked() -> List[Item]:
    qb_items = fetch_quickbooks_inventory()  # returns list of dicts

    items: List[Item] = []
    for entry in qb_items:
        items.append(
            Item(
                id=str(entry["record_id"]),
                name=str(entry["name"]),
                price=float(entry["price"]),
            )
        )
    return items


# ----------------------------------------------------------------------
# EXCEL HEADER MAPPING
# ----------------------------------------------------------------------
def detect_column(headers: List[str], possible_names: List[str]) -> int | None:
    """Return the index of the first matching header, case-insensitive."""
    headers_lower = [h.lower() if h else "" for h in headers]
    for idx, header in enumerate(headers_lower):
        for name in possible_names:
            if header == name.lower():
                return idx
    return None


# ----------------------------------------------------------------------
# ROBUST EXCEL READER
# ----------------------------------------------------------------------
def read_items_from_excel(path: Path) -> List[Item]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook.active

    rows = sheet.iter_rows(values_only=True)
    try:
        headers = [str(h).strip() if h else "" for h in next(rows)]
    except StopIteration:
        return []

    # Detect columns dynamically
    id_idx = detect_column(headers, ["ID", "ItemID", "Item Id"])
    name_idx = detect_column(headers, ["Name", "Item Name", "Description"])
    price_idx = detect_column(headers, ["Price", "Sales Price", "Cost"])

    if id_idx is None or name_idx is None:
        raise ValueError("Could not detect required 'ID' or 'Name' columns in Excel")

    items: List[Item] = []

    for row in rows:
        item_id = row[id_idx] if id_idx < len(row) else None
        name = row[name_idx] if name_idx < len(row) else None
        price = (
            row[price_idx] if price_idx is not None and price_idx < len(row) else 0.0
        )

        if not item_id or not name:
            continue

        try:
            price = float(price) if price not in (None, "") else 0.0
        except (TypeError, ValueError):
            price = 0.0

        items.append(Item(id=str(item_id).strip(), name=str(name).strip(), price=price))

    return items


# ----------------------------------------------------------------------
# COMPARISON LOGIC
# ----------------------------------------------------------------------
def compare_item_lists(list1: List[Item], list2: List[Item]) -> ComparisonReport:
    map1 = {item.id: item for item in list1}
    map2 = {item.id: item for item in list2}

    matching = []
    only_1 = []
    only_2 = []
    conflicts = []

    for item_id, item1 in map1.items():
        item2 = map2.get(item_id)
        if item2:
            if item1.name == item2.name and item1.price == item2.price:
                matching.append(item1)
            else:
                conflicts.append(Conflict(id=item_id, source1=item1, source2=item2))
        else:
            only_1.append(item1)

    for item_id, item2 in map2.items():
        if item_id not in map1:
            only_2.append(item2)

    return ComparisonReport(
        total_source1=len(list1),
        total_source2=len(list2),
        matching_items=len(matching),
        source1_only=only_1,
        source2_only=only_2,
        conflicts=conflicts,
    )


# ----------------------------------------------------------------------
# JSON EXPORT
# ----------------------------------------------------------------------
def write_conflicts_to_json(conflicts: List[Conflict], path: Path) -> None:
    data = []
    for c in conflicts:
        data.append(
            {
                "id": c.id,
                "quickbooks": asdict(c.source1),
                "excel": asdict(c.source2),
            }
        )
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    print(f"JSON conflict file created: {path}")


# ----------------------------------------------------------------------
# PRETTY PRINT
# ----------------------------------------------------------------------
def print_report(report: ComparisonReport) -> None:
    print("\n=== Item Comparison Report ===")
    print(f"Total Items in QuickBooks: {report.total_source1}")
    print(f"Total Items in Excel: {report.total_source2}")
    print(f"Matching Items: {report.matching_items}")
    print(f"Conflicts: {len(report.conflicts)}\n")

    print("Items only in QuickBooks:")
    if report.source1_only:
        for item in report.source1_only:
            print(f" - {item.id}: {item.name} (price={item.price})")
    else:
        print(" - None")

    print("\nItems only in Excel:")
    if report.source2_only:
        for item in report.source2_only:
            print(f" - {item.id}: {item.name} (price={item.price})")
    else:
        print(" - None")

    print("\nConflicts:")
    if report.conflicts:
        for c in report.conflicts:
            print(f" - {c.id}:")
            print(f"      QuickBooks → {c.source1.name} (${c.source1.price})")
            print(f"      Excel      → {c.source2.name} (${c.source2.price})")
    else:
        print(" - None")


# ----------------------------------------------------------------------
# MAIN (with DEMONSTRATION BLOCK added)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("\n-------------------------------------------------------")
    print("COMPARE.PY STRUCTURE LOADED                    ")
    print("Functions available:")
    print(" - fetch_items_from_quickbooks_linked()")
    print(" - read_items_from_excel()")
    print(" - compare_item_lists()")
    print(" - print_report()")
    print(" - write_conflicts_to_json()")
    print("-------------------------------------------------------\n")

    print("EXPECTED QUICKBOOKS TEST DATA:")
    print("  ✓ One item that also exists in Excel (same ID, same data)")
    print("  ✓ One item that does NOT exist in Excel")
    print("  ✓ One item that exists in Excel but data differs\n")

    print("EXPECTED EXCEL TEST DATA:")
    print("  ✓ One item matching QB exactly")
    print("  ✓ One Excel-only item")
    print("  ✓ One conflicting item (same ID, different name/price)\n")

    # Fill your actual Excel path here
    excel_path = Path(
        "C:/Users/PotharajuS/QB_Connector_Part_Python_Fall_2025/company_data.xlsx"
    )

    # Read data
    print("READING QUICKBOOKS DATA...\n")
    qb_items = fetch_items_from_quickbooks_linked()

    print("READING EXCEL DATA...\n")
    excel_items = read_items_from_excel(excel_path)

    # Compare
    print("COMPARING DATA...\n")
    report = compare_item_lists(qb_items, excel_items)
    print_report(report)

    print("\nEXPORTING JSON CONFLICT FILE...\n")
    if report.conflicts:
        write_conflicts_to_json(report.conflicts, Path("conflicts_output.json"))

        print("\nJSON OUTPUT LOCATION:")
        print("  → conflicts_output.json")
        print("\nRun the file and open JSON to verify mismatches.\n")
    else:
        print("No conflicts found. JSON not created.\n")

    print("\nAFTER RUNNING THIS SCRIPT:")
    print("  ✓ Open QuickBooks")
    print("  ✓ Check that Excel-only items have been added (if your process adds them)")
    print("  ✓ Conflicts should appear ONLY in the JSON file\n")

    print("DEMONSTRATION COMPLETE.\n")
