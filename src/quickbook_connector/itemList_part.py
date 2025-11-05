"""Excel extraction for inventory items (Parts to purchase/sell)."""

from __future__ import annotations
from pathlib import Path
from typing import List
from openpyxl import load_workbook

from quickbook_connector.models import InventoryItem


def extract_inventory_items(workbook_path: Path) -> List[InventoryItem]:
    """Return inventory items parsed from the Excel workbook.

    This function reads the ``parts`` worksheet in the Excel file,
    extracts columns "ID", "Name", and "Sales Price", and constructs
    :class:`~itemlist_parts_cli.models.InventoryItem` instances with ``source="excel"``.

    Raises:
        FileNotFoundError: If the workbook cannot be found.
        ValueError: If the "parts" worksheet is missing.
    """

    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    workbook = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    try:
        sheet = workbook["parts"]
    except KeyError as exc:
        workbook.close()
        raise ValueError("Worksheet 'parts' not found in workbook") from exc

    rows = sheet.iter_rows(values_only=True)
    headers_row = next(rows, None)
    if headers_row is None:
        workbook.close()
        return []

    headers = [
        str(header).strip() if header is not None else "" for header in headers_row
    ]
    header_index = {header: idx for idx, header in enumerate(headers)}

    def _value(row, column_name: str):
        idx = header_index.get(column_name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    items: List[InventoryItem] = []
    try:
        for row in rows:
            item_id = _value(row, "ID")
            name = _value(row, "Name")
            price = _value(row, "Price")

            if not item_id or not name:
                continue

            try:
                item_id = str(item_id).strip()
                name = str(name).strip()
                price = float(price) if price not in (None, "") else 0.0
            except (TypeError, ValueError):
                continue

            items.append(
                InventoryItem(
                    record_id=item_id,
                    name=name,
                    price=price,
                    source="excel",
                )
            )
    finally:
        workbook.close()

    return items


__all__ = ["extract_inventory_items"]

if __name__ == "__main__":  # pragma: no cover - manual invocation
    import sys

    # Allow running as a script: poetry run python itemList_parts_cli/excel_reader.py
    try:
        parts = extract_inventory_items(
            Path(
                "C:/Users/PotharajuS/QB_Connector_Part_Python_Fall_2025/company_data.xlsx"
            )
        )
        for part in parts:
            print(part)
    except Exception as e:
        print(f"Error: {e}")
        print("Usage:  poetry run python <path-to-workbook.xlsx>")
        sys.exit(1)
