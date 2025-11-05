"""Comparison helpers for inventory items between Excel and QuickBooks."""

from __future__ import annotations
from typing import Iterable
from .models import ComparisonReport, Conflict, InventoryItem


def compare_inventory_items(
    excel_items: Iterable[InventoryItem],
    qb_items: Iterable[InventoryItem],
) -> ComparisonReport:
    """Compare Excel and QuickBooks inventory items and identify discrepancies.

    This function compares item lists from Excel (parts sheet) and QuickBooks (Item List)
    by their ``record_id`` field — where:
        - Excel ID Field = "ID"
        - QB Company ID Field = "Manufacturer's Part Number"

    It identifies:
        1. Items only in Excel
        2. Items only in QuickBooks
        3. Conflicts (same ID but different name or sales price)
    """

    # Index items by record_id for quick lookup
    excel_dict = {item.record_id: item for item in excel_items}
    qb_dict = {item.record_id: item for item in qb_items}

    excel_ids = set(excel_dict.keys())
    qb_ids = set(qb_dict.keys())

    excel_only_ids = excel_ids - qb_ids
    qb_only_ids = qb_ids - excel_ids
    common_ids = excel_ids & qb_ids

    # 1. Items only in Excel
    excel_only = [excel_dict[i] for i in sorted(excel_only_ids)]

    # 2. Items only in QuickBooks
    qb_only = [qb_dict[i] for i in sorted(qb_only_ids)]

    # 3. Conflicts (name or sales price mismatch)
    conflicts = []
    for record_id in sorted(common_ids):
        excel_item = excel_dict[record_id]
        qb_item = qb_dict[record_id]

        if (
            excel_item.name != qb_item.name
            or abs(excel_item.price - qb_item.sales_price) > 0.0001
        ):
            conflicts.append(
                Conflict(
                    record_id=record_id,
                    excel_name=excel_item.name,
                    qb_name=qb_item.name,
                    excel_price=excel_item.Price,
                    qb_price=qb_item.sales_price,
                    reason="name_or_price_mismatch",
                )
            )

    report = ComparisonReport()
    report.excel_only = excel_only
    report.qb_only = qb_only
    report.conflicts = conflicts

    return report


__all__ = ["compare_inventory_items"]
