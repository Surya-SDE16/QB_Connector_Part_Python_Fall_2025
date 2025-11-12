from __future__ import annotations
from dataclasses import dataclass
from typing import Literal

SourceLiteral = Literal["excel", "quickbooks"]

@dataclass(slots=True)
class InventoryItem:
    """Represents an inventory item found in Excel or QuickBooks."""

    record_id: str      # Unique ID of the inventory item
    name: str           # Item name
    price: float        # Item sales price
    source: SourceLiteral  # 'excel' or 'quickbooks'

    def __str__(self) -> str:
        """Return a readable string representation of the item."""
        return (
            f"InventoryItem("
            f"id={self.record_id}, "
            f"name={self.name}, "
            f"price={self.price}, "
            f"source={self.source})"
        )

__all__ = ["InventoryItem", "SourceLiteral"]
