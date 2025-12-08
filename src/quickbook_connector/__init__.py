"""Customer CLI toolkit.

Exposes the high-level ``<>`` API for programmatic use.
"""

from .qb_comparer import (
    fetch_items_from_quickbooks_linked,
    read_items_from_excel,
    compare_item_lists,
    print_report,
    write_conflicts_to_json,
    push_new_items_to_quickbooks,
)
# Public API for synchronisation

__all__ = [
    "fetch_items_from_quickbooks_linked",
    "read_items_from_excel",
    "compare_item_lists",
    "print_report",
    "write_conflicts_to_json",
    "push_new_items_to_quickbooks",
]  # Re-exported symbol
