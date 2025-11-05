from __future__ import annotations
import os
from typing import List
import xml.etree.ElementTree as ET
import logging

from quickbook_connector.models import InventoryItem

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _parse_items_from_qbxml(qbxml_response: str) -> List[InventoryItem]:
    """Parse QBXML and build InventoryItem model objects."""
    root = ET.fromstring(qbxml_response)

    item_nodes = (
        root.findall(".//ItemInventoryRet")
        + root.findall(".//ItemNonInventoryRet")
        + root.findall(".//ItemServiceRet")
    )

    items: List[InventoryItem] = []

    for node in item_nodes:
        name_el = node.find("Name")
        mpn_el = node.find("ManufacturerPartNumber")

        price_el = (
            node.find("SalesOrPurchase/SalesPrice")
            or node.find("SalesAndPurchase/SalesPrice")
            or node.find("SalesPrice")
        )

        name = name_el.text.strip() if name_el is not None and name_el.text else ""
        record_id = mpn_el.text.strip() if mpn_el is not None and mpn_el.text else name

        try:
            price = (
                float(price_el.text.strip())
                if price_el is not None and price_el.text
                else 0.0
            )
        except Exception:
            price = 0.0

        if name:
            items.append(
                InventoryItem(
                    record_id=record_id,
                    name=name,
                    price=price,
                    source="quickbooks",
                )
            )

    return items


def _build_item_query_qbxml() -> str:
    return """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <ItemQueryRq requestID="1">
      <ActiveStatus>ActiveOnly</ActiveStatus>
    </ItemQueryRq>
  </QBXMLMsgsRq>
</QBXML>
"""


def _call_quickbooks(qbxml_request: str) -> str:
    """Send a QBXML request to QuickBooks and return response XML."""
    if win32com is None:
        raise RuntimeError(
            "win32com is not available. Install pywin32 and run on Windows with QuickBooks Desktop."
        )

    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor.2")
    app_name = "QB Inventory Gateway"

    try:
        rp.OpenConnection2("", app_name, 1)
        ticket = rp.BeginSession("", 0)

        logger.info("Connected to QuickBooks.")

        response = rp.ProcessRequest(ticket, qbxml_request)

        rp.EndSession(ticket)
        rp.CloseConnection()
        logger.info("QuickBooks session closed.")

        return response

    except Exception as exc:
        try:
            rp.CloseConnection()
        except Exception:
            pass
        raise RuntimeError(f"QuickBooks processing error: {exc}") from exc


def fetch_quickbooks_inventory() -> List[InventoryItem]:
    """Reads inventory items from QuickBooks and returns model objects."""
    qbxml_request = _build_item_query_qbxml()
    response_xml = _call_quickbooks(qbxml_request)
    return _parse_items_from_qbxml(response_xml)


if __name__ == "__main__":
    print("Running fetch_quickbooks_inventory() ...")

    try:
        items = fetch_quickbooks_inventory()
        print(f"Found {len(items)} inventory item(s):\n")

        for item in items:
            print(item)

    except Exception as e:
        print("Error while reading QuickBooks inventory:", e)
        exit(1)
