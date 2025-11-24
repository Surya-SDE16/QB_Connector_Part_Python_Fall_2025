from __future__ import annotations
import xml.etree.ElementTree as ET
from typing import List
import logging

from quickbook_connector.models import InventoryItem

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _call_quickbooks(qbxml_request: str) -> str:
    if win32com is None:
        raise RuntimeError(
            "win32com is not available. Use Windows with pywin32 + QuickBooks Desktop."
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


def _parse_items_from_qbxml(qbxml_response: str) -> List[dict]:
    root = ET.fromstring(qbxml_response)
    item_nodes = root.findall(".//ItemInventoryRet")
    items = []
    for node in item_nodes:
        name = node.findtext("Name", default="")
        record_id = node.findtext("ManufacturerPartNumber", default=name)
        price_text = node.findtext("SalesPrice")
        try:
            price = float(price_text) if price_text is not None else 0.0
        except ValueError:
            price = 0.0
        items.append({"record_id": record_id, "name": name, "price": price})
    return items


def fetch_quickbooks_inventory() -> List[dict]:
    qbxml_request = _build_item_query_qbxml()
    response_xml = _call_quickbooks(qbxml_request)
    return _parse_items_from_qbxml(response_xml)


def _build_items_add_qbxml(items: List[InventoryItem]) -> str:
    msg_rq = ""
    for item in items:
        msg_rq += f"""
    <ItemInventoryAddRq>
      <ItemInventoryAdd>
        <Name>{item.name}</Name>
        <ManufacturerPartNumber>{item.record_id}</ManufacturerPartNumber>
        <SalesPrice>{item.price:.2f}</SalesPrice>
        <IncomeAccountRef><FullName>Sales</FullName></IncomeAccountRef>
        <COGSAccountRef><FullName>Cost of Goods Sold</FullName></COGSAccountRef>
        <AssetAccountRef><FullName>Inventory Asset</FullName></AssetAccountRef>
      </ItemInventoryAdd>
    </ItemInventoryAddRq>
        """
    return f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    {msg_rq}
  </QBXMLMsgsRq>
</QBXML>
"""


def add_items_to_quickbooks(new_items: List[InventoryItem]) -> None:
    # Fetch current inventory to prevent duplicate IDs
    current_items = fetch_quickbooks_inventory()
    current_by_id = {i["record_id"] for i in current_items}
    items_to_add = []
    for item in new_items:
        if item.record_id in current_by_id:
            print(f"Warning: '{item.record_id}' already exists—skipping '{item.name}'.")
        else:
            items_to_add.append(item)

    if not items_to_add:
        print("No new items to add.")
        return

    qbxml_request = _build_items_add_qbxml(items_to_add)
    response_xml = _call_quickbooks(qbxml_request)
    root = ET.fromstring(response_xml)
    for add_rs in root.findall(".//ItemInventoryAddRs"):
        status_code = add_rs.attrib.get("statusCode", "")
        status_msg = add_rs.attrib.get("statusMessage", "")
        name = add_rs.findtext(".//ItemInventoryRet/Name", default="UNKNOWN")
        if status_code == "0":
            print(f"Add Success: {name}")
        else:
            print(f"Add Failed: {name} - {status_msg}")


if __name__ == "__main__":
    print("Reading inventory from QuickBooks ...")
    try:
        items = fetch_quickbooks_inventory()
        for item in items:
            print(f"{item['name']} (id={item['record_id']}, price={item['price']})")
    except Exception as e:
        print("Error while reading QuickBooks inventory:", e)

    print("\nAdding new inventory parts to QuickBooks ...")
    try:
        new_parts = [
            InventoryItem(
                record_id="TEST03", name="Gadget", price=10.99, source="quickbooks"
            ),
            InventoryItem(
                record_id="TEST04", name="Bolt", price=13.20, source="quickbooks"
            ),
            InventoryItem(
                record_id="TEST06", name="Nail", price=14.69, source="quickbooks"
            ),
        ]
        add_items_to_quickbooks(new_parts)
    except Exception as e:
        print("Error while adding inventory items:", e)
