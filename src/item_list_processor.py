from dataclasses import dataclass
from typing import Any
import win32com.client
from openpyxl import load_workbook

@dataclass
class Item:
    name: str
    sales_price: float
    item_id: str  # Manufacturer's Part Number

@dataclass
class ItemComparison:
    same_id_diff_data: list[tuple[str, str, float]]  # (excel_name, qb_name, excel_price, qb_price)
    only_in_excel: list[Item]
    only_in_qb: list[Item]
    matching_count: int

def read_items(file_path: str) -> list[Item]:
    workbook = load_workbook(file_path, read_only=True)
    sheet = workbook.active  # Or specify sheet if known
    items = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, sales_price, item_id = row[0], row[1], row[2]
        if name and sales_price is not None and item_id:
            try:
                items.append(Item(name=str(name).strip(),
                                  sales_price=float(sales_price),
                                  item_id=str(item_id).strip()))
            except (ValueError, TypeError):
                continue
    return items

def connect_to_quickbooks() -> Any:
    qb_app = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    qb_app.OpenConnection("", "Item List Import")
    session = qb_app.BeginSession("", 2)
    return qb_app, session

def get_qb_items() -> list[Item]:
    qb_app, session = None, None
    try:
        qb_app, session = connect_to_quickbooks()
        qbxml_query = """<?xml version="1.0"?>
<?qbxml version="13.0"?>
<QBXML>
    <QBXMLMsgsRq onError="continueOnError">
        <ItemInventoryQueryRq></ItemInventoryQueryRq>
    </QBXMLMsgsRq>
</QBXML>"""
        response = qb_app.ProcessRequest(session, qbxml_query)
        import xml.etree.ElementTree as ET
        root = ET.fromstring(response)
        items = []
        for item_ret in root.findall(".//ItemInventoryRet"):
            name = item_ret.findtext("Name")
            sales_price_str = item_ret.findtext("SalesPrice")
            item_id = item_ret.findtext("ManufacturerPartNumber")
            try:
                sales_price = float(sales_price_str) if sales_price_str else 0.0
                if name and item_id:
                    items.append(Item(name=name, sales_price=sales_price, item_id=item_id))
            except (ValueError, TypeError):
                continue
        return items
    finally:
        if qb_app and session:
            qb_app.EndSession(session)
            qb_app.CloseConnection()

def compare_items(excel_items: list[Item], qb_items: list[Item]) -> ItemComparison:
    qb_dict = {item.item_id: item for item in qb_items}
    excel_dict = {item.item_id: item for item in excel_items}

    same_id_diff_data = []
    only_in_excel = []
    only_in_qb = []
    matching_count = 0

    for item_id, excel_item in excel_dict.items():
        qb_item = qb_dict.get(item_id)
        if not qb_item:
            only_in_excel.append(excel_item)
        else:
            if (excel_item.name != qb_item.name) or (excel_item.sales_price != qb_item.sales_price):
                same_id_diff_data.append((excel_item.name, qb_item.name, excel_item.sales_price, qb_item.sales_price))
            else:
                matching_count += 1
    
    for item_id, qb_item in qb_dict.items():
        if item_id not in excel_dict:
            only_in_qb.append(qb_item)

    return ItemComparison(
        same_id_diff_data=same_id_diff_data,
        only_in_excel=only_in_excel,
        only_in_qb=only_in_qb,
        matching_count=matching_count,
    )

def create_items_batch_qbxml(items: list[Item]) -> str:
    requests = []
    for item in items:
        name = item.name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        price = f"{item.sales_price:.2f}"
        item_id = item.item_id.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        requests.append(f"""
        <ItemInventoryAddRq>
            <ItemInventoryAdd>
                <Name>{name}</Name>
                <SalesPrice>{price}</SalesPrice>
                <ManufacturerPartNumber>{item_id}</ManufacturerPartNumber>
                <IncomeAccountRef>
                    <FullName>Sales of Product Income</FullName>
                </IncomeAccountRef>
            </ItemInventoryAdd>
        </ItemInventoryAddRq>""")

    qbxml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
    <QBXMLMsgsRq onError="continueOnError">
{chr(10).join(requests)}
    </QBXMLMsgsRq>
</QBXML>"""
    return qbxml

def save_items_to_quickbooks(items: list[Item]) -> list[str]:
    if not items:
        return []
    qb_app, session = None, None
    try:
        qb_app, session = connect_to_quickbooks()
        qbxml = create_items_batch_qbxml(items)
        response = qb_app.ProcessRequest(session, qbxml)
        import xml.etree.ElementTree as ET
        root = ET.fromstring(response)
        created = []
        for add_rs in root.findall(".//ItemInventoryAddRs"):
            status_code = add_rs.get("statusCode")
            if status_code == "0":
                name = add_rs.findtext("Name")
                if name:
                    created.append(name)
            elif status_code == "3100":
                # Item already exists
                continue
            else:
                print(f"Warning: Failed to add item: {add_rs.get('statusMessage','Unknown error')}")
        return created
    finally:
        if qb_app and session:
            qb_app.EndSession(session)
            qb_app.CloseConnection()

def process_items(file_path: str) -> ItemComparison:
    excel_items = read_items(file_path)
    print(f"Found {len(excel_items)} items in Excel.")
    qb_items = get_qb_items()
    print(f"Found {len(qb_items)} items in QuickBooks.")

    comparison = compare_items(excel_items, qb_items)

    print(f"Matching items (same id, name and price): {comparison.matching_count}")
    if comparison.same_id_diff_data:
        print(f"Items with same ID but different data ({len(comparison.same_id_diff_data)}):")
        for e_name, q_name, e_price, q_price in comparison.same_id_diff_data:
            print(f"  Excel: {e_name}, QB: {q_name}, Price Excel: {e_price}, Price QB: {q_price}")
    if comparison.only_in_qb:
        print(f"Items only in QuickBooks ({len(comparison.only_in_qb)}):")
        for item in comparison.only_in_qb:
            print(f"  ID {item.item_id}: {item.name} Price: {item.sales_price}")
    if comparison.only_in_excel:
        print(f"Adding {len(comparison.only_in_excel)} new items to QuickBooks...")
        created = save_items_to_quickbooks(comparison.only_in_excel)
        print(f"Successfully added {len(created)} items.")
    else:
        print("No new items to add to QuickBooks.")
    return comparison
