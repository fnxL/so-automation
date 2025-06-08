# TODO - make it flexible to parse
import pdfplumber
import re
from typing import List
from datetime import datetime
from .po_data import POData

PDF_PARSING_RULES = {
    "po": r"Order Number.*\n\s*(\d+)",
    "port_of_shipment": r"FOB -\s*(\w+)",
    "channel_type": r"Order Indicator",
    "ship_window": r"Shipment Window.*\n.*(\d{4}-\d{2}-\d{2} / \d{4}-\d{2}-\d{2})",
    "notify": r"\bLI\s*&\s*FUNG\b",
}

TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_tolerance": 10,
    "min_words_vertical": 1,
    "min_words_horizontal": 1,
}

NOTIFY_ADDRESS = "Li & Fung (Trading) Limited\n7/F, HK SPINNERS INDUSTRIAL BUILDING\nPhase I & II,\n800 CHEUNG SHA WAN ROAD,\nKOWLOON, HONGKONG\nAir8 Pte Ltd,\n3 Kallang Junction\n#05-02 Singapore 339266\n2% commission to WUSA"


def extract_text_from_pdf(
    pdf_path: str,
    page_range: tuple = None,
    last_page_only: bool = False,
    from_end: int = None,
):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if last_page_only:
                page = pdf.pages[-1]
                text = page.extract_text(layout=True)
            elif from_end is not None:
                start_page = max(0, len(pdf.pages) - from_end)
                for i in range(start_page, len(pdf.pages)):
                    text += pdf.pages[i].extract_text(layout=True) + "\n"
            else:
                pages = pdf.pages[
                    slice(*page_range) if page_range else slice(len(pdf.pages))
                ]
                for page in pages:
                    text += page.extract_text(layout=True) + "\n"
    except Exception as e:
        raise IOError(f"Error extracting text from {pdf_path}: {e}")
    return text


def _remove_duplicates(data: List[List[str | None]]):
    unique_data = []
    for row in data:
        if row not in unique_data:
            unique_data.append(row)
    return unique_data


def _remove_none(data: List[List[str | None]]):
    new_data = []
    for row in data:
        new_row = []
        for item in row:
            if item is not None:
                new_row.append(item)
        if new_row:
            new_data.append(new_row)
    return new_data


def _clean_data(data: List[List[str | None]]):
    clean_data = []
    line_item = []
    for row in data:
        match = re.search(r"UPC/EAN \(GTIN\) (\d+)", row[0])
        if match:
            line_item.append(int(match.group(1)))
            clean_data.append(line_item)
        try:
            line_num = int(row[0])
            line_item = row
        except ValueError:
            continue

    for row in clean_data:
        row[3] = row[3].split(" ")[0]
        row[3] = int(row[3].replace(",", ""))

    return clean_data


def extract_fields_from_text(text, rules):
    extracted_data = {}
    for field, pattern in rules.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            if field == "ship_window":
                data = match.group(1).split(" / ")
                extracted_data["ship_start_date"] = data[0]
                extracted_data["ship_end_date"] = data[1]
            elif field == "port_of_shipment":
                data = match.group(1)
                extracted_data[field] = "MUNDRA" if "MUNDRA" in data else "JNPT"
            elif field == "channel_type":
                channel = match.group()
                if channel is not None:
                    extracted_data[field] = "ECOM"
            elif field == "po":
                extracted_data[field] = int(match.group(1).strip())
            elif field == "notify":
                extracted_data[field] = True
            else:
                extracted_data[field] = match.group(1).strip()
        else:
            if field == "channel_type":
                extracted_data[field] = "RETAIL"

    return extracted_data


def _parse_ship_date(date_str: str):
    return datetime.fromisoformat(date_str)


def get_total_qty_and_value(pdf_path: str):
    rules = {
        "total_qty": r"Total Item Qty\s+([\d,]+)",
        "order_value": r"Order Total\s+([\d,.]+)",
    }
    text = extract_text_from_pdf(pdf_path, from_end=2)

    data = extract_fields_from_text(text, rules)
    print(pdf_path)
    data["total_qty"] = int(data["total_qty"].replace(",", ""))
    data["order_value"] = float(data["order_value"].replace(",", ""))

    return data


def parse_po_metadata(pdf_path: str):
    text = extract_text_from_pdf(pdf_path, (0, 1))
    po_metadata = extract_fields_from_text(text, PDF_PARSING_RULES)

    ship_start_date = _parse_ship_date(po_metadata["ship_start_date"])
    ship_end_date = _parse_ship_date(po_metadata["ship_end_date"])
    sub_channel_type = (
        "PURE PLAY ECOM" if po_metadata["channel_type"] == "ECOM" else "NIL"
    )
    packing_type = "BULK" if po_metadata["channel_type"] == "RETAIL" else "ECOM"

    notify = NOTIFY_ADDRESS if "notify" in po_metadata else "2% commission to WUSA"

    return POData(
        po=po_metadata["po"],
        port_of_shipment=po_metadata["port_of_shipment"],
        channel_type=po_metadata["channel_type"],
        ship_start_date=ship_start_date,
        ship_end_date=ship_end_date,
        sub_channel_type=sub_channel_type,
        packing_type=packing_type,
        notify=notify,
    )


def parse_po_line_items(pdf_path):
    tables_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        tables_data.append(row)
        tables_data = _remove_duplicates(tables_data)
        tables_data = _remove_none(tables_data)
        tables_data = _clean_data(tables_data)
    except Exception as e:
        raise IOError(f"Error parsing tables from {pdf_path}: {e}")
    return tables_data
