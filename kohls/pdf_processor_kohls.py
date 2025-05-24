# pdf_processor.py
import pdfplumber
import re
from typing import List

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


def remove_duplicates(data: List[List[str | None]]):
  unique_data = []
  for row in data:
    if row not in unique_data:
      unique_data.append(row)
  return unique_data


def remove_none(data: List[List[str | None]]):
  new_data = []
  for row in data:
    new_row = []
    for item in row:
      if item is not None:
        new_row.append(item)
    if new_row:
      new_data.append(new_row)
  return new_data


def clean_data(data: List[List[str | None]]):
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


def extract_text_from_pdf(pdf_path: str, page_range: tuple = None):
  """
  Extracts all text from a PDF file.
  """
  text = ""
  try:
    with pdfplumber.open(pdf_path) as pdf:
      pages = pdf.pages[slice(*page_range) if page_range else slice(len(pdf.pages))]
      for page in pages:
        text += page.extract_text(layout=True) + "\n"
  except Exception as e:
    raise IOError(f"Error extracting text from {pdf_path}: {e}")
  return text


def parse_table_from_pdf(pdf_path):
  tables_data = []
  try:
    with pdfplumber.open(pdf_path) as pdf:
      for page in pdf.pages:
        table = page.extract_table()
        if table:
          for row in table:
            tables_data.append(row)
    tables_data = remove_duplicates(tables_data)
    tables_data = remove_none(tables_data)
    tables_data = clean_data(tables_data)
  except Exception as e:
    raise IOError(f"Error parsing tables from {pdf_path}: {e}")
  return tables_data


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


def process_pdf(pdf_path: str):
  extracted_data = {}
  full_text = extract_text_from_pdf(pdf_path, (0, 1))

  extracted_data.update(extract_fields_from_text(full_text, PDF_PARSING_RULES))
  table_data = parse_table_from_pdf(pdf_path)

  return extracted_data, table_data
