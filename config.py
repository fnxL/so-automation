MASTERSHEET_PATH = "mastersheet.xlsx"
MACRO_TEMPLATE_PATH = "macro.xlsb"

CUSTOMER_CONFIGS = {
  "Kohls Towel": {
    "display_message": "Please ensure all PO files are in .pdf format.",
    "mastersheet_columns": {
      "program name": None,
      "upc": None,
      "plant": None,
      "material number": None,
      "sort number": None,
      "shade name": None,
      "set type": None,
      "yarn dyed matching": None,
      "pis": None,
      "sales unit": None,
    },
    "macro_population_procedures": {
      "invoice_cell": "B2",
      "date_cell": "B3",
      "items_start_row": 10,
      "items_column_mapping": {
        "Item": "A",
        "Quantity": "B",
        "Price": "C",
        "Amount": "D",
      },
      "final_total_cell": "D50",
    },
    "expected_pdf_fields": ["invoice_number", "date", "items_table"],
    "output_mapping": {
      "invoice_number": "Invoice Number",
      "date": "Invoice Date",
      "items_table": "Line Items",
    },
  },
}


def get_customer_config(customer_name):
  """
  Retrieves the configuration for a specific SO Customer.
  """
  return CUSTOMER_CONFIGS.get(customer_name)
