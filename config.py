CUSTOMER_CONFIGS = {
  "Kohls Towel": {
    "display_message": "Please ensure all PO files are in .pdf format.\n\nMasterdata must be present in the first worksheet in Mastersheet excel file, followed by 'PIS' worksheet.\n\nThere must be only one sheet in the macro excel file.\n\nMacro name for this program must be 'vtowels' in order to run the macro. Ensure that running macro is enabled in your system",
    "macro_name": "vtowels",
    "source_folder_cell": "AK1",
    "design_split": ["abstract", "medal", "stripe"],
  },
}


def get_customer_config(customer_name):
  """
  Retrieves the configuration for a specific SO Customer.
  """
  return CUSTOMER_CONFIGS.get(customer_name)
