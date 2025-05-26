CUSTOMER_CONFIGS = {
    "Kohls Towel": {
        "display_message": "Please ensure all PO files are in .pdf format.\nMasterdata must be present in the first worksheet in Mastersheet excel file, followed by 'PIS' worksheet.\nThere must be only one sheet in the macro excel file.\nMacro name for this program must be 'vtowels' in order to run the macro. Ensure that running macro is enabled in your system",
        "macro_name": "vtowels",
        "source_folder_cell": "AK1",
        "design_split": ["abstract", "medal", "stripe"],
        "notify_address": "Li & Fung (Trading) Limited\n7/F, HK SPINNERS INDUSTRIAL BUILDING\nPhase I & II,\n800 CHEUNG SHA WAN ROAD,\nKOWLOON, HONGKONG\nAir8 Pte Ltd,\n3 Kallang Junction\n#05-02 Singapore 339266\n 2% commission to WUSA",
    },
    "Kohls Rugs": {
        "display_message": "Test",
        "macro_name": "vtowels",
        "source_folder_cell": "AK1",
        "design_split": ["abstract", "medal", "stripe"],
    },
    "Kohls Bedsheet": {
        "display_message": "Hello, this is a test message for Kohls Bedsheet.",
        "source_folder_cell": "AK1",
        "design_split": ["abstract", "medal", "stripe"],
    },
}


def get_customer_config(customer_name):
    """
    Retrieves the configuration for a specific SO Customer.
    """
    return CUSTOMER_CONFIGS.get(customer_name)
