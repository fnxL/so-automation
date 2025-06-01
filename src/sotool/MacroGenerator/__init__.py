from loguru import Logger
from ..config import config
from ..utils import Utils
from openpyxl import load_workbook
from datetime import datetime
import os


class MacroGenerator:
    def __init__(
        self,
        source_folder: str,
        mastersheet_path: str,
        macro_path: str,
        customer_name: str,
        logger: Logger = None,
        stop_after_create_macro: bool = False,
    ):
        self.source_folder = source_folder
        self.mastersheet_path = mastersheet_path
        self.macro_path = macro_path
        self.logger = logger
        self.stop_after_create_macro = stop_after_create_macro
        self.customer_config = config.get(customer_name)

        # Load excel files
        self.mastersheet_df = Utils.get_df_from_excel(path=mastersheet_path)
        pass
        self.macro_wb = load_workbook(filename=macro_path, keep_vba=True)
        self.macro_ws = self.macro_wb.worksheets[0]

    def _get_pdf_files_in_source_folder(self) -> list[str]:
        return [f for f in os.listdir(self.source_folder) if f.lower().endswith(".pdf")]

    def _parse_ship_date(self, date_str: str):
        return datetime.fromisoformat(date_str)

    def _get_mastersheet_row(self, upc):
        result = self.mastersheet_df.query(f"upc=={upc}")
        if result.empty:
            raise ValueError(f"No mastersheet row found for UPC: {upc}")
        return result.iloc[0]
