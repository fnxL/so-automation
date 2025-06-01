from loguru import logger
from ..config import config
from ..utils import get_df_from_excel
from openpyxl import load_workbook
from datetime import datetime
import os


class MacroGenerator:
    def __init__(
        self,
        source_folder: str,
        customer_name: str,
        logger=logger,
        stop_after_create_macro: bool = False,
    ):
        self.customer_config = config.get(customer_name)
        self.source_folder = source_folder
        self.mastersheet_path = self.customer_config.get("mastersheet_path")
        self.macro_path = self.customer_config.get("macro_path")
        self.stop_after_create_macro = stop_after_create_macro
        self.logger = logger

        # Load excel files
        self.mastersheet_df = get_df_from_excel(path=self.mastersheet_path)
        pass
        self.macro_wb = load_workbook(filename=self.macro_path, keep_vba=True)
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
