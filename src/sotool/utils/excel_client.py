import win32com.client as win32
from loguru import logger


class ExcelClient:
    def __init__(self, excel_path: str, logger=logger):
        self.excel_path = excel_path
        self.logger = logger

    def open_excel(self):
        try:
            self.logger.info("Attempting to open Excel...")
            self.excel = win32.Dispatch("Excel.Application")
            self.excel.Visible = True
            self.workbook = self.excel.Workbooks.Open(self.excel_path)
            return self
        except Exception as e:
            self.logger.error(f"Failed to open Excel/workbook: {e}")
            raise e

    def run_macro(self, macro_name: str):
        try:
            self.logger.info(f"Running macro: {macro_name}")
            result = self.excel.Application.Run(macro_name)
            self.logger.success(
                f"Macro {macro_name} ran successfully. Result: {result}"
            )
        except Exception as e:
            self.logger.warning(f"Macro error: {e}")

    def copy_table(self, sheet=1):
        sheet = self.workbook.Sheets(sheet)
        sheet.UsedRange.Copy()
        return

    def cleanup(self):
        self.excel.DisplayAlerts = False
        self.workbook.Close(SaveChanges=True)
        self.excel.Quit()
        self.logger.info("Excel closed successfully.")
