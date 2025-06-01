from loguru import logger
import win32com.client
from ..utils import SAPUtils


class MacroRunner:
    def __init__(self, macro_path: str, macro_name: str, logger=logger):
        self.macro_path = macro_path
        self.macro_name = macro_name
        self.logger = logger

    def run(self):
        self.logger.info(f"Attempting to open Excel and Run Macro: {self.macro_name}")
        excel = None
        workbook = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True

            self.logger.info(
                f"Excel started successfully, opening workbook: {self.macro_path}"
            )
            workbook = excel.Workbooks.Open(self.macro_path)
        except Exception as e:
            self.logger.error(f"Failed to start Excel/Open workbook: {e}")
            raise e

        try:
            self.logger.info(f"Running Macro: {self.macro_name}")
            sap_alert_thread = SAPUtils.start_sap_alert_thread(self.logger)
            result = excel.Application.Run(self.macro_name)
            self.logger.success(
                f"Macro {self.macro_name} ran successfully. Result: {result}"
            )

        except Exception as e:
            self.logger.warning(f"Macro Error: {e}")
        finally:
            excel.DisplayAlerts = False
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.logger.info("Excel closed successfully.")
            sap_alert_thread.join(timeout=10)
