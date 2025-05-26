import win32com.client
import threading
from Logger import Logger
from utils import PywinUtils
from SAPUtils import SAPUtils


class MacroRunner:
    def __init__(self, macro_path: str, macro_name: str, logger: Logger):
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

            # Create different threads here
            macro_error_thread = threading.Thread(
                target=PywinUtils.handle_excel_macro_errors, args=(self.logger,)
            )
            macro_error_thread.start()

            sap_alert_thread = threading.Thread(
                target=SAPUtils.handle_sap_scripting_alert, args=(self.logger,)
            )
            sap_alert_thread.start()

            excel.Application.Run(self.macro_name)
            self.logger.success(f"Macro {self.macro_name} ran successfully.")

        except Exception as e:
            self.logger.error(f"Macro Error: {e}")
            raise e

        finally:
            # Clean up
            excel.DisplayAlerts = False
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.logger.info("Excel closed successfully.")

            # Join threads here
            macro_error_thread.join(timeout=10)
            sap_alert_thread.join(timeout=10)
