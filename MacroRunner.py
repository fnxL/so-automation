from pywinauto import Application
import threading
import win32com.client
from Logger import Logger


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

            self.logger.log(
                f"Excel started successfully, opening workbook: {self.macro_path}"
            )
            workbook = excel.Workbooks.Open(self.macro_path)
        except Exception as e:
            self.logger.error(f"Failed to start Excel/Open workbook: {e}")
            raise e

        try:
            self.logger.log(f"Running Macro: {self.macro_name}")

            # Create different threads here
            error_thread = threading.Thread(target=self.handle_excel_macro_errors)
            error_thread.start()

            sap_alert_thread = threading.Thread(target=self.handle_sap_alert)
            sap_alert_thread.start()

            excel.Application.Run(self.macro_name)
            self.logger.success(f"Macro {self.macro_name} ran successfully.")

        except Exception as e:
            self.logger.error(f"Error running macro: {e}")
            raise e
        finally:
            # Clean up
            excel.DisplayAlerts = False
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.logger.log("Excel closed successfully.")

            # Join threads here
            error_thread.join(timeout=10)
            sap_alert_thread.join(timeout=10)

    def handle_excel_macro_errors(self):
        try:
            app = Application().connect(title_re=".*Excel", class_name="XLMAIN")
            dlg = app.window(title_re=".*Visual Basic", class_name="#32770")
            dlg.wait("exists", timeout=50000)
            error_description = dlg["Static"].texts()[0]
            dlg["End"].click()
            self.logger.error(f"Macro Error: {error_description}")
            return
        except Exception as e:
            self.logger.error(f"Failed to handle Excel macro error: {e}")
            raise e

    def handle_sap_alert(self):
        try:
            app = Application().connect(title_re="SAP Easy Access.*")
            alert = app.window(title="SAP Logon", class_name="#32770")
            alert.wait("exists", timeout=50000)
            alert["OK"].click()
            self.logger.log("SAP Alert handled successfully.")
        except Exception as e:
            self.logger.error(f"Failed to handle SAP alert: {e}")
            raise
