from pywinauto import Application
import threading
import win32com.client
from Logger import Logger


# pwa_app = pywinauto.application.Application()
# w_handle = pywinauto.findwindows.find_windows(
#     title="Microsoft Visual Basic", class_name="#32770"
# )[0]
# window = pwa_app.window_(handle=w_handle)
# window.SetFocus()
# ctrl = window["&End"]
# ctrl.Click()


class MacroRunner:
    def __init__(self, macro_path: str, macro_name: str, logger: Logger):
        self.macro_path = macro_path
        self.macro_name = macro_name
        self.logger = logger

    def run(self):
        self.logger.info(f"Attempting to open Excel and Run Macro: {self.macro_name}")
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True

            self.logger.log(
                f"Excel started successfully, opening workbook: {self.macro_path}"
            )
            workbook = excel.Workbooks.Open(self.macro_path)

            self.logger.log(f"Running Macro: {self.macro_name}")
            # Start macro error thread
            error_thread = threading.Thread(target=self.handle_excel_macro_errors)
            error_thread.start()

            excel.Application.Run(self.macro_name)

            error_thread.join(timeout=10)

        except Exception as e:
            self.logger.error(f"Failed to start Excel: {e}")
            return
        finally:
            # Clean up
            excel.DisplayAlerts = False
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.logger.log("Excel closed successfully.")

    def handle_excel_macro_errors(self):
        app = Application().connect(title_re=".*Excel", class_name="XLMAIN")
        dlg = app.window(title_re=".*Visual Basic", class_name="#32770")
        dlg.wait("exists", timeout=50000)
        error_description = dlg["Static"].texts()[0]
        dlg["End"].click()
        self.logger.error(f"Macro Error: {error_description}")
        raise Exception(f"Macro Error: {error_description}")
