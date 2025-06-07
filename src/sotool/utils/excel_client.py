import win32com.client as win32
from pywinauto.application import Application
from loguru import logger


class ExcelClient:
    def __init__(self, excel_path: str = "", logger=logger):
        self.excel_path = excel_path
        self.logger = logger

    def open_excel(self):
        if not self.excel_path:
            raise ValueError("Excel path not provided")

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
            raise e

    def copy_table(self, sheet=1):
        sheet = self.workbook.Sheets(sheet)
        sheet.UsedRange.Copy()
        return

    @staticmethod
    def close_workbook(workbook_title_contains: str, logger=logger):
        try:
            app = Application(backend="uia").connect(
                title_re=f".*{workbook_title_contains}.*"
            )
            app.top_window().close()
            logger.info(f"{workbook_title_contains}: Workbook closed successfully.")
        except Exception as e:
            logger.error(f"Failed to close workbook/no workbook found: {e}")
            raise e

    @staticmethod
    def close_workbook_win32(workbook_title_contains: str, logger=logger):
        excel = None
        try:
            excel = win32.GetActiveObject("Excel.Application")
            logger.success("Connected to DispatchReport Excel instance")
        except Exception as e:
            logger.error(f"Failed to connect to Excel/No Excel instance found: {e}")
            raise e

        try:
            if excel.Workbooks.Count == 0:
                logger.info("No workbooks are currently open")
                return False

            dispatch_workbooks = []
            for i in range(1, excel.Workbooks.Count + 1):
                wb = excel.Workbooks(i)
                if workbook_title_contains in wb.Name:
                    dispatch_workbooks.append(wb)
                    logger.info(f"Found matching workbook: {wb.Name}")

            # Close found workbooks
            for wb in dispatch_workbooks:
                try:
                    wb_name = wb.Name
                    excel.DisplayAlerts = False
                    wb.Close(SaveChanges=False)
                    excel.DisplayAlerts = True
                    logger.success(f"Successfully closed workbook: {wb_name}")
                except Exception as e:
                    logger.error(f"Error closing workbook {wb.Name}: {e}")

            return len(dispatch_workbooks) > 0

        except Exception as e:
            logger.error(f"Error while searching for workbooks: {e}")
            return False
        finally:
            if excel:
                excel = None

    def cleanup(self):
        self.excel.DisplayAlerts = False
        self.workbook.Close(SaveChanges=True)
        self.excel.Quit()
        self.logger.info("Excel closed successfully.")
