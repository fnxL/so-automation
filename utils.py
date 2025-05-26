import win32com.client
import threading
from Logger import Logger
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import NamedStyle
from pywinauto import Application


def create_thread(target, args):
    thread = threading.Thread(target=target, args=args)
    thread.start()
    return thread


class PywinUtils:
    @staticmethod
    def handle_dialogs(app_title: str, window_title: str, ctrl: str, logger: Logger):
        try:
            app = Application().connect(title_re=app_title)
            dlg = app.window(title_re=window_title, class_name="#32770")
            dlg.wait("exists", timeout=50000)
            dlg[ctrl].click()
            logger.log(f"{window_title} dialog handled successfully.")
        except Exception as e:
            logger.error(f"Failed to handle dialog {window_title}: {e}")
            raise e

    @staticmethod
    def handle_sap_scripting_alert(logger: Logger):
        try:
            app = Application().connect(title_re="SAP Easy Access.*")
            alert = app.window(title="SAP Logon", class_name="#32770")
            alert.wait("exists", timeout=50000)
            alert["OK"].click()
            logger.log("SAP GUI Scripting alert handled successfully.")
        except Exception as e:
            logger.error(f"Failed to handle SAP alert: {e}")
            raise

    @staticmethod
    def handle_excel_macro_errors(logger: Logger):
        try:
            app = Application().connect(title_re=".*Excel", class_name="XLMAIN")
            dlg = app.window(title_re=".*Visual Basic", class_name="#32770")
            dlg.wait("exists", timeout=50000)
            error_description = dlg["Static"].texts()[0]
            dlg["End"].click()
            logger.error(f"Macro Error: {error_description}")
            return
        except Exception as e:
            logger.error(f"Failed to handle Excel macro error: {e}")
            raise e


def get_df_from_excel(path: str, sheet_name: str | int = 0, **kwargs) -> pd.DataFrame:
    df = pd.read_excel(io=path, engine="calamine", sheet_name=sheet_name, **kwargs)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.lower()
    return df


def format_number(ws: Worksheet, col: int):
    for row in ws.iter_cols(min_col=col, max_col=col):
        for cell in row:
            cell.number_format = "0"
    return


def format_date(ws: Worksheet, startcol=0, endcol=1, date_format="DD-MM-YYYY"):
    for column in ws.iter_cols(min_col=startcol, max_col=endcol):
        for cell in column:
            cell.style = NamedStyle(name="datetime", number_format=date_format)
