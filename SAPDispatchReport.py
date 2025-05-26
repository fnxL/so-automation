import win32com.client
import win32clipboard
import time
import os
import pandas as pd
from datetime import datetime
from Logger import Logger
from utils import get_df_from_excel


class SAPDispatchReport:
    def __init__(self, macro_path: str, logger: Logger):
        self.macro_path = macro_path
        self.logger = logger
        self.source_folder = os.path.dirname(macro_path)

    def run(self):
        self._connect_to_sap()
        reports = self._download_dispatch_reports()
        return reports

    def _connect_to_sap(self):
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            if not sap_gui:
                self.logger.error("SAP GUI is not running.")
                raise Exception("SAP GUI is not running.")

            self.logger.success("Connected to SAP GUI successfully.")

            application = sap_gui.GetScriptingEngine()
            if not application:
                self.logger.error("Failed to get SAP scripting engine.")
                raise Exception("Failed to get SAP scripting engine.")

            connection = application.Children(0)
            if not connection:
                self.logger.error("No active SAP connection found.")
                raise Exception("No active SAP connection found.")

            session = connection.Children(0)
            if not session:
                self.logger.error("No active SAP session found.")
                raise Exception("No active SAP session found.")

            self.logger.success("Connected to SAP session successfully.")
            self.session = session
        except Exception as e:
            self.logger.error(f"Error connecting to SAP: {e}")
            raise e

    def _download_dispatch_reports(self):
        so_list = self._get_so_list()
        result = []
        for plant in so_list.keys():
            self._copy_list_to_clipboard(so_list[plant])
            timestamp = datetime.now().strftime("%Y.%m.%d")
            file_name = os.path.join(
                self.source_folder, f"DispatchReport_{plant}_{timestamp}.xlsx"
            )
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "zsddr"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS").text = str(plant)
            self.session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
            self.session.findById("wnd[1]/tbar[0]/btn[24]").press
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press
            self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
            self.session.findById(
                "wnd[0]/usr/cntlGRID1/shellcont/shell"
            ).selectContextMenuItem = "&XXL"
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.source_folder
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press
            result.append(file_name)
        return result

    def _copy_list_to_clipboard(self, so_list):
        text = "\n".join(map(str, so_list))
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
        return text

    def _get_so_list(self):
        df = get_df_from_excel(path=self.macro_path, sheet_name=0, header=1)
        required_columns = ["so#", "plant"]

        if not all(col in df.columns for col in required_columns):
            self.logger.error(f"Missing required columns: {required_columns}")
            raise ValueError(f"Missing required columns: {required_columns}")

        unique_tuples = (
            df[required_columns]
            .dropna()
            .drop_duplicates()
            .apply(tuple, axis=1)
            .to_list()
        )

        cleaned = [
            (int(so), int(plant))
            for so, plant in unique_tuples
            if not any(isinstance(x, str) for x in (so, plant))
            and float(plant).is_integer()
        ]

        result = {}
        for so, plant in cleaned:
            if plant not in result:
                result[plant] = []
            result[plant].append(so)

        return result
