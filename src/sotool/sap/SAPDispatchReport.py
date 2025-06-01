import win32clipboard
import time
import os
from datetime import datetime
from loguru import logger
from ..utils import SAPUtils, get_df_from_excel


class SAPDispatchReport:
    def __init__(self, macro_path: str, logger=logger):
        self.macro_path = macro_path
        self.source_folder = os.path.dirname(macro_path)
        self.session = SAPUtils.connect_to_sap(logger)
        self.logger = logger

    def run(self):
        reports = self._download_dispatch_reports()
        return reports

    def _download_dispatch_reports(self):
        so_list = self._get_so_list()
        result = []
        print(so_list)
        for plant in so_list.keys():
            self._copy_list_to_clipboard(so_list[plant])
            timestamp = datetime.now().strftime("%Y.%m.%d")
            file_name = os.path.join(f"DispatchReport_{plant}_{timestamp}.xlsx")
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/Nzsddr"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS").text = str(plant)
            self.session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
            self.session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press()
            self._copy_list_to_clipboard(so_list[plant])
            self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            shell = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            shell.contextMenu()
            shell.selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.source_folder
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

            file_path = os.path.join(self.source_folder, file_name)
            result.append(file_path)
        return result

    def _copy_list_to_clipboard(self, so_list):
        text = "\r\n".join(map(str, so_list))
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
        time.sleep(1)
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
