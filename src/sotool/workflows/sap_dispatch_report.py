import win32clipboard
import time
import os
from datetime import datetime
from loguru import logger
from ..integrations.excel import get_df_from_excel, ExcelClient
from ..integrations import SAPConnector


def _copy_list_to_clipboard(self, so_list):
    text = "\r\n".join(map(str, so_list))
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
    finally:
        win32clipboard.CloseClipboard()
    time.sleep(1)
    return text


def _get_so_list_from_macro(macro_path: str):
    df = get_df_from_excel(path=macro_path, sheet_name=0, header=1)
    required_columns = ["so#", "plant"]

    if not all(col in df.columns for col in required_columns):
        logger.error(f"Missing required columns: {required_columns}")
        raise ValueError(f"Missing required columns: {required_columns}")

    unique_tuples = (
        df[required_columns].dropna().drop_duplicates().apply(tuple, axis=1).to_list()
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


def download_dispatch_reports(macro_path: str, source_folder: str, logger=logger):
    logger.info("Downloading SAP Dispatch Reports...")
    so_list = _get_so_list_from_macro(macro_path)
    downloaded_reports = []
    with SAPConnector() as session:
        for plant in so_list.keys():
            _copy_list_to_clipboard(so_list[plant])
            timestamp = datetime.now().strftime("%Y.%m.%d")
            file_name = os.path.join(f"DispatchReport_{plant}_{timestamp}.xlsx")
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/Nzsddr"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtS_WERKS").text = str(plant)
            session.findById("wnd[0]/usr/ctxtS_WERKS").caretPosition = 4
            session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press()
            _copy_list_to_clipboard(so_list[plant])
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            shell = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            shell.contextMenu()
            shell.selectContextMenuItem("&XXL")
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = source_folder
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            file_path = os.path.join(source_folder, file_name)
            downloaded_reports.append((plant, file_path))
            time.sleep(5)
            with ExcelClient(logger=logger) as excel:
                excel.find_and_close_workbook(title_contains="DispatchReport")
    logger.success(f"Finished downloading {len(downloaded_reports)} dispatch reports.")
    return downloaded_reports
