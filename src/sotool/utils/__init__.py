import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from .excel_client import ExcelClient
from .outlook_client import OutlookClient
from .sap_utils import SAPUtils


def get_df_from_excel(path: str, sheet_name: str | int = 0, **kwargs) -> pd.DataFrame:
    df = pd.read_excel(io=path, engine="calamine", sheet_name=sheet_name, **kwargs)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.lower()
    return df


def format_number(ws: Worksheet, startcol=0, endcol=1, format="0"):
    for row in ws.iter_cols(min_col=startcol, max_col=endcol):
        for cell in row:
            cell.number_format = format
    return


__all__ = [
    "get_df_from_excel",
    "format_number",
    "OutlookClient",
    "SAPUtils",
    "ExcelClient",
]
