import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet


def get_df_from_excel(path: str, sheet_name: str | None = None):
  if sheet_name:
    df = pd.read_excel(io=path, engine="calamine", sheet_name=sheet_name)
    df.columns = df.columns.str.lower()
    return df

  df = pd.read_excel(io=path, engine="calamine")
  df.columns = df.columns.str.lower()
  return df


def format_number(ws: Worksheet, col: int):
  for row in ws.iter_cols(min_col=col, max_col=col):
    for cell in row:
      cell.number_format = "0"
  return
