# automation_logic.py
import os
from config import get_customer_config
from kohls.Kohls import Kohls


def run_customer_automation(customer_name, source_folder, log_callback=None):
  if log_callback:
    log_callback(f"Starting automation for '{customer_name}'...", "info")

  customer_config = get_customer_config(customer_name)
  if not customer_config:
    error_msg = f"Error: Configuration not found for customer '{customer_name}'."
    if log_callback:
      log_callback(error_msg, "error")
    raise ValueError(error_msg)

  mastersheet = [f for f in os.listdir(source_folder) if "mastersheet" in f.lower()]
  if not mastersheet:
    error_msg = "Error: Mastersheet not found in the given path."
    if log_callback:
      log_callback(error_msg, "error")
    return FileNotFoundError(error_msg)

  macro = [f for f in os.listdir(source_folder) if "macro" in f.lower()]
  if not macro:
    error_msg = "Error: Macro file not found in the given path."
    if log_callback:
      log_callback(error_msg, "error")
    return FileNotFoundError(error_msg)

  if log_callback:
    log_callback(f"Using mastersheet: {mastersheet[0]}", "warning")
    log_callback(f"Using macro template: {macro[0]}", "warning")
    log_callback(f"Processing: {mastersheet[0]}...", "info")

  mastersheet_path = os.path.join(source_folder, mastersheet[0])
  macro_path = os.path.join(source_folder, macro[0])

  if "kohls" in customer_name.lower():
    Kohls(
      source_folder=source_folder,
      macro_path=macro_path,
      mastersheet_path=mastersheet_path,
      logger=log_callback,
    ).start()
