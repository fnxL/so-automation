import time
import pywinauto
from pywinauto import Application
import threading
import win32com.client
import pythoncom
import win32process  # Added for getting process ID
import logging  # Added for logging

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

pwa_app = Application()


def handle_excel_macro_end() -> None:
    vba_dialog_title = "Microsoft Visual Basic"
    vba_dialog_class = "#32770"
    end_button_text = "&End"
    while True:
        try:
            # Use the connected pwa_app to find the VBA dialog
            # Using .wait('visible', timeout=10) for robustness
            vba_dialog = pwa_app.window(
                title=vba_dialog_title, class_name=vba_dialog_class, timeout=10
            )
            logging.info(f"Found VBA dialog: '{vba_dialog_title}'")

            vba_dialog.SetFocus()
            end_button = vba_dialog[end_button_text]
            end_button.Click()
            logging.info(f"Clicked '{end_button_text}' button in VBA dialog.")
            break
        except pywinauto.timings.TimeoutError:
            logging.warning(
                f"VBA dialog '{vba_dialog_title}' not found within timeout."
            )
        except pywinauto.findwindows.ElementNotFoundError:
            logging.warning(f"'{end_button_text}' button not found in VBA dialog.")
        except Exception as e:
            logging.error(
                f"An unexpected error occurred while handling VBA dialog: {e}"
            )
            raise


def run_excel_macro(path: str, macro_name: str):
    """
    Runs an Excel macro and handles the VBA dialog that may appear afterwards.

    Args:
        path (str): The full path to the Excel workbook.
        macro_name (str): The name of the macro to run.
    """
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        workbook = excel.Workbooks.Open(path)
        logging.info(f"Opened Excel workbook: {path}")

        # Connect pywinauto to the Excel process
        # Get the process ID from the Excel application's main window handle
        _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        pwa_app.connect(process=excel_pid)
        logging.info(f"Connected pywinauto to Excel process (PID: {excel_pid}).")

        # Start the alert handler thread
        alert_thread = threading.Thread(target=handle_excel_macro_end, daemon=True)
        alert_thread.start()
        logging.info(f"Started thread to handle Excel macro end dialog.")

        excel.Application.Run(macro_name)
        logging.info(f"Executed Excel macro: {macro_name}")

        alert_thread.join(timeout=15)  # Increased timeout for thread to finish
        if alert_thread.is_alive():
            logging.warning("Alert handler thread did not finish within timeout.")

    except Exception as e:
        logging.error(f"An error occurred during macro execution: {e}", exc_info=True)
        raise
    finally:
        # Clean up
        if excel:
            logging.info("Cleaning up Excel application.")
            excel.DisplayAlerts = False
            if workbook:
                try:
                    workbook.Close(SaveChanges=True)
                    logging.info("Workbook closed.")
                except Exception as e:
                    logging.warning(f"Error closing workbook: {e}")
            excel.Quit()
            logging.info("Excel application quit.")
