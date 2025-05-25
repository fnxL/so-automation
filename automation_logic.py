# automation_logic.py
import os
from kohls.Kohls import Kohls
from Logger import Logger


def run_customer_automation(customer_name, source_folder, logger: Logger):
    mastersheet = [f for f in os.listdir(source_folder) if "mastersheet" in f.lower()]
    if not mastersheet:
        error_msg = "Error: Mastersheet not found in the given path."
        logger.error(error_msg, "error")
        return FileNotFoundError(error_msg)

    macro = [f for f in os.listdir(source_folder) if "macro" in f.lower()]
    if not macro:
        error_msg = "Error: Macro file not found in the given path."
        logger.error(error_msg, "error")
        return FileNotFoundError(error_msg)

    logger.log(f"Using mastersheet: {mastersheet[0]}")
    logger.log(f"Using macro template: {macro[0]}")
    logger.log(f"Processing: {mastersheet[0]}")

    mastersheet_path = os.path.join(source_folder, mastersheet[0])
    macro_path = os.path.join(source_folder, macro[0])

    if "kohls" in customer_name.lower():
        Kohls(
            source_folder=source_folder,
            macro_path=macro_path,
            mastersheet_path=mastersheet_path,
            logger=logger,
        ).start()
