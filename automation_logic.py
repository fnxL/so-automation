# automation_logic.py
import os
from kohls.Kohls import KohlsMacroGenerator
from Logger import Logger


def run_customer_automation(
    customer_name, source_folder, logger: Logger, stopAfterMacro: bool = False
):
    mastersheet = [f for f in os.listdir(source_folder) if "mastersheet" in f.lower()]
    if not mastersheet:
        error_msg = "Error: Mastersheet not found in the given path."
        logger.error(error_msg)
        return FileNotFoundError(error_msg)

    macro = [f for f in os.listdir(source_folder) if "macro" in f.lower()]
    if not macro:
        error_msg = "Error: Macro file not found in the given path."
        logger.error(error_msg)
        return FileNotFoundError(error_msg)

    logger.info(f"Using mastersheet: {mastersheet[0]}")
    logger.info(f"Using macro template: {macro[0]}")
    logger.info(f"Processing: {mastersheet[0]}")

    mastersheet_path = os.path.join(source_folder, mastersheet[0])
    macro_path = os.path.join(source_folder, macro[0])

    if "kohls" in customer_name.lower():
        KohlsMacroGenerator(
            source_folder=source_folder,
            macro_path=macro_path,
            mastersheet_path=mastersheet_path,
            logger=logger,
            customer_name=customer_name,
            stopAfterMacro=stopAfterMacro,
        ).start()
