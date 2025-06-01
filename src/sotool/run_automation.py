from .config import config
from loguru import logger
from .kohls import KohlsMacroGenerator
import time
import os


def validate_config(automation_name, logger=logger):
    automation_config = config.get(automation_name)
    if not automation_config:
        msg = f"No configuration found for {automation_name}"
        logger.error(msg)
        raise ValueError(f"No configuration found for {automation_name}")

    if not os.path.isfile(automation_config["mastersheet_path"]):
        msg = f"Mastersheet not found: {automation_config['mastersheet_path']}"
        logger.error(msg)
        raise FileNotFoundError(
            f"Mastersheet not found: {automation_config['mastersheet']}"
        )

    if not os.path.isfile(automation_config["macro_path"]):
        msg = f"Macro not found: {automation_config['macro_path']}"
        raise FileNotFoundError(f"Macro not found: {automation_config['macro_path']}")


def run_automation(
    automation_name,
    source_folder,
    stop_after_create_macro: bool = False,
    logger=logger,
):
    validate_config(automation_name, logger)
    start_time = time.time()

    if "kohls" in automation_name.lower():
        KohlsMacroGenerator(
            source_folder=source_folder,
            customer_name=automation_name,
            stop_after_create_macro=stop_after_create_macro,
        ).start()

    end_time = time.time()
    total_time = end_time - start_time
    logger.success(
        f"Automation {automation_name} ran successfully. Time taken: {total_time:.2f} seconds."
    )
    return
