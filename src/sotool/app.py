from .config import config
from loguru import logger
from .automations.kohls.kohls_towel import KohlsTowelAutomation
import time
import os


def validate_config(automation_name, logger=logger):
    customer_config = config.get(automation_name)
    if not customer_config:
        msg = f"No configuration found for {automation_name}"
        logger.error(msg)
        raise ValueError(f"No configuration found for {automation_name}")

    base_folder = customer_config.get("base_folder")
    if not base_folder:
        msg = "No base folder found in config"
        logger.error(msg)
        logger.info(config["base_folder"])
        raise ValueError("No base folder found in config")

    if not os.path.isdir(base_folder):
        msg = f"Base folder directory not found in file system: {base_folder}"
        logger.error(msg)
        raise FileNotFoundError(f"Base folder not found in file system: {base_folder}")

    # detect macro and mastersheet files
    mastersheet = [f for f in os.listdir(base_folder) if "mastersheet" in f.lower()]
    if not mastersheet:
        msg = f"Mastersheet not found in base folder: {base_folder}"
        logger.error(msg)
        raise FileNotFoundError(f"Mastersheet not found in base folder: {base_folder}")

    macro = [f for f in os.listdir(base_folder) if "macro" in f.lower()]
    if not macro:
        msg = f"Macro not found in base folder: {base_folder}"
        logger.error(msg)
        raise FileNotFoundError(f"Macro not found in base folder: {base_folder}")

    mastersheet_path = os.path.join(base_folder, mastersheet[0])
    macro_path = os.path.join(base_folder, macro[0])

    customer_config.update(
        {"mastersheet_path": mastersheet_path, "macro_path": macro_path}
    )

    return customer_config


def run_automation(
    automation_name,
    source_folder,
    stop_after_create_macro: bool = False,
    logger=logger,
):
    customer_config = validate_config(automation_name, logger)
    start_time = time.time()

    match automation_name.lower():
        case "kohls_towel":
            KohlsTowelAutomation(
                config=customer_config,
                source_folder=source_folder,
                logger=logger,
            ).start(stop_after_create_macro=stop_after_create_macro)

    end_time = time.time()
    total_time = end_time - start_time
    logger.success(
        f"Automation {automation_name} ran successfully. Time taken: {total_time:.2f} seconds."
    )
    return
