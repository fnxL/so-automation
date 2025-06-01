from loguru import logger
from ..utils import SAPUtils, ExcelUtils


class MacroRunner:
    def __init__(self, macro_path: str, macro_name: str, logger=logger):
        self.macro_path = macro_path
        self.macro_name = macro_name
        self.logger = logger

    def run(self):
        excel = ExcelUtils(self.macro_path, logger).open_excel()
        self.logger.info(f"Attempting to open Excel and Run Macro: {self.macro_name}")

        sap_alert_thread = SAPUtils.start_sap_alert_thread(self.logger)
        excel.run_macro(self.macro_name)
        excel.cleanup()
        sap_alert_thread.join(timeout=10)
