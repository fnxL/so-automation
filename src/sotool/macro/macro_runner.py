from loguru import logger
from ..utils import ExcelClient, SAPUtils


class MacroRunner:
    @staticmethod
    def run(macro_path, macro_name, logger=logger):
        excel = ExcelClient(macro_path, logger).open_excel()
        sap_alert_thread = SAPUtils.start_sap_alert_thread(logger)
        excel.run_macro(macro_name)
        excel.cleanup()
        sap_alert_thread.join(timeout=10)
