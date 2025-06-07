import win32com.client
import threading
import time
from loguru import logger
from pywinauto import Application

# TODO - implement SAP Custom Errors


class SAPConnector:
    """
    Usage:
    ```python
    try:
        with SAPConnector() as session:
           session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
           session.findById("wnd[0]").sendVKey(0)
    except Exception  as e:
        logger.error(f"Please ensure you are logged into SAP system.")
    ```
    """

    def __init__(self, logger=logger):
        self.logger = logger
        self.session = None

    def __enter__(self):
        self.connect()

        if not self.session:
            raise "Failed to establish a valid SAP session."

        return self.session

    def __exit__(self, exc_type, exc_value, traceback):
        self.logger.info("Disconnecting from SAP session.")
        self.session = None

    def connect(self):
        self.logger.info("Connecting to SAP GUI Scripting Engine...")
        alert_thread = start_sap_alert_thread(logger=self.logger)
        try:
            sap_gui_app = win32com.client.GetObject("SAPGUI")
            engine = sap_gui_app.GetScriptingEngine
            connection = engine.Children(0)
            self.session = connection.Children(0)
            if not self.session:
                logger.error("No active SAP session found.")
                raise Exception("No active SAP session found.")

            self.logger.success("Connected to SAP session successfully.")
            alert_thread.join(timeout=5)
        except Exception as e:
            self.logger.error(f"Error connecting to SAP: {e}")
            raise e


def start_sap_alert_thread(logger=logger):
    sap_alert_thread = threading.Thread(target=_handle_sap_alert, args=(logger,))
    sap_alert_thread.start()
    time.sleep(2)
    return sap_alert_thread


def _handle_sap_alert(logger=logger):
    try:
        app = Application().connect(path="saplogon.exe")
        alert = app.window(title="SAP Logon", class_name="#32770")
        alert.wait("exists", timeout=50000)
        alert["OK"].click()
        logger.success("SAP GUI Scripting alert handled successfully.")
    except Exception as e:
        logger.error(f"Error connecting to SAP GUI Window: {e}")
        logger.warning(
            "Ensure SAP GUI window is opened and logged in with your credentials."
        )
        raise e
