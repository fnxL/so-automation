import win32com.client
import threading
import time
from loguru import logger
from pywinauto import Application


class SAPUtils:
    @staticmethod
    def start_sap_alert_thread(logger=logger):
        sap_alert_thread = threading.Thread(
            target=SAPUtils.handle_sap_scripting_alert, args=(logger,)
        )
        sap_alert_thread.start()
        time.sleep(2)
        return sap_alert_thread

    @staticmethod
    def connect_to_sap(logger=logger):
        sap_alert_thread = threading.Thread(
            target=SAPUtils.handle_sap_scripting_alert, args=(logger,)
        )
        sap_alert_thread.start()

        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            if not sap_gui:
                logger.error("SAP GUI is not running.")
                raise Exception("SAP GUI is not running.")

            logger.success("Connected to SAP GUI successfully.")

            application = sap_gui.GetScriptingEngine
            if not application:
                logger.error("Failed to get SAP scripting engine.")
                raise Exception("Failed to get SAP scripting engine.")

            connection = application.Children(0)
            if not connection:
                logger.error("No active SAP connection found.")
                raise Exception("No active SAP connection found.")

            session = connection.Children(0)
            if not session:
                logger.error("No active SAP session found.")
                raise Exception("No active SAP session found.")

            logger.success("Connected to SAP session successfully.")

            sap_alert_thread.join(timeout=10)
            return session
        except Exception as e:
            logger.error(f"Error connecting to SAP: {e}")
            raise e

    @staticmethod
    def handle_sap_scripting_alert(logger=logger):
        try:
            app = Application().connect(path="saplogon.exe")
            alert = app.window(title="SAP Logon", class_name="#32770")
            alert.wait("exists", timeout=50000)
            alert["OK"].click()
            logger.info("SAP GUI Scripting alert handled successfully.")
        except Exception as e:
            logger.error(f"Error connecting to SAP GUI Window: {e}")
            logger.warning(
                "Ensure SAP GUI window is opened and logged in with your credentials."
            )
            raise e
