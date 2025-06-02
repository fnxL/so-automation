from loguru import logger
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import subprocess
import time
import os


class OutlookClient:
    def __init__(self, logger=logger):
        self.logger = logger
        self.app = None
        self.main_window = None

    def _get_outlook_path(self):
        outlook_paths = [
            "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office15\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office14\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office13\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office12\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office11\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\root\\Office10\\OUTLOOK.EXE",
        ]
        for path in outlook_paths:
            if os.path.exists(path):
                return path

        raise Exception("Could not find Outlook executable.")

    def connect(self):
        try:
            self.logger.info("Attempting to open/connect to Outlook...")
            try:
                self.app = Application().connect(path="outlook.exe")
                self.logger.success("Connected to existing Outlook instance")
            except Exception as e:
                outlook_path = self._get_outlook_path()
                subprocess.Popen(outlook_path)
                time.sleep(5)  # Wait for Outlook to start
                self.app = Application().connect(path="outlook.exe")
                self.logger.info("Created a new Outlook instance")

            self.main_window = self.app.window(
                title_re=".*Outlook.*", class_name="rctrl_renwnd32"
            )
            if self.main_window.exists(timeout=60):
                self.main_window.set_focus()
                self.logger.success("Connected to Outlook successfully.")
                return self
            else:
                raise Exception("Could not find Outlook main window")

        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            raise e

    def create_mail_and_paste(self, to="", cc="", subject="", body_text=""):
        self.logger.info("Creating new email...")
        try:
            self.main_window.set_focus()
            send_keys("^n")
            time.sleep(2)

            email_window = self.app.window(title_re=".*Untitled - Message.*")
            if not email_window.exists(timeout=30):
                self.logger.error("Could not find email window")
                raise Exception("Email window not found")

            email_window.set_focus()
            email_window.ToEdit.set_text(to)
            email_window.CcEdit.set_text(cc)
            email_window.SubjectEdit.set_text(subject)

            email_window._WwG.click_input()  # message body

            send_keys(body_text, with_spaces=True, with_newlines=True)
            send_keys("{ENTER}{ENTER}")
            send_keys("^v")
            time.sleep(1)
            send_keys("^s")
            self.logger.info("Draft email saved.")
            self.logger.info("Closing email window.")
            time.sleep(2)
            send_keys("%{F4}")
            self.logger.info("Email window closed.")

        except Exception as e:
            self.logger.error(f"Failed to create email: {e}")
            raise e

    def disconnect(self):
        if self.app:
            self.app = None
            self.main_window = None
            self.logger.info("Disconnected from Outlook successfully.")
