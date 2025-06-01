from loguru import logger
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import subprocess
import time


class OutlookClient:
    def __init__(self, logger=logger):
        self.logger = logger
        self.app = None
        self.main_window = None

    def connect(self):
        try:
            self.logger.info("Attempting to open/connect to Outlook...")
            try:
                self.app = Application().connect(path="OUTLOOK.EXE")
                self.logger.info("Connected to existing Outlook instance")
            except Exception as e:
                subprocess.Popen([self.outlook_path])
                time.sleep(5)  # Wait for Outlook to start
                self.app = Application().connect(path="OUTLOOK.EXE")
                self.logger.info("Created a new Outlook instance")

            self.main_window = self.app.window(
                title_re=".*Outlook.*", class_name="Olk Host"
            )
            if self.main_window.exists():
                self.main_window.set_focus()
                self.logger.success("Connected to Outlook successfully.")
                return self
            else:
                raise Exception("Could not find Outlook main window")

        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            raise e

    def create_mail_and_paste(self, to="", cc="", subject="", body=""):
        self.logger.info("Creating new email...")
        try:
            self.main_window.set_focus()
            send_keys("^n")
            time.sleep(2)

            email_window = self.app.window(title_re=".*- Message.*")
            if not email_window.exists():
                self.logger.error("Could not find email window")
                raise Exception("Email window not found")

            email_window.set_focus()

            to_field = email_window.child_window(title="To", control_type="Edit")
            to_field.set_text(to)

            cc_button = email_window.child_window(title="Cc", control_type="Button")
            if cc_button.exists():
                cc_button.click()
                time.sleep(1)

            cc_field = email_window.child_window(title="Cc", control_type="Edit")
            if cc_field.exists():
                cc_field.set_text(cc)

            # Subject line
            subject_field = email_window.child_window(
                title="Subject", control_type="Edit"
            )
            subject_field.set_text(subject)

            # Focus on the body
            email_window.set_focus()
            send_keys("{TAB}")  # Move to the body field

            # Type some text
            send_keys("This is a test mail{ENTER}{ENTER}")
            send_keys("^v")  # Ctrl+V to paste clipboard content
            self.logger.info("Pasting text from clipboard.")

            time.sleep(1)

            # Save the draft (Ctrl+S)
            email_window.set_focus()
            send_keys("^s")
            self.logger.info("Draft email saved.")
            self.logger.info("Closing email window.")
            time.sleep(2)
            email_window.set_focus()
            send_keys("%{F4}")
            self.logger.info("Email window closed.")
            time.sleep(1)

        except Exception as e:
            self.logger.error(f"Failed to create email: {e}")
            raise e

    def disconnect(self):
        if self.app:
            self.app = None
            self.main_window = None
            self.logger.info("Disconnected from Outlook successfully.")
