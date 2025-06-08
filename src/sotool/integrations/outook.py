from loguru import logger
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import subprocess
import time
import os
import re


def _find_outlook_executable_path():
    possible_outlook_paths = [
        "C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office15\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office14\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office13\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office12\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office11\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\Office10\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office15\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office14\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office13\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office12\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office11\\OUTLOOK.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office10\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office15\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office14\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office13\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office12\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office11\\OUTLOOK.EXE",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office10\\OUTLOOK.EXE",
    ]
    for path in possible_outlook_paths:
        if os.path.exists(path):
            return path

    raise FileNotFoundError("Could not locate OUTLOOK.EXE via  common paths.")


class OutlookClient:
    """
    Usage:
    ```python
        try:
            with OutlookClient() as outlook:
                outlook.create_draft_mail(to="test@example.com", subject="UI Test")
        except Exception as e:
            logger.error(f"Outlook UI automation failed: {e}")
    ```
    """

    def __init__(self, logger=logger):
        self.logger = logger
        self.app = None
        self.main_window = None

    def __enter__(self):
        self.logger.info("Attempting to open/connect to Outlook...")
        try:
            self.app = Application().connect(path="outlook.exe")
            self.logger.success("Connected to existing Outlook instance")
        except Exception:
            self.logger.warning("Outlook not running. Attempting to start it...")
            outlook_path = _find_outlook_executable_path()
            subprocess.Popen(outlook_path)
            time.sleep(5)  # Wait for Outlook to start
            self.app = Application().connect(path="outlook.exe", timeout=60)
            self.logger.success("Started and connected to a new Outlook instance.")

        self.main_window = self.app.window(
            title_re=".*Outlook.*", class_name="rctrl_renwnd32"
        )
        self.main_window.wait("ready", timeout=60)

        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.logger.info("Releasing control of Outlook application.")
        self.app = None
        self.main_window = None

    def create_draft_mail(
        self, to="", cc="", subject="", body="", paste_from_clipboard=False
    ):
        if not self.main_window:
            raise RuntimeError("Outlook is not connected. Use within a 'with' block.")

        self.logger.info("Creating a new draft email.")
        self.main_window.set_focus()
        send_keys("^n")
        try:
            email_window = self.app.window(title_re=".*Untitled - Message.*")
            email_window.wait("exists", timeout=60)

            email_window.set_focus()
            email_window.ToEdit.set_text(to)
            email_window.CcEdit.set_text(cc)
            email_window.SubjectEdit.click_input()
            email_window.SubjectEdit.set_text(subject)
            email_window.type_keys("{TAB}")
            # update window
            email_window = self.app.window(title_re=f".*{re.escape(subject)}.*")
            email_window.wait("exists", timeout=60)
            email_window.child_window(class_name="_WwG").set_focus()

            if body:
                email_window.type_keys(body, with_spaces=True, with_newlines=True)

            if paste_from_clipboard:
                email_window.type_keys("{ENTER}{ENTER}^v")

            self.logger.info("Saving and closing draft email window...")
            email_window.type_keys("^s")
            time.sleep(2)
            email_window.close()
            self.logger.info("Draft email created and saved.")
        except Exception as e:
            self.logger.error(f"Failed to find or interact with the email window: {e}")
            self.logger.warning(
                "This may be due to slow system performance or unexpected dialogs."
            )
            raise RuntimeError("Could not complete email creation via UI.") from e
