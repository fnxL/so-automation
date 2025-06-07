import win32com.client as win32
import win32gui
import win32con
import time
import pythoncom
from loguru import logger


class OutlookClientWin32:
    def __init__(self, logger=logger):
        self.logger = logger
        self.outlook = None
        self.namespace = None
        self.inbox = None

    def connect(self):
        pythoncom.CoInitialize()
        try:
            self.logger.info("Attempting to Open/Connect to Outlook...")
            try:
                self.outlook = win32.GetActiveObject("Outlook.Application")
                self.logger.info("Connected to existing Outlook instance")
            except Exception as e:
                self.outlook = win32.Dispatch("Outlook.Application")
                self.logger.info("Created a new Outlook instance")

            time.sleep(1)
            self._close_dialogs()

            self.namespace = self.outlook.GetNameSpace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)
            self.logger.success("Connected to Outlook successfully.")
            return self
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            raise e

    def _close_dialogs(self):
        """Attempt to find and close any open Outlook dialogs"""

        # Find windows with "Outlook" or "Microsoft Outlook" in the title
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "Outlook" in window_text and "dialog" in window_text.lower():
                    windows.append(hwnd)
                    self.logger.info(f"Found potential Outlook dialog: {window_text}")
            return True

        dialog_windows = []
        win32gui.EnumWindows(enum_windows_callback, dialog_windows)

        # Try to close found dialog windows
        for hwnd in dialog_windows:
            try:
                self.logger.info(
                    f"Attempting to close dialog: {win32gui.GetWindowText(hwnd)}"
                )
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception as e:
                self.logger.warning(f"Failed to close dialog: {e}")

    def create_mail_and_paste_from_clipboard(self, to="", cc="", subject="", body=""):
        self._close_dialogs()
        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        time.sleep(1)
        max_retires = 3
        for attempt in range(max_retires):
            try:
                mail.Display()
            except Exception as e:
                if attempt < max_retires - 1:
                    self.logger.warning(f"Display attempt {attempt + 1} failed: {e}")
                    self._close_dialogs()
                    time.sleep(2)

        time.sleep(2)

        word_editor = mail.GetInspector.WordEditor
        selection = word_editor.Application.Selection
        selection.EndKey(6)  # 6 = wdStory - move to the end of the document
        selection.TypeText("\n\n")
        selection.PasteAndFormat(16)
        self.logger.info("Draft email created successfully.")

    def disconnect(self):
        if self.outlook:
            self.outlook = None
            self.namespace = None
            self.inbox = None
            self.logger.info("Disconnected from Outlook successfully.")
            pythoncom.CoUninitialize()
