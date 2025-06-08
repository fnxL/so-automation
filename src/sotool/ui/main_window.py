import sys
import ttkbootstrap as ttk
import threading
from ..logger import Logger
from ..config import config
from .dialogs import Dialog
from tkinter.filedialog import askdirectory
from ..app import run_automation
from ttkbootstrap.constants import *


class SOAutomation(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, padding=(20, 10))
        self.master = master
        self.pack(fill=BOTH, expand=YES)
        self.automation_display_names = [
            data["display_name"] for key, data in config.items() if key != "default"
        ]

        self.automation_map = {
            data["display_name"]: key
            for key, data in config.items()
            if key != "default"
        }
        self.configure_layout()
        self.create_select_automation()
        self.create_source_folder_selection()
        self.create_stop_automation_check()
        self.create_run_automation_button()
        self.create_automation_logs()

        self.master.withdraw()  # Hide the main window initially
        self.master.after(
            0, self.master.deiconify
        )  # Show it after all widgets are rendered

    def configure_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

    def create_automation_logs(self):
        frame = ttk.LabelFrame(self, text="Automation Output Logs", padding=10)
        frame.grid(row=5, column=0, rowspan=100, sticky="nsew", pady=5)
        self.log_text = ttk.Text(frame, state=DISABLED, wrap=CHAR, height=15)
        self.log_text.pack(fill=BOTH, expand=YES)

        log_scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        log_scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)

        self.log_text.tag_configure("info", foreground="white")
        self.log_text.tag_configure("success", foreground="#90ee90")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")

        self.logger = Logger(self._log_message).get_logger()

    def create_run_automation_button(self):
        # Run Automation Button
        self.run_automation_button = ttk.Button(
            self,
            text="Run Automation",
            bootstyle=SUCCESS,
            command=self._run_automation,
        )
        self.run_automation_button.grid(
            row=4, column=0, sticky="ew", pady=10, padx=(0, 10)
        )

    def create_stop_automation_check(self):
        self.stop_after_create_macro = ttk.BooleanVar(value=False)
        ttk.Checkbutton(
            self,
            text="Stop after creating macro",
            variable=self.stop_after_create_macro,
            bootstyle="round-toggle",
        ).grid(row=3, column=0, sticky="w", pady=5, padx=(0, 10))

    def create_source_folder_selection(self):
        frame = ttk.Frame(self)
        frame.grid(row=2, sticky=EW, pady=5)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)

        ttk.Label(frame, text="Select Source Folder:").grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )

        self.select_source_folder = ttk.StringVar()
        ttk.Entry(frame, textvariable=self.select_source_folder, state=READONLY).grid(
            row=1, column=0, sticky=EW, padx=(0, 10)
        )

        ttk.Button(frame, text="Browse", command=self._get_directory, width=6).grid(
            row=1, column=1, sticky=E
        )

    def create_select_automation(self):
        frame = ttk.Frame(self)
        frame.grid(row=0, sticky=EW, pady=5)
        frame.columnconfigure(0, weight=0)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Select Automation:").grid(row=0, column=0, sticky=SW)

        ttk.Button(
            frame, text="Info", command=self._show_automation_info, width=6
        ).grid(
            row=0,
            column=1,
            sticky=E,
        )
        self.select_automation = ttk.Combobox(
            frame, values=self.automation_display_names
        )
        self.select_automation.grid(
            row=1, column=0, sticky=EW, columnspan=2, pady=(10, 0)
        )
        if self.automation_display_names:
            self.select_automation.set(self.automation_display_names[0])

        self.select_automation.bind("<Button-1>", self._on_combobox_click)

    def _on_combobox_click(self, event):
        event.widget.event_generate("<Down>")

    def _show_automation_info(self):
        selected_display_name = self.select_automation.get()
        automation_info = self.automation_map.get(selected_display_name)
        if automation_info and "description" in automation_info:
            Dialog.show_info(
                title=f"Information: {selected_display_name}",
                message=automation_info["description"],
                parent=self.master,
            )
        else:
            Dialog.show_warning(
                title="Description not available",
                message="No description available for this automation.",
            )

    def _log_message(self, message, level="info"):
        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, f"{message}\n", level)
        self.log_text.see(END)  # Scroll to the end
        self.log_text.config(state=DISABLED)

    def _get_directory(self):
        folder = askdirectory()
        if folder:
            self.select_source_folder.set(folder)

    def _run_automation(self):
        selected_automation = self.select_automation.get()
        source_folder = self.select_source_folder.get()
        stop_after_create_macro = self.stop_after_create_macro.get()

        if not selected_automation:
            Dialog.show_error(
                message="Please select an automation.",
                title="Automation not selected",
                parent=self.master,
            )
            return

        if not source_folder:
            Dialog.show_error(
                message="Please select a source folde path that contains all required files.",
                title="Source folder path not selected",
                parent=self.master,
            )
            return

        self.logger.info(f"Running: {selected_automation}", "info")
        self.logger.info(f"Source Folder Path: {source_folder}", "info")

        automation_name = self.automation_map.get(selected_automation)
        thread = threading.Thread(
            target=run_automation,
            args=(automation_name, source_folder, stop_after_create_macro, self.logger),
        )
        thread.start()
