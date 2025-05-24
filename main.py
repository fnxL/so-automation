import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os
import threading
import sv_ttk

# Backend modules
from automation_logic import run_customer_automation
from config import CUSTOMER_CONFIGS, get_customer_config


class AutomationToolGUI:
    def __init__(self, master):
        self.master = master
        master.title("SO Automation Tool")
        # master.geometry("800x600")  # Removed fixed size to allow resizing

        # Configure grid for responsiveness
        master.grid_rowconfigure(0, weight=0)
        master.grid_rowconfigure(1, weight=0)
        master.grid_rowconfigure(2, weight=0)  # For the customer message
        master.grid_rowconfigure(3, weight=0)  # No longer used
        master.grid_rowconfigure(4, weight=1)  # Text area for logs gets all extra space
        master.grid_columnconfigure(0, weight=1)
        master.grid_columnconfigure(1, weight=1)

        # --- SO Customer Selection ---
        self.customer_frame = ttk.LabelFrame(master, text="SO Customer Selection")
        self.customer_frame.grid(
            row=0, column=0, columnspan=2, padx=10, pady=5, sticky="ew"
        )
        self.customer_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(self.customer_frame, text="Select Customer:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.so_customer_var = tk.StringVar()
        self.so_customers = list(
            CUSTOMER_CONFIGS.keys()
        )  # Dynamically get customer list
        self.customer_combobox = ttk.Combobox(
            self.customer_frame,
            textvariable=self.so_customer_var,
            values=self.so_customers,
            state="readonly",
        )
        self.customer_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.customer_combobox.set(self.so_customers[0])  # Set default
        self.customer_combobox.bind(
            "<<ComboboxSelected>>", self.update_customer_message
        )
        self.customer_combobox.bind("<Return>", lambda e: self.folder_entry.focus())

        # Customer-specific message display
        self.customer_message_var = tk.StringVar()
        self.customer_message_label = ttk.Label(
            self.customer_frame,
            textvariable=self.customer_message_var,
            wraplength=700,
            justify=tk.LEFT,
            foreground="orange",
        )
        self.customer_message_label.grid(
            row=1, column=0, columnspan=2, padx=5, pady=2, sticky="w"
        )
        self.update_customer_message()  # Set initial message

        # --- Source Folder Selection ---
        self.folder_frame = ttk.LabelFrame(
            master, text="Select path of the folder with required files"
        )
        self.folder_frame.grid(
            row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew"
        )  # Adjusted row to 2
        self.folder_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(self.folder_frame, text="Source Folder:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.source_folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(
            self.folder_frame, textvariable=self.source_folder_path, width=50
        )
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.folder_entry.bind("<Return>", lambda e: self.run_automation())
        self.browse_button = ttk.Button(
            self.folder_frame, text="Browse", command=self.browse_folder
        )
        self.browse_button.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.browse_button.bind("<Return>", lambda e: self.browse_folder())
        self.browse_button.bind("<space>", lambda e: self.browse_folder())

        # --- Run Automation Button ---
        self.run_button = ttk.Button(
            master, text="Run Automation", command=self.run_automation
        )
        self.run_button.grid(
            row=2, column=2, padx=10, pady=10, sticky="e"
        )  # Moved to row 2, column 2 (right side of folder frame)
        self.run_button.bind("<Return>", lambda e: self.run_automation())

        # --- Progress and Feedback Display ---
        self.log_frame = ttk.LabelFrame(master, text="Automation Log")
        self.log_frame.grid(
            row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew"
        )  # Adjusted row to 4
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            self.log_frame, wrap=tk.WORD, width=80, height=15
        )  # Reduced height from 20 to 15 lines
        self.log_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.log_text.config(state="disabled")  # Make it read-only

        # Configure tags for colored logging
        self.log_text.tag_config("info", foreground="white")
        self.log_text.tag_config("success", foreground="green")
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("warning", foreground="yellow")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.source_folder_path.set(folder_selected)
            self.log_message(f"Source folder selected: {folder_selected}", "info")

    def log_message(self, message, message_type="info"):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n", message_type)
        self.log_text.see(tk.END)  # Auto-scroll to the end
        self.log_text.config(state="disabled")

    def update_customer_message(self, event=None):
        selected_customer = self.so_customer_var.get()
        customer_config = get_customer_config(selected_customer)
        if customer_config and "display_message" in customer_config:
            self.customer_message_var.set(customer_config["display_message"])
        else:
            self.customer_message_var.set("")

    def run_automation(self):
        selected_customer = self.so_customer_var.get()
        source_folder = self.source_folder_path.get()

        if not source_folder or not os.path.isdir(source_folder):
            self.log_message("Error: Please select a valid source folder.", "error")
            return

        self.log_message(
            f"Initiating automation for customer: {selected_customer}", "info"
        )
        self.log_message(f"Source folder: {source_folder}", "info")
        self.run_button.config(state="disabled")  # Disable button during automation

        # Run automation in a separate thread to keep GUI responsive
        self.automation_thread = threading.Thread(
            target=self._perform_automation_task,
            args=(selected_customer, source_folder),
        )
        self.automation_thread.start()

    def _perform_automation_task(self, customer, folder):
        try:
            self.log_message("Automation started...", "info")
            error = run_customer_automation(customer, folder, self.log_message)
            if error:
                self.log_message(f"Automation process failed: {error}", "error")
        except Exception as e:
            self.log_message(f"An error occurred: {e}", "error")
            self.log_message("Automation process failed.", "error")
        finally:
            self.master.after(
                100, lambda: self.run_button.config(state="normal")
            )  # Re-enable button in main thread


def main():
    root = tk.Tk()
    app = AutomationToolGUI(root)
    sv_ttk.set_theme("dark")
    # Set initial focus and tab order
    app.customer_combobox.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
