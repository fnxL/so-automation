from nicegui import app, ui
from config import CUSTOMER_CONFIGS
from Logger import Logger
from automation_logic import run_customer_automation
import webview
import threading


class AutomationGUI:
    def __init__(self):
        ui.dark_mode().enable()
        self.create_layout()
        ui.query(".nicegui-content").style(
            "box-sizing: border-box; padding: 1rem; height: 100vh;"
        )

    def create_layout(self):
        with ui.grid(columns="2fr 2fr").classes("w-full h-full"):
            with ui.column().classes("w-full"):
                with ui.card().classes(
                    "no-shadow w-full bg-neutral-800 rounded-lg border border-neutral-700"
                ):
                    self.create_customer_select()
                self.create_select_source_path()
                self.create_run_automation_button()
            with ui.column().classes(
                "w-full mb-0 pb-0 h-full overflow-auto box-border"
            ):
                self.create_automation_log()

    def create_run_automation_button(self):
        self.run_automation = ui.button(
            text="Run Automation",
            color="bg-green-900",
            on_click=self._handle_run_automation,
        ).classes(
            "w-full bg-green-600 text-white font-semibold rounded-lg px-4 py-2 transition-colors hover:bg-green-700"
        )

    def create_select_source_path(self):
        with ui.card().classes(
            "no-shadow w-full bg-neutral-800 rounded-lg border border-neutral-700"
        ):
            with ui.row().classes("items-center justify-between gap-2"):
                ui.icon("sym_r_folder").classes("text-lg text-green-400")
                ui.label("Select Source Path").classes(
                    "text-md font-semibold text-gray-200"
                )
            with ui.row().classes("flex items-center gap-2 w-full"):
                self.source_path_input = (
                    ui.input()
                    .classes(
                        "flex-1 bg-neutral-700 text-gray-200 rounded-xl px-2 py-0 border border-neutral-600 w-[70%]"
                    )
                    .props("borderless dense readonly")
                )
                self.select_button = ui.button(
                    "Browse", on_click=self._on_browse_button_click
                ).classes(
                    "text-gray-200 font-medium rounded-lg border border-gray-700 transition-colors flex items-center gap-2"
                )
            self.checkbox = ui.checkbox(text="Stop after creating macro sheet")

    def create_automation_log(self):
        with ui.card().classes(
            "no-shadow w-full h-full bg-neutral-800 rounded-lg border border-neutral-700 "
        ):
            ui.label("Automation Logs").classes("text-md font-semibold text-gray-200")
            self.log = ui.log().classes(
                "h-full bg-gray-950 border-gray-700 font-mono rounded-lg text-sm resize-none focus:ring-1 focus:ring-blue-500 text-wrap break-all overflow-y-auto"
            )
            self.logger = Logger(self.log)

    def create_customer_select(self):
        with (
            ui.dialog() as dialog,
            ui.card().classes(
                "no-shadow w-full p-3 bg-stone-800 rounded-lg border border-stone-700"
            ),
        ):
            self.info_label = (
                ui.textarea(
                    label="Information", value="Select a customer to see information."
                )
                .classes("text-sm text-gray-300 leading-relaxed overflow-y-auto w-full")
                .props("borderless readonly autogrow")
            )
            ui.button("Close", on_click=dialog.close)

        with ui.row().classes("flex items-center w-full justify-between gap-2"):
            ui.icon("sym_r_groups").classes("text-lg text-blue-400")
            ui.label("Select Customer").classes(
                "flex-1 text-md font-semibold text-gray-200"
            )
            with ui.button(icon="info", on_click=dialog.open).classes(
                "h-0 w-5 text-sm border rounded-xl bg-transparent text-gray-200"
            ):
                ui.tooltip("Customer Information").classes("text-xs")

        self.customer_select = (
            ui.select(
                options=list(CUSTOMER_CONFIGS.keys()),
                on_change=self._on_customer_selected,
            )
            .classes("w-full bg-neutral-700 text-gray-200 rounded-xl px-2")
            .props(
                'borderless transition-show="scale" transition-hide="scale" dense stack-label'
            )
        )

    def _on_customer_selected(self) -> None:
        self.current_customer = self.customer_select.value
        if self.current_customer:
            self._load_customer_data()
            self._update_info_display()
            self.logger.log(f"Selected Customer: {self.current_customer}")

    async def _on_browse_button_click(self):
        files = await app.native.main_window.create_file_dialog(
            dialog_type=webview.FOLDER_DIALOG, allow_multiple=False
        )
        if not files:
            self.logger.warn("Source path selection was cancelled.")

        self.source_path_input.value = files
        self.source_path_folder = files[0]
        self.logger.log(f"Selected Source Path: {files[0]}")

    def _load_customer_data(self):
        self.customer_config = CUSTOMER_CONFIGS.get(self.current_customer)
        if not self.customer_config:
            self.logger.error(
                f"Configuration not found for customer '{self.current_customer}'."
            )
            return

    def _update_info_display(self):
        message = self.customer_config.get(
            "display_message", "No information available."
        )
        self.info_label.value = message

    def _handle_run_automation(self):
        if not self.source_path_input.value:
            self.logger.error(
                "Please select a source folder path before running automation."
            )
            return

        if not self.current_customer:
            self.logger.error("Please select a customer before running automation.")
            return

        try:
            self.logger.warn(
                f"Starting automation for customer: {self.current_customer}"
            )
            automation_thread = threading.Thread(
                target=run_customer_automation,
                args=(
                    self.current_customer,
                    self.source_path_folder,
                    self.logger,
                    self.checkbox.value,
                ),
            )
            automation_thread.start()
        except Exception as e:
            self.logger.error(f"An error occurred while running automation: {e}")
            return


AutomationGUI()

app.native.start_args["debug"] = True
app.native.window_args["resizable"] = True
app.native.window_args["min_size"] = (800, 528)

ui.run(
    native=True,
    title="🚀 SO Automation Tool",
    dark=True,
    window_size=(1024, 650),
)
