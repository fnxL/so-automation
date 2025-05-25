import os
from .pdf_processor_kohls import process_pdf
from utils import get_df_from_excel, format_number, format_date
from config import get_customer_config
from openpyxl import load_workbook
from Logger import Logger
from datetime import datetime
from MacroRunner import MacroRunner

NOTIFY_ADDRESS = """Li & Fung (Trading) Limited\n7/F, HK SPINNERS INDUSTRIAL BUILDING\nPhase I & II,\n800 CHEUNG SHA WAN ROAD,\nKOWLOON, HONGKONG\nAir8 Pte Ltd,\n3 Kallang Junction\n#05-02 Singapore 339266\n 2% commission to WUSA"""


class Kohls:
    def __init__(
        self, source_folder: str, mastersheet_path: str, macro_path: str, logger: Logger
    ):
        self.source_folder = source_folder
        self.macro_path = macro_path
        self.pis_df = get_df_from_excel(path=mastersheet_path, sheet_name="PIS")
        self.mastersheet_df = get_df_from_excel(path=mastersheet_path)
        self.macro_wb = load_workbook(filename=macro_path, keep_vba=True)
        self.macro_ws = self.macro_wb.worksheets[0]
        self.logger = logger
        self.customer_config = get_customer_config("Kohls Towel")

    def start(self):
        pdf_files = [
            f for f in os.listdir(self.source_folder) if f.lower().endswith(".pdf")
        ]
        if not pdf_files:
            warning_msg = f"No PDF files found in source folder: {self.source_folder}"
            self.logger.warn(warning_msg)

        self.logger.info(
            f"Found {len(pdf_files)} PDF files in source folder: {self.source_folder}"
        )

        for pdf_file in pdf_files:
            pdf_path = os.path.join(self.source_folder, pdf_file)
            self.logger.log(f"Processing PDF: {pdf_file}")
            self.create_macro(pdf_path=pdf_path)

        macro_filename_with_ext = self.macro_path.split(os.sep)[-1]

        macro_filename = "FILLED_" + macro_filename_with_ext
        macro_path = os.path.join(self.source_folder, macro_filename)

        format_number(
            self.macro_ws, 32
        )  # 32 = upc col idx, format=long number instead of scientific notation
        format_date(
            self.macro_ws, startcol=13, endcol=14, date_format="DD-MM-YYYY"
        )  # ship dates

        self.macro_ws[self.customer_config["source_folder_cell"]] = self.source_folder
        self.macro_wb.save(macro_path)

        self.logger.success(
            f"Macro file successfully filled and saved copy to {macro_path}",
        )
        MacroRunner(
            macro_path=macro_path,
            macro_name=self.customer_config["macro_name"],
            logger=self.logger,
        ).run()

    def create_macro(self, pdf_path):
        po_metadata, line_items = process_pdf(pdf_path=pdf_path)
        (
            po,
            port_of_shipment,
            channel_type,
            ship_start_date,
            ship_end_date,
        ) = (
            po_metadata["po"],
            po_metadata["port_of_shipment"],
            po_metadata["channel_type"],
            po_metadata["ship_start_date"],
            po_metadata["ship_end_date"],
        )
        # convert ship_start_date (str) and ship_end_date (str) which is the format of YYYY-mm-dd to datetime objects with the format "%d-%m-%Y"

        ship_start_date = datetime.strptime(ship_start_date, "%Y-%m-%d").strftime(
            "%d-%m-%Y"
        )
        ship_end_date = datetime.strptime(ship_end_date, "%Y-%m-%d").strftime(
            "%d-%m-%Y"
        )

        sub_channel_type = "PURE PLAY ECOM" if channel_type == "ECOM" else "NIL"
        notify = NOTIFY_ADDRESS if "notify" in po_metadata else "2% commission to WUSA"
        row_data = {}
        for row in line_items:
            _, _, _, qty, _, _, upc = row
            mastersheet_row = self.mastersheet_df.query(f"upc=={upc}")
            plant = mastersheet_row.iloc[0]["plant"]
            program_name = mastersheet_row.iloc[0]["program name"]
            sales_unit = mastersheet_row.iloc[0]["sales unit"]
            design = str(mastersheet_row.iloc[0]["design"]).lower().strip()

            new_po = po
            if design in self.customer_config["design_split"]:
                new_po = f"{po} {design}"
                if f"{design}_{sales_unit}" not in row_data:
                    row_data[f"{design}_{sales_unit}"] = []

            if sales_unit not in row_data:
                row_data[sales_unit] = []

            packing_type = "BULK" if channel_type == "RETAIL" else "ECOM"
            filtered_pis = self.pis_df[
                (self.pis_df["program name"] == program_name)
                & (self.pis_df["sales unit"] == sales_unit)
                & (self.pis_df["packing type"] == packing_type)
            ]

            pis_value = filtered_pis.iloc[0]["pis"] if not filtered_pis.empty else "N/A"
            f_part = filtered_pis.iloc[0]["f part"] if not filtered_pis.empty else "N/A"
            date_month = int(ship_start_date.split("-")[1])
            s_part = ""
            if plant == 2100:
                if packing_type == "BULK" and sales_unit == "PC":
                    s_part = f"ALT{date_month}"
                elif packing_type == "ECOM" and sales_unit == "PC":
                    s_part = f"ALT{date_month}"
                elif packing_type == "ECOM" and (
                    sales_unit == "12 PC SET" or sales_unit == "6 PC SET"
                ):
                    s_part = f"ALT{date_month + 12}"

            macro_row = [
                new_po,  # PO
                "",  # SO
                102083,  # SOLD TO PARTY
                102083,  # SHIP TO PARTY
                "W137",  # PAYMENT TERM
                "",  # INCO TERMS
                "",  # INCO TERM 2
                "",  # order reason
                100023,  # end customer
                channel_type,
                sub_channel_type,
                "REPLENISHMENT",  # order type
                ship_start_date,  # ship start date
                ship_end_date,  # ship cancel date
                port_of_shipment,  # port of shipment
                "NEW YORK",  # destinatoin
                "USA",  # country
                "NEW YORK",  # port of loading
                mastersheet_row.iloc[0]["material number"],  # matcode,
                qty,  # order qty
                "",  # amount
                mastersheet_row.iloc[0]["sort number"],  # TT sort no
                mastersheet_row.iloc[0]["shade name"],  # TT shade
                mastersheet_row.iloc[0]["set type"],  # TT set
                "",  # Embroidery code L
                "",  # sublistatic code
                s_part,  # TT packing type for S part
                f_part,  # TT packing type for F part
                "NA",  # Destination
                mastersheet_row.iloc[0]["yarn dyed matching"],  # yarn dyed
                mastersheet_row.iloc[0]["plant"],  # plant
                upc,  # customer material
                pis_value,  # PIS
                "",  # PO AVAIL DATE
                "Saluja Tirkey",
                notify,
                po,
                "pdf",  # PO file format
            ]
            if design in self.customer_config["design_split"]:
                row_data[f"{design}_{sales_unit}"].append(macro_row)
            else:
                row_data[sales_unit].append(macro_row)
        # loop through the row_data dictionary and append the rows to the macro worksheet

        for sales_unit, macro_rows in row_data.items():
            for macro_row in macro_rows:
                self.macro_ws.append(macro_row)
            # add a blank row after each sales unit
            self.macro_ws.append([""])
