from ..macro import MacroGenerator, MacroRunner
from ..integrations import (
    format_number,
    ExcelClient,
    OutlookClient,
    apply_borders,
)
from ..sap import SAPDispatchReport
from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple
from .pdf_processor import process_pdf
from openpyxl import load_workbook
import time
import os


@dataclass
class POData:
    po: int | str
    port_of_shipment: str
    channel_type: str
    sub_channel_type: str
    ship_start_date: datetime
    ship_end_date: datetime
    packing_type: str
    notify: str


class KohlsMacroGenerator(MacroGenerator):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def start(self):
        pdf_files = self._get_pdf_files_in_source_folder()
        if not pdf_files:
            self.logger.error("No PDF files found in the source folder.")
            raise FileNotFoundError("No PDF files found in the source folder.")

        self.logger.info(f"Found {len(pdf_files)} PDF files to process")

        for pdf_file in pdf_files:
            self._process_single_po(pdf_file)

        filled_macro_path = self._finalize_macro_file()

        if self.stop_after_create_macro:
            self.logger.success("Macro file created successfully. Stopping here.")
            return

        MacroRunner.run(
            macro_path=filled_macro_path,
            macro_name=self.config["macro_name"],
            logger=self.logger,
        )

        self.reports = SAPDispatchReport(
            macro_path=filled_macro_path, logger=self.logger
        ).run()
        self._create_draft_mail()

        return True

    def _create_draft_mail(self):
        outlook_client = OutlookClient(logger=self.logger).connect()
        for report in self.reports:
            plant = str(report[0])
            report_path = report[1]

            # apply borders first
            wb = load_workbook(filename=report_path)
            ws = wb.active
            apply_borders(ws)
            wb.save(report_path)
            time.sleep(2)

            self.logger.info(f"Copying dispatch report to clipboard: {report_path}")
            excel = ExcelClient(
                report_path,
                logger=self.logger,
            ).open_excel()
            excel.copy_table()

            to = self.config["mail"][plant]["to"]
            cc = self.config["mail"][plant]["cc"]
            subject = self.config["mail"][plant]["subject"]
            body = self.config["mail"][plant]["body_template"]
            self.logger.info(f"Creating email for plant: {plant}")
            outlook_client.create_mail_and_paste(to, cc, subject, body)
            excel.cleanup()

        outlook_client.disconnect()
        self.logger.info("Draft emails created successfully.")

    def _parse_po_metadata(self, po_metadata: dict) -> POData:
        ship_start_date = self._parse_ship_date(po_metadata["ship_start_date"])
        ship_end_date = self._parse_ship_date(po_metadata["ship_end_date"])
        sub_channel_type = (
            "PURE PLAY ECOM" if po_metadata["channel_type"] == "ECOM" else "NIL"
        )
        packing_type = "BULK" if po_metadata["channel_type"] == "RETAIL" else "ECOM"
        notify = (
            self.config["notify_address"]
            if "notify" in po_metadata
            else "2% commission to WUSA"
        )

        return POData(
            po=po_metadata["po"],
            port_of_shipment=po_metadata["port_of_shipment"],
            channel_type=po_metadata["channel_type"],
            ship_start_date=ship_start_date,
            ship_end_date=ship_end_date,
            sub_channel_type=sub_channel_type,
            packing_type=packing_type,
            notify=notify,
        )

    def _prepare_row_data(self, po_data: POData, line_items: List[Tuple]):
        row_data = {}
        for row in line_items:
            _, _, _, qty, _, _, upc = row
            mastersheet_row = self._get_mastersheet_row(upc)
            # Prepare entire excel macro row
            macro_row = self._create_macro_row(
                po_data=po_data,
                qty=qty,
                upc=upc,
                mastersheet_row=mastersheet_row,
            )

            row_group = self._get_row_group_key(mastersheet_row)

            if row_group not in row_data:
                row_data[row_group] = []
            row_data[row_group].append(macro_row)
        return row_data

    def _get_row_group_key(self, mastersheet_row):
        sales_unit = mastersheet_row["sales unit"]
        design = str(mastersheet_row["design"]).lower().strip()

        if design in self.config["design_split"]:
            return f"{design}_{sales_unit}"
        return sales_unit

    def _get_adjusted_po(self, base_po: int | str, design: str):
        if design in self.config["design_split"]:
            return f"{base_po} {design}"
        return base_po

    def _get_pis_data(self, mastersheet_row, po_data: POData):
        sales_unit = mastersheet_row["sales unit"]
        program_name = mastersheet_row["program name"]
        packing_type = po_data.packing_type
        filtered_pis = self.pis_df[
            (self.pis_df["program name"] == program_name)
            & (self.pis_df["sales unit"] == sales_unit)
            & (self.pis_df["packing type"] == packing_type)
        ]

        return {
            "pis": filtered_pis.iloc[0]["pis"] if not filtered_pis.empty else "N/A",
            "f_part": filtered_pis.iloc[0]["f part"]
            if not filtered_pis.empty
            else "N/A",
        }

    def _get_s_part(
        self, plant: int, packing_type: str, sales_unit: str, ship_month: int
    ):
        if plant == 2100:
            if packing_type == "BULK" and sales_unit == "PC":
                return f"ALT{ship_month}"
            elif packing_type == "ECOM" and sales_unit == "PC":
                return f"ALT{ship_month}"
            elif packing_type == "ECOM" and (
                sales_unit == "12 PC SET" or sales_unit == "6 PC SET"
            ):
                return f"ALT{ship_month + 12}"
        return ""

    def _create_macro_row(
        self,
        po_data: POData,
        qty: int,
        upc,
        mastersheet_row,
    ):
        ship_month = po_data.ship_start_date.month
        pis_data = self._get_pis_data(mastersheet_row, po_data)
        s_part = self._get_s_part(
            plant=mastersheet_row["plant"],
            packing_type=po_data.packing_type,
            sales_unit=mastersheet_row["sales unit"],
            ship_month=ship_month,
        )

        # Adjust PO if design split is applicable
        new_po = self._get_adjusted_po(
            po_data.po, str(mastersheet_row["design"]).lower().strip()
        )

        return [
            new_po,  # PO
            "",  # SO
            102083,  # SOLD TO PARTY
            102083,  # SHIP TO PARTY
            "W137",  # PAYMENT TERM
            "",  # INCO TERMS
            "JNPT / MUNDRA",  # INCO TERM 2
            "",  # order reason
            100023,  # end customer
            po_data.channel_type,
            po_data.sub_channel_type,
            "REPLENISHMENT",  # order type
            po_data.ship_start_date,  # ship start date
            po_data.ship_end_date,  # ship cancel date
            po_data.port_of_shipment,  # port of shipment
            "NEW YORK",  # destinatoin
            "USA",  # country
            "NEW YORK",  # port of loading
            mastersheet_row["material number"],  # matcode,
            qty,  # order qty
            "",  # amount
            mastersheet_row["sort number"],  # TT sort no
            mastersheet_row["shade name"],  # TT shade
            mastersheet_row["set type"],  # TT set
            "",  # Embroidery code L
            "",  # sublistatic code
            s_part,  # TT packing type for S part
            pis_data["f_part"],  # TT packing type for F part
            "NA",  # Destination
            mastersheet_row["yarn dyed matching"],  # yarn dyed
            mastersheet_row["plant"],  # plant
            upc,  # customer material
            pis_data["pis"],  # PIS
            "",  # PO AVAIL DATE
            "Saluja Tirkey",
            po_data.notify,
            po_data.po,  # PO file name
            "pdf",  # PO file format
        ]

    def _process_single_po(self, pdf_file: str):
        pdf_path = os.path.join(self.source_folder, pdf_file)
        self.logger.info(f"Processing PDF: {pdf_file}")

        po_metadata, line_items = process_pdf(pdf_path=pdf_path)
        po_data = self._parse_po_metadata(po_metadata)
        row_data = self._prepare_row_data(po_data, line_items)
        self.logger.info("Parsed PO File")
        self.logger.info("Writing data to macro worksheet...")
        self._write_rows_to_worksheet(row_data)

    def _write_rows_to_worksheet(self, row_data):
        for group, rows in row_data.items():
            for row in rows:
                self.macro_ws.append(row)
            self.macro_ws.append([""])  # Add a blank row after each group

    def _finalize_macro_file(self):
        format_number(
            self.macro_ws, startcol=32, endcol=32, format="0"
        )  # UPC column long number format

        format_number(
            self.macro_ws, startcol=13, endcol=14, format="DD-MM-YYYY"
        )  # ship dates

        # Set source folder in macro file
        self.macro_ws[self.config["source_folder_cell"]] = self.source_folder

        # Save filled macro
        macro_filename = "FILLED_" + os.path.basename(self.macro_path)
        output_path = os.path.join(self.source_folder, macro_filename)
        self.macro_wb.save(output_path)
        self.logger.info(f'Macro file saved to "{output_path}"')
        return output_path
