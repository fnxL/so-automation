from .kohls import Kohls, POData
from ...integrations import ExcelClient
from ..sap_dispatch_report import download_dispatch_reports


class KohlsTowel(Kohls):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def start(self, stop_after_create_macro: bool = False):
        pdf_files = self._get_pdf_files(self.source_folder)
        if not pdf_files:
            self.logger.error("No PDF files found in the source folder.")
            raise FileNotFoundError("No PDF files found in the source folder.")

        self.logger.info(f"Found {len(pdf_files)} PO files to process")

        for pdf_file in pdf_files:
            self._process_single_po(pdf_file)

        filled_macro_path = self._finalize_macro_file()

        if stop_after_create_macro:
            self.logger.success("Macro file created successfully. Stopping here.")
            return

        # Run Macro
        with ExcelClient(filled_macro_path, logger=self.logger) as excel:
            excel.run_macro(self.config["macro_name"])

        # Get Reports
        reports = download_dispatch_reports(
            macro_path=filled_macro_path,
            source_folder=self.source_folder,
            logger=self.logger,
        )
        # Create Mails
        self._create_draft_mails(reports)
        return True

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
