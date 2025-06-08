from .kohls import Kohls, POData

from ...integrations.excel import get_df_from_excel


class KohlsRugs(Kohls):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _get_pis_data(self, mastersheet_row, po_data):
        if self.pis_df is None:
            self.pis_df = get_df_from_excel(
                path=self._mastersheet_path, sheet_name="PIS"
            )

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
            "product_pac_type": filtered_pis.iloc[0]["product pac type"]
            if not filtered_pis.empty
            else "N/A",
        }

    def _create_macro_row(
        self,
        po_data: POData,
        qty: int,
        upc,
        mastersheet_row,
    ):
        pis_data = self._get_pis_data(mastersheet_row, po_data)
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
            "",  # INCO TERM 2
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
            mastersheet_row["sort number"],  # Rugs Sort No.
            mastersheet_row["shade name"],  # Shade No - /Washin
            mastersheet_row["printing shade no"],  # Printing Shade No
            "",  # product packing style
            pis_data["product_pac_type"],
            "",  # Production type
            mastersheet_row["set type"],  # SET-NO
            mastersheet_row["yarn dyed matching"],  # shade number - yd
            "NA",  # destination
            mastersheet_row["plant"],  # plant
            upc,  # customer material
            pis_data["pis"],  # PIS
            "",  # PO AVAIL DATE
            po_data.notify,
            po_data.po,  # PO file name
            "pdf",  # PO file format
        ]
