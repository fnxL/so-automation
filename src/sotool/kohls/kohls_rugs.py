from .kohls_macro_generator import KohlsMacroGenerator


class KohlsRugs(KohlsMacroGenerator):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _get_pis_data(self, mastersheet_row, po_data):
        # Remove f_part
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

    def _get_row_group_key(self, mastersheet_row):
        # Remove design
        return mastersheet_row["sales unit"]

    def _create_macro_row(self, po_data, qty, upc, mastersheet_row):
        pis_data = self._get_pis_data(mastersheet_row, po_data)
        return [
            po_data.po,  # PO
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
            mastersheet_row["material number"],  # mat code
            qty,  # order qty
            "",  # amount
            mastersheet_row["sort number"],  # TT sort no
            mastersheet_row["shade name"],  # TT shade
            mastersheet_row["printing shade no"],
            "",  # product packing style
            pis_data["product_pac_type"],
            "",  # production type
            mastersheet_row["set type"],  # Set No.
            mastersheet_row["yarn dyed matching"],
            "NA",  # destination
            mastersheet_row["plant"],
            upc,  # customer material
            pis_data["pis"],  # PIS
            "",  # PO AVAIL DATE
            "Saluja Tirkey",  # AC HOLDER
            po_data.notify,
            po_data.po,  # PO file name
            "pdf",  # PO file format
        ]
