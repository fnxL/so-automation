from dataclasses import dataclass
from datetime import datetime


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
