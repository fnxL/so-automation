import sys

sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
import ttkbootstrap as ttk
from .ui.main_window import SOAutomation


def main():
    app = ttk.Window(title="SO Automation Tool", themename="darkly", size=(800, 600))
    app.place_window_center()
    SOAutomation(app)
    app.mainloop()
