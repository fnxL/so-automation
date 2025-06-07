from ttkbootstrap.dialogs import MessageDialog
from ttkbootstrap.icons import Icon
import math


class CustomDialog(MessageDialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def _locate(self):
        if self._parent != None:
            # Define values for parent widget width & height.
            parent_width = self._parent.winfo_width()
            parent_height = self._parent.winfo_height()

            # Define values for center x & y coordinates of
            # parent widget.
            parent_center_x = math.trunc(self._parent.winfo_x() + (parent_width / 2))
            parent_center_y = math.trunc(self._parent.winfo_y() + (parent_height / 2))

            # Define values for MessageDialog instance width & height
            # based on its _toplevel instance.
            widget_width = self._toplevel.winfo_reqwidth()
            widget_height = self._toplevel.winfo_reqheight()

            # Calculate the center_x & center_y coordinates where
            # the MessageDialog will need to be situated in order
            # to be placed at the center of its parent widget.
            center_x = math.trunc(parent_center_x - (widget_width / 2))
            center_y = math.trunc(parent_center_y - (widget_height / 2))

            # Then, set the x & y coordinates used
            # by the geometry() method below to use
            # center_x & center_y instead.
            x = center_x
            y = center_y
        else:
            # Define values as in original method,
            # using top-left x & y coordinates of
            # root Tk instance.
            x = self._toplevel.master.winfo_rootx()
            y = self._toplevel.master.winfo_rooty()

        self._toplevel.geometry(f"+{x}+{y}")


class Dialog:
    @staticmethod
    def show_info(message, title=" ", parent=None, alert=False, **kwargs):
        dialog = CustomDialog(
            message=message,
            title=title,
            alert=alert,
            parent=parent,
            buttons=["OK:primary"],
            icon=Icon.info,
            localize=True,
        )
        if "position" in kwargs:
            position = kwargs.pop("position")
        else:
            position = "center"
        dialog.show(position)

    @staticmethod
    def show_error(message, title=" ", parent=None, alert=True, **kwargs):
        dialog = CustomDialog(
            message=message,
            title=title,
            parent=parent,
            buttons=["OK:primary"],
            icon=Icon.error,
            alert=alert,
            localize=True,
            **kwargs,
        )
        if "position" in kwargs:
            position = kwargs.pop("position")
        else:
            position = None
        dialog.show(position)

    @staticmethod
    def show_warning(message, title=" ", parent=None, alert=True, **kwargs):
        dialog = CustomDialog(
            message=message,
            title=title,
            parent=parent,
            buttons=["OK:primary"],
            icon=Icon.warning,
            alert=alert,
            localize=True,
            **kwargs,
        )
        if "position" in kwargs:
            position = kwargs.pop("position")
        else:
            position = None
        dialog.show(position)
