import sys
from loguru import logger


class Logger:
    def __init__(self, log_widget=None):
        self.logger = logger.bind(name="Logger")
        self.logger.remove()
        if sys.stdout:
            self.logger.add(
                sys.stdout,
                format="<level>{time:%I:%M %p} [{level.name}]</level> {message}",
                level="INFO",
                colorize=True,
            )

        if log_widget:
            self.log_widget = log_widget
            self.logger.add(
                self._loguru_sink_to_widget,
                format="{message}",
                level="INFO",
                colorize=False,
            )

    def get_logger(self):
        return self.logger

    def _loguru_sink_to_widget(self, message):
        record = message.record
        level_name = record["level"].name.lower()
        timestamp = record["time"].strftime("%I:%M %p")
        formatted_message = record["message"]
        prefix = ""

        if level_name == "info":
            prefix = "[INFO] "
        elif level_name == "warning":
            prefix = "[WARNING] "
        elif level_name == "error":
            prefix = "[ERROR] "
        elif level_name == "success":
            prefix = "[SUCCESS] "

        message_format = f"{timestamp} {prefix}{formatted_message}".strip()

        self.log_widget(message=message_format, level=level_name)
