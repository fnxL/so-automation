from loguru import logger
import sys


class Logger:
    def __init__(self, log_widget=None):
        self.logger = logger.bind(name="Logger")
        self.logger.remove()
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
        level_name = record["level"].name
        timestamp = record["time"].strftime("%I:%M %p")
        formatted_message = record["message"]
        classes = ""
        prefix = ""

        if level_name == "INFO":
            prefix = "[INFO] "
        elif level_name == "WARNING":
            classes = "text-orange-400"
            prefix = "[WARNING] "
        elif level_name == "ERROR":
            classes = "text-red-400"
            prefix = "[ERROR] "
        elif level_name == "SUCCESS":
            classes = "text-green-400"
            prefix = "[SUCCESS] "

        self.log_widget.push(
            f"{timestamp} {prefix}{formatted_message}".strip(), classes=classes
        )
