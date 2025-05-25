from datetime import datetime


class Logger:
    def __init__(self, log_widget):
        self.log_widget = log_widget

    def log(self, message):
        timestamp = datetime.now().strftime("%I:%M %p")
        self.log_widget.push(f"{timestamp} {message}")
        print(f"{timestamp} {message}")

    def info(self, message):
        timestamp = datetime.now().strftime("%I:%M %p")
        self.log_widget.push(f"{timestamp} {message}", classes="text-blue-400")
        print(f"{timestamp} [INFO] {message}")

    def warn(self, message):
        timestamp = datetime.now().strftime("%I:%M %p")
        self.log_widget.push(
            f"{timestamp} [WARNING] {message}", classes="text-orange-400"
        )
        print(f"{timestamp} [WARNING] {message}")

    def error(self, message):
        timestamp = datetime.now().strftime("%I:%M %p")

        self.log_widget.push(f"{timestamp} [ERROR] {message}", classes="text-red-400")
        print(f"{timestamp} [ERROR] {message}")

    def success(self, message):
        timestamp = datetime.now().strftime("%I:%M %p")

        self.log_widget.push(
            f"{timestamp} [SUCCESS] {message}", classes="text-green-400"
        )
        print(f"{timestamp} [SUCCESS] {message}")
