import json
import os


def get_config_path():
    config_path = os.environ.get("SOTOOL_CONFIG")
    if config_path and os.path.isfile(config_path):
        return config_path

    if os.path.isfile("sotool.json"):
        return "sotool.json"

    home_config = os.path.join(os.path.expanduser("~"), "sotool.json")
    if os.path.isfile(home_config):
        return home_config

    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    if os.path.isdir(downloads):
        config_path = os.path.join(downloads, "sotool.json")
        if os.path.isfile(config_path):
            return config_path

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.isdir(desktop):
        config_path = os.path.join(desktop, "sotool.json")
        if os.path.isfile(config_path):
            return config_path

    raise FileNotFoundError("sotool.json config file not found!")


def read_config():
    config_path = get_config_path()
    print(f"Reading config from {config_path}")
    with open(config_path, "r") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            print(f"sotool.config is not a valid json file: {e}")
            raise e


config = read_config()
