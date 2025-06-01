import json
import os


def get_config_path():
    config_path = os.environ.get("SOTOOL_CONFIG")
    if config_path and os.path.isfile(config_path):
        return config_path

    home_config = os.path.join(os.path.expanduser("~"), "sotool.json")

    if os.path.isfile(home_config):
        return home_config

    if os.path.isfile("sotool.json"):
        return "sotool.json"

    return FileNotFoundError("sotool.json config file not found!")


def read_config():
    with open(get_config_path(), "r") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            print(f"sotool.config is not a valid json file: {e}")
            raise e


config = read_config()
