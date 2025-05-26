import json

def get_config() -> dict:
    try:
        with open("config.json", "r") as file:
            return json.load(file)
    except Exception as e:
        print(f"config.json is not a valid json: {e}")
        raise e


CUSTOMER_CONFIGS = get_config()
