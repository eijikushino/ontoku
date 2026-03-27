import json
import os

CONFIG_FILE = "config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def update_config_value(key_path, value):
    config = load_config()
    current = config
    for key in key_path[:-1]:
        current = current.setdefault(key, {})
    current[key_path[-1]] = value
    save_config(config)
