import os
import sys
import json
from utilities.resource_utils import get_config_dir

HEADER_CONFIG_PATH = os.path.join(get_config_dir(), "sheet_headers.json")

def load_header_config():
    try:
        with open(HEADER_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_header_config(config):
    try:
        with open(HEADER_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        print("Error saving header config:", e)

def get_header_index(filename, sheet_name):
    config = load_header_config()
    return config.get(filename, {}).get(sheet_name)

def set_header_index(filename, sheet_name, index):
    config = load_header_config()
    if filename not in config:
        config[filename] = {}
    config[filename][sheet_name] = index
    save_header_config(config)
