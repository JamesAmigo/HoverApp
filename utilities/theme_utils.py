import os
import sys
from pathlib import Path
import json
from utilities.resource_utils import get_config_dir

# Use user's AppData or home directory
THEME_CONFIG_PATH = os.path.join(get_config_dir(), "theme_config.json")

def load_stylesheet(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        print(f"Failed to load stylesheet {path}: {e}")
        return ""

def save_theme_preference(theme_name):
    try:
        with open(THEME_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"theme": theme_name}, f)
    except Exception as e:
        print("Error saving theme preference:", e)

def load_theme_preference():
    try:
        with open(THEME_CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("theme", "light")
    except:
        return "light"