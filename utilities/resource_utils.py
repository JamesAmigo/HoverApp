import os
import sys
from pathlib import Path
import json

# Use user's AppData or home directory


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

def get_config_dir(app_name="excel_search_tool"):
    if sys.platform == "win32":
        base = os.getenv("APPDATA", Path.home())
    elif sys.platform == "darwin":
        base = os.path.join(Path.home(), "Library", "Application Support")
    else:
        base = os.getenv("XDG_CONFIG_HOME", os.path.join(Path.home(), ".config"))

    config_dir = os.path.join(base, app_name)
    os.makedirs(config_dir, exist_ok=True)
    return config_dir

THEME_CONFIG_PATH = os.path.join(get_config_dir(), "theme_config.json")



def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)