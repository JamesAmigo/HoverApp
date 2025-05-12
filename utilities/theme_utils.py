import os
import sys
from pathlib import Path
import json
from utilities.resource_utils import get_config_dir, get_resource_path

# Use user's AppData or home directory
THEME_CONFIG_PATH = os.path.join(get_config_dir(), "theme_config.json")
RESOURCES_PATH = get_resource_path(r"Resources\Stylesheets")

def load_themes():
    themes = {}
    for filename in os.listdir(RESOURCES_PATH):
        if filename.endswith(".qss"):
            theme_name = filename.split(".")[0]
            full_path = os.path.join(RESOURCES_PATH, filename)
            themes[theme_name] = load_stylesheet(full_path)
    return themes

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