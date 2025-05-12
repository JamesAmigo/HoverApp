import os
import sys
import json
from PyQt5.QtWidgets import QApplication
from utilities.resource_utils import get_config_dir, get_resource_path

THEME_CONFIG_PATH = os.path.join(get_config_dir(), "theme_config.json")
RESOURCES_PATH = get_resource_path("Resources/Stylesheets")
_theme_cache = None  # Internal lazy cache

def load_stylesheet(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        print(f"Failed to load stylesheet {path}: {e}")
        return ""

def get_all_themes():
    global _theme_cache
    if _theme_cache is None:
        _theme_cache = {}
        for filename in os.listdir(RESOURCES_PATH):
            if filename.endswith(".qss"):
                theme_name = filename.split(".")[0]
                full_path = os.path.join(RESOURCES_PATH, filename)
                _theme_cache[theme_name] = load_stylesheet(full_path)
    return _theme_cache


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

def get_next_theme(current_theme):
    return {
        "light": "pink",
        "pink": "dark",
        "dark": "light"
    }.get(current_theme, "light")
def get_theme_icon(theme_name):
    icon_map = {
        "light": "🌸",
        "pink": "🌙",
        "dark": "💥"
    }
    return icon_map.get(theme_name, "🌸")

def apply_theme(theme_name):
    themes = get_all_themes()
    if theme_name in themes:
        QApplication.instance().setStyleSheet(themes[theme_name])
        save_theme_preference(theme_name)
        return theme_name
    return None
