import os
import sys
from pathlib import Path

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

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)