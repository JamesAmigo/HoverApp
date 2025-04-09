from PyQt5.QtGui import QFontDatabase
import os
from utilities.resource_utils import get_resource_path

def load_fonts_from_folder():
    fonts_dir = get_resource_path("Resources/Fonts")
    loaded_fonts = []

    if not os.path.exists(fonts_dir):
        print("Fonts folder not found:", fonts_dir)
        return loaded_fonts

    for filename in os.listdir(fonts_dir):
        if filename.lower().endswith((".ttf", ".otf")):
            font_path = os.path.join(fonts_dir, filename)
            font_id = QFontDatabase.addApplicationFont(font_path)
            if font_id != -1:
                families = QFontDatabase.applicationFontFamilies(font_id)
                if families:
                    print(f"Loaded font: {families[0]}")
                    loaded_fonts.append(families[0])
            else:
                print(f"Failed to load font: {filename}")
    
    return loaded_fonts
