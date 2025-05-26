import ctypes
import os

def set_wallpaper(image_path):
    # Diese Funktion setzt das Wallpaper unter Windows
    SPI_SETDESKWALLPAPER = 20
    result = ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, 3)
    if result:
        print("Wallpaper wurde jesetzt.")
    else:
        print("Konnte Wallpaper nich setzen.")

# Hier den Pfad zu deinem Wunschbild eintragen
wallpaper_path = r"C:\wallpaper\green.jpg"

# Auf absolute Pfade prüfen
if os.path.exists(wallpaper_path):
    set_wallpaper(wallpaper_path)
else:
    print("Bild nich jefunden. Prüf mal den Pfad, Alter.")
