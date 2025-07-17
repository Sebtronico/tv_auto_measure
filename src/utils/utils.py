import sys
from pathlib import Path

def rpath(relative_path):
    """Devuelve la ruta absoluta para archivos dentro de resources/templates, etc."""
    # Cuando PyInstaller crea el .exe, guarda los archivos en una carpeta temporal (sys._MEIPASS)
    if getattr(sys, "frozen", False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).parent.parent.parent
    return base_path / relative_path