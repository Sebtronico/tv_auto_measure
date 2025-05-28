import sys
from pathlib import Path

def rpath(relative_path):
    """Devuelve la ruta absoluta para archivos dentro de resources/templates, etc."""
    # Cuando PyInstaller crea el .exe, guarda los archivos en una carpeta temporal (sys._MEIPASS)
    base_path = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    return base_path / relative_path