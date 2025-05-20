"""
Punto de entrada principal para la aplicación de automatización de mediciones de TV.
Este script inicia la interfaz gráfica de usuario.
"""

import sys
import os

# Asegurar que el directorio principal esté en el path de Python
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.gui.gui import iniciar_aplicacion as gui_main

if __name__ == "__main__":
    gui_main()