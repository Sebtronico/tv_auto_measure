"""
Pendiente: configurar la aplicación para usar en modo CLI, sin necesidad de usar la interfaz gráfica.
"""
import sys
import os

# Asegurar que el directorio principal esté en el path de Python
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core.MeasurementManager import MeasurementManager

def main():
    """Función principal para ejecutar la aplicación en modo CLI."""
    print("Iniciando automatización de mediciones de TV...")
    
    # Ejemplo de uso del MeasurementManager
    manager = MeasurementManager()
    # Configuración y ejecución de mediciones
    
    print("Proceso completado.")

if __name__ == "__main__":
    main()