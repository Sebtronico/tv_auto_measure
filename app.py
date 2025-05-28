import logging
import sys
import os
from datetime import datetime

def setup_logging() -> None:
    """
    Crea la carpeta 'logs' junto al ejecutable (o al proyecto cuando es script),
    configura un archivo log con marca de tiempo y redirige stdout / stderr
    al log cuando el programa está 'frozen' (empaquetado con PyInstaller).
    """
    # Carpeta base donde se guardarán los logs
    if getattr(sys, "frozen", False):           # Estamos corriendo como .exe
        base_path = os.path.dirname(sys.executable)
    else:                                       # Estamos corriendo como script
        base_path = os.path.abspath(os.path.dirname(__file__))

    log_dir = os.path.join(base_path, "logs")
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(
        log_dir,
        datetime.now().strftime("log_%Y-%m-%d_%H-%M-%S.log")
    )

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")]
    )

    # Redirigir stdout y stderr solo cuando estamos en modo ejecutable
    if getattr(sys, "frozen", False):
        class StreamToLogger:
            def __init__(self, logger, level):
                self.logger = logger
                self.level = level

            def write(self, message):
                for line in message.rstrip().splitlines():
                    self.logger.log(self.level, line.rstrip())

            def flush(self):
                pass

        sys.stdout = StreamToLogger(logging.getLogger("STDOUT"), logging.INFO)
        sys.stderr = StreamToLogger(logging.getLogger("STDERR"), logging.ERROR)

    # Captura global de excepciones no manejadas
    def global_exception_handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            # Permite Ctrl+C sin traza extensa
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logging.critical(
            "Excepción no controlada",
            exc_info=(exc_type, exc_value, exc_traceback)
        )

    sys.excepthook = global_exception_handler

# Llamar lo antes posible para capturar todo
setup_logging()
if hasattr(sys, "_MEIPASS"):
    # Configurar PROJ_LIB
    for root, _, files in os.walk(sys._MEIPASS):
        if "proj.db" in files:
            proj_path = root
            os.environ["PROJ_LIB"] = proj_path
            logging.debug("PROJ_LIB establecido en %s", proj_path)
            try:
                from pyproj import datadir
                datadir.set_data_dir(proj_path)
            except Exception:
                logging.exception("Error configurando pyproj.datadir")
            break

    # Configurar GDAL_DATA
    for root, _, files in os.walk(sys._MEIPASS):
        if "gdalvrt.xsd" in files:
            os.environ["GDAL_DATA"] = root
            logging.debug("GDAL_DATA establecido en %s", root)
            break

# Se asegura que el directorio principal del proyecto esté en el PATH
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

try:
    from src.gui.gui import iniciar_aplicacion as gui_main
except Exception:
    logging.exception("No se pudo importar la GUI")
    raise  # Re-lanzar para que PyInstaller muestre error en modo desarrollo
if __name__ == "__main__":
    try:
        gui_main()
    except Exception:
        logging.exception("Excepción no controlada durante la ejecución de la GUI")
        raise
