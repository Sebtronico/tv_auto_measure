from pyproj import Transformer

# Definir el sistema de referencia de entrada y salida
crs_origen = "ESRI:103599"  # Sistema del DEM
crs_destino = "EPSG:3857"   # Sistema de destino (usado en QGIS)

# Crear el transformador
transformer = Transformer.from_crs(crs_origen, crs_destino, always_xy=True)

# Coordenadas límite en ESRI:103599
bounds = {
    "left": 4280523.4007429285,  # Oeste
    "bottom": 1037468.1706524522,  # Sur
    "right": 5769172.745110866,  # Este
    "top": 2996969.6092570713  # Norte
}

# Convertir las coordenadas a EPSG:3857
west, south = transformer.transform(bounds["left"], bounds["bottom"])  # Oeste y Sur
east, north = transformer.transform(bounds["right"], bounds["top"])  # Este y Norte

# Imprimir los límites
print(f"Límite Oeste: {west}")
print(f"Límite Sur: {south}")
print(f"Límite Este: {east}")
print(f"Límite Norte: {north}")
