# Enlace de descarga del geotiff
#https://serviciosgeovisor.igac.gov.co:8080/Geovisor/descargas?cmd=download&token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxNjQ4MzYiLCJleHAiOjE3Mzc2NTk5NjksImp0aSI6InNlcnZpY2lvXzAtMTU5In0.nUpEh3C4LmdvOwv-v7LNbkDvtIEW-nhuZ6Cew_zt-2aUtJ0TGVhk7_OHf4ybb4ALzU40WjZt8S6uAd-UhZHWUw

import rasterio
import numpy as np
import matplotlib.pyplot as plt
from shapely.geometry import LineString
from pyproj import Transformer

# Ruta al archivo DEM
dem_file = "./resources/SRTM_30_Col1.tif"

# Coordenadas de inicio y fin (latitud, longitud)
start_coord = (4.302305556,	-73.74183333)  # Ejemplo: Bogotá (4.60971, -74.08175)
end_coord = (4.3067319, -72.1111886)    # Ejemplo: Cali (3.43722, -76.5225)

# Sistema de referencia espacial
with rasterio.open(dem_file) as dem:
    dem_crs = dem.crs  # Obtener CRS del DEM

# Transformador de coordenadas
transformer = Transformer.from_crs("EPSG:4326", dem_crs.to_string(), always_xy=True)

# Convertir coordenadas a sistema proyectado del DEM
start_coord_projected = transformer.transform(*start_coord[::-1])
end_coord_projected = transformer.transform(*end_coord[::-1])

# Número de puntos a lo largo del perfil
n_points = 500

# Crear una línea entre los puntos de inicio y fin
line = LineString([start_coord_projected, end_coord_projected])

# Generar puntos intermedios a lo largo de la línea
interpolated_points = [line.interpolate(dist, normalized=True).coords[0] for dist in np.linspace(0, 1, n_points)]

# Lista para guardar las elevaciones y distancias
elevations = []
distances = []

# Abrir el DEM y procesar punto por punto
with rasterio.open(dem_file) as dem:
    for i, (x, y) in enumerate(interpolated_points):
        # Convertir x/y a fila/columna en el DEM
        try:
            row, col = dem.index(x, y)

            # Validar que la fila y columna estén dentro de los límites del DEM
            if 0 <= row < dem.height and 0 <= col < dem.width:
                # Leer solo un píxel del DEM
                elevation = dem.read(1, window=rasterio.windows.Window(col, row, 1, 1))[0, 0]
                elevations.append(elevation)

                # Calcular distancia acumulada
                if i > 0:
                    prev_x, prev_y = interpolated_points[i - 1]
                    dist = np.sqrt((x - prev_x)**2 + (y - prev_y)**2) / 1000  # Convertir a km
                    distances.append(distances[-1] + dist if distances else dist)
                else:
                    distances.append(0)
            else:
                print(f"Punto fuera de límites: {x}, {y}")
                elevations.append(None)
        except Exception as e:
            print(f"Error procesando el punto ({x}, {y}): {e}")
            elevations.append(None)

# Graficar el perfil de elevación
elevations = np.array(elevations, dtype=float)
distances = np.array(distances, dtype=float)

plt.figure(figsize=(13, 4))
plt.plot(distances, np.nan_to_num(elevations), color='brown')
plt.fill_between(distances, np.nan_to_num(elevations), color='lightcoral', alpha=0.5)
plt.title("Perfil de Elevación")
plt.xlabel("Distancia (km)")
plt.ylabel("Elevación (m)")
plt.grid()
plt.xlim(min(distances), max(distances))
plt.ylim(0, max(elevations))
plt.savefig("perfil_elevacion.png", dpi=300, bbox_inches="tight", pad_inches=0.25)  # Guarda la imagen con alta resolución
print("Gráfico guardado como 'perfil_elevacion.png'")
