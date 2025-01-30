import matplotlib.pyplot as plt
import geopandas as gpd
import rasterio
from rasterio.plot import show
from shapely.geometry import Point, LineString

# Coordenadas del punto de medición y estaciones
punto_medicion = (-72.1111886, 4.3067319)  # Coordenadas del punto de medición (lon, lat)
estaciones = [
    {"nombre": "Estación Azul", "coordenadas": (-73.74183333, 4.302305556)},  # Coordenadas estación azul
    {"nombre": "Estación Roja", "coordenadas": (-72.3886527777778, 5.37683333333333)},  # Coordenadas estación roja
]

# Crear un GeoDataFrame para las líneas
geometries = [
    LineString([punto_medicion, estacion["coordenadas"]]) for estacion in estaciones
]
lineas = gpd.GeoDataFrame(
    {"color": ["blue", "red"], "nombre": ["Línea Azul", "Línea Roja"]},
    geometry=geometries,
    crs="EPSG:4326",  # Sistema de coordenadas geográficas
)

# Crear un GeoDataFrame para el punto de medición
punto = gpd.GeoDataFrame(
    {"nombre": ["Punto de Medición"]},
    geometry=[Point(punto_medicion)],
    crs="EPSG:4326",
)

# Cargar el archivo raster del mapa base
raster_path = "./resources/Colombia_Satelital.tif"  # Ruta al archivo raster descargado
with rasterio.open(raster_path) as src:
    # Reproyectar las capas al CRS del raster
    lineas = lineas.to_crs(src.crs)
    punto = punto.to_crs(src.crs)

    # Recortar la región de interés automáticamente
    all_coords_proj = [
        punto.geometry[0].coords[0]
    ] + [lineas.geometry[idx].coords[1] for idx in range(len(estaciones))]
    lons, lats = zip(*all_coords_proj)

    margin = 10000  # Margen para los límites en las unidades del raster (metros)
    min_x, max_x = min(lons) - margin, max(lons) + margin
    min_y, max_y = min(lats) - margin, max(lats) + margin

    margin_factor = 0.2  # Factor del 10% para añadir márgenes
    x_min, x_max = min(lons), max(lons)
    y_min, y_max = min(lats), max(lats)
    margin_x = (x_max - x_min) * margin_factor
    margin_y = (y_max - y_min) * margin_factor

    # Crear la ventana (bounds -> ventana en el espacio del raster)
    window = rasterio.windows.from_bounds(min_x - margin_x, min_y - margin_y, max_x + margin_x, max_y + margin_y, src.transform)
    data = src.read(window=window)  # Leer todas las bandas de la región de interés
    transform = src.window_transform(window)  # Ajustar el transform para la ventana

    # Mostrar el mapa base recortado con colores originales
    fig, ax = plt.subplots(figsize=(10, 10))
    if data.shape[0] >= 3:  # Si el raster tiene al menos 3 bandas (RGB)
        show(data[:3], transform=transform, ax=ax)  # Mostrar las 3 primeras bandas como RGB
    else:
        show(data, transform=transform, ax=ax, cmap="gray")  # Fallback a escala de grises

    # Dibujar líneas y puntos sobre el mapa base
    lineas.plot(ax=ax, color=lineas["color"], linewidth=3)

    # Agregar marcadores al final de las líneas
    for idx, estacion in enumerate(estaciones):
        coord_x, coord_y = lineas.geometry[idx].coords[1]  # Coordenadas finales
        # print(f'Coordenada en x: {coord_x} de {estacion} (?)')
        # print(f'Coordenada en y: {coord_y} de {estacion} (?)')
        ax.plot(
            coord_x,
            coord_y,
            marker="*",  # Marcador circular
            color=lineas["color"][idx],
            markersize=15,
            label=estacion["nombre"],
        )

    # Dibujar el punto de medición
    coord_x_1, coord_y_1 = punto.geometry[0].coords[0]
    print(f'Coordenada en x: {coord_x_1}')
    print(f'Coordenada en y: {coord_y_1}')

    # punto.plot(ax=ax, color="black", marker="*", markersize=100, label="Punto de Medición")
    ax.plot(
        coord_x_1,
        coord_y_1,
        marker=".",  # Marcador circular
        color='white',
        markersize=15,
        label='Punto',
    )

    # Ajustar límites dinámicamente
    margin_factor = 0.2  # Factor del 10% para añadir márgenes
    x_min, x_max = min(lons), max(lons)
    y_min, y_max = min(lats), max(lats)
    margin_x = (x_max - x_min) * margin_factor
    margin_y = (y_max - y_min) * margin_factor
    ax.set_xlim(x_min - margin_x, x_max + margin_x)
    ax.set_ylim(y_min - margin_y, y_max + margin_y)
    # ax.set_xlim(x_min, x_max)
    # ax.set_ylim(y_min, y_max)

    # Títulos y etiquetas
    # plt.title("Líneas desde Punto de Medición hacia Estaciones")
    plt.xlabel("Coordenadas Este (m)")
    plt.ylabel("Coordenadas Norte (m)")

    # Agregar la leyenda
    ax.legend(loc="upper right", fontsize=10, title_fontsize=12)

    # Ocultar los ejes y las etiquetas
    ax.axis("off")

    # Guardar el mapa como archivo de imagen
    plt.savefig("mapa_ajustado_con_leyenda.png", dpi=300, bbox_inches="tight", pad_inches=0)
    print("Gráfico guardado como 'mapa_ajustado_con_leyenda.png'")
