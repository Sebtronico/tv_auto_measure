import xlwings as wx
import pandas as pd
from rapidfuzz import process, fuzz
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pyproj import Transformer
import datetime
from src.utils.constants import *
import os
import shutil
from src.utils.utils import rpath
import geopandas as gpd
from shapely.geometry import Point, LineString
import rasterio
from rasterio.plot import show
from rasterio.windows import Window

class ExcelReport:
    def __init__(self):
        self.analog_filename = rpath('./templates/FOR_Registro Monitoreo In Situ TV Analógica_V0.xlsm')
        self.digital_filename = rpath('./templates/FOR_Registro Monitoreo In Situ TDT_V0.xlsm')

        # Se inicia la app
        self.app = wx.App(visible=False)

        # Abrimos workbooks
        self.wb_analog = self.app.books.open(self.analog_filename)
        self.register_sheet = self.wb_analog.sheets["Registro"]
        self.graphical_supports_sheet = self.wb_analog.sheets["Soportes Gráficos"]

        self.wb_digital = self.app.books.open(self.digital_filename)
        self.general_info_sheet = self.wb_digital.sheets["Información Gral"]
        self.channel_template_sheet = self.wb_digital.sheets["Template"]

        # Cargue de archivo de referencias
        self.filename_references = rpath('./src/utils/Referencias.xlsx')
        self.stations = pd.read_excel(self.filename_references, sheet_name=2)

    def _get_closest_station_name(self, station: str):
        stations = self.stations['TX_TDT'].tolist()
        if station == "SUBA - RTVC":
            found_station = "CERRO SUBA - CITYTV"
        elif station == "MUX OSAL - RTVC":
            found_station = "MUX OSAL - CALI"
        elif station == "TELEPASTO - RTVC":
            found_station = "TELEPASTO - TELEPASTO"
        elif station == "TELESANTIAGO - RTVC":
            found_station = "TELESANTIAGO - TELESANTIAGO"
        elif station == "TV IPIALES - RTVC":
            found_station = "TV IPIALES - TV IPIALES"
        elif station == "U.DEL PACÍFICO - RTVC":
            found_station = "U.DEL PACÍFICO - U.DEL PACÍFICO"
        elif station == "U. DEL VALLE - RTVC":
            found_station = "U. DEL VALLE - U. DEL VALLE"
        elif station == "TELEPETROLEO - RTVC":
            found_station = "TELEPETROLEO - ENLACE TV"
        else:
            found_station, score, index = process.extractOne(
                station, stations, scorer=fuzz.ratio)

        return found_station

    def get_station_list(self, digital_measurement_dictionary: dict):
        station_list = []
        for channel, results in digital_measurement_dictionary.items():
            if results['service_name'] in ['Caracol', 'RCN']:
                station = f"{results['station']} - CCNP".upper()
            else:
                station = f"{results['station']} - RTVC".upper()
            found_station = self._get_closest_station_name(station)
            station_list.append(found_station)

        return list(set(station_list))

    def plot_elevation_profile(self, lat_point: float, lon_point: float, station_list: list, path: str):
        # Coordenadas del punto de medición
        point_coord = (lat_point, lon_point)

        # Cargue del archivo DEM
        # dem_file = rpath("./resources/SRTM_30_Col1.tif")
        dem_file = rpath("./resources/Colombia_DEM.tif")
        with rasterio.open(dem_file) as dem:
            dem_crs = dem.crs  # Obtener CRS del DEM

        # Transformador de coordenadas
        transformer = Transformer.from_crs("EPSG:4326", dem_crs.to_string(), always_xy=True)

        # Convertir coordenadas a sistema proyectado del DEM
        point_coord_projected = transformer.transform(*point_coord[::-1])

        # Número de puntos a lo largo del perfil
        # number_of_points = 500

        for station in station_list:
            # Se obtienen las coordenadas de la estación desde el documento de referencias
            index = self.stations.index[self.stations['TX_TDT'] == station].tolist()[0]
            lat_station = self.stations.at[index, 'LAT_D']
            lon_station = self.stations.at[index, 'LONG_D']
            station_coord = (lat_station, lon_station)

            # Convertir coordenadas a sistema proyectado DEM
            station_coord_projected = transformer.transform(*station_coord[::-1])

            # Se crea una línea entre la estación y el punto
            line = LineString([station_coord_projected, point_coord_projected])
            
            # Get the length of the line in meters
            line_length_meters = line.length 

            # Determine an appropriate number of points based on DEM resolution
            # Aim for roughly 1 point per pixel, or maybe 2 points per pixel for a slightly smoother look
            dem_resolution = 100 # Your DEM resolution is 100m
            desired_points_per_pixel = 1.5 # Example: 1.5 points for every DEM pixel length

            number_of_points = max(50, int(line_length_meters / dem_resolution * desired_points_per_pixel)) 
            # Use max(50, ...) to ensure a minimum number of points for very short lines

            interpolated_points = [line.interpolate(dist, normalized=True).coords[0] for dist in np.linspace(0, 1, number_of_points)]

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
                                dist = np.sqrt((x - prev_x) ** 2 + (y - prev_y) ** 2) / 1000  # Convertir a km
                                distances.append(distances[-1] + dist if distances else dist)
                            else:
                                distances.append(0)
                        else:
                            elevations.append(None)
                    except Exception as e:
                        elevations.append(None)

            # Graficar el perfil de elevación
            elevations = np.array(elevations, dtype=float)
            distances = np.array(distances, dtype=float)

            plt.figure(figsize=(13.19, 3.35))
            plt.plot(distances, np.nan_to_num(elevations), color='brown')
            plt.fill_between(distances, np.nan_to_num(elevations), color='lightcoral', alpha=0.5)
            plt.xlabel("Distancia [km]")
            plt.ylabel("Elevación [m]")
            plt.grid()
            plt.xlim(min(distances), max(distances))
            plt.ylim(min(elevations) * 0.9, max(elevations) * 1.05)
            plt.savefig(rpath(f"{path}/profile_images/elevation_profile-{station}.png"), dpi=300, bbox_inches="tight", pad_inches=0.25)
            plt.close()

    def plot_distances_image(self, lat_point: float, lon_point: float, station_list: list, path: str):
        # Coordenadas del punto de medición (asegurándose que sea lon, lat para shapely)
        point_coord = (lon_point, lat_point)

        stations_data = []
        for station_name in station_list:
            # Se obtienen las coordenadas de la estación desde el documento de referencias
            index = self.stations.index[self.stations['TX_TDT'] == station_name].tolist()[0]
            lat_station = self.stations.at[index, 'LAT_D']
            lon_station = self.stations.at[index, 'LONG_D']
            station_coord = (lon_station, lat_station)

            stations_data.append({'station_name': station_name, 'coordinates': station_coord})

        # Crear un GeoDataFrame para las líneas
        geometries = [LineString([point_coord, station['coordinates']]) for station in stations_data]

        # Creación de líneas y punto
        colors = ['blue', 'green', 'red', 'cyan', 'magenta', 'yellow', 'black', 'purple', 'orange', 'brown'] 

        lines = gpd.GeoDataFrame({'color': [colors[i % len(colors)] for i in range(len(station_list))], 'name': station_list},
                                 geometry=geometries, crs='EPSG:4326')
        point_gdf = gpd.GeoDataFrame({'name': ['Punto de medición']}, geometry=[Point(point_coord)], crs='EPSG:4326')

        # Cargar el archivo raster del mapa base
        raster_path = rpath('./resources/Colombia_Satelital.tif')
        
        with rasterio.open(raster_path) as src:
            # Reproyectar las capas al CRS del raster
            lines_proj = lines.to_crs(src.crs)
            point_proj = point_gdf.to_crs(src.crs)

            # Obtener todos los puntos de interés en el CRS proyectado
            all_x_coords = [p.x for p in point_proj.geometry]
            all_y_coords = [p.y for p in point_proj.geometry]
            for line_geom in lines_proj.geometry:
                all_x_coords.extend([c[0] for c in line_geom.coords])
                all_y_coords.extend([c[1] for c in line_geom.coords])

            min_x_data, max_x_data = min(all_x_coords), max(all_x_coords)
            min_y_data, max_y_data = min(all_y_coords), max(all_y_coords)

            # Calcular los límites mínimos que contienen todos los datos
            data_width = max_x_data - min_x_data
            data_height = max_y_data - min_y_data

            # Definir la relación de aspecto deseada
            desired_aspect_ratio = 1.77 # 16:9

            # Ajustar los límites para que la región sea cuadrada (o la relación de aspecto deseada)
            if data_width / data_height > desired_aspect_ratio:
                # Ajustar el alto
                target_height = data_width / desired_aspect_ratio
                y_center = (min_y_data + max_y_data) / 2
                min_y_final = y_center - target_height / 2
                max_y_final = y_center + target_height / 2
                min_x_final, max_x_final = min_x_data, max_x_data
            else:
                # Ajustar el ancho
                target_width = data_height * desired_aspect_ratio
                x_center = (min_x_data + max_x_data) / 2
                min_x_final = x_center - target_width / 2
                max_x_final = x_center + target_width / 2
                min_y_final, max_y_final = min_y_data, max_y_data
            
            # Añadir un margen (ej. 20% extra) a los límites finales
            margin_factor = 0.2 
            total_width = max_x_final - min_x_final
            total_height = max_y_final - min_y_final

            margin_x = total_width * margin_factor
            margin_y = total_height * margin_factor
            
            final_bounds = (min_x_final - margin_x, min_y_final - margin_y, 
                            max_x_final + margin_x, max_y_final + margin_y)

            # Recortar el raster usando los límites finales y la función window.from_bounds
            window = src.window(*final_bounds)
            window = window.intersection(Window(col_off=0, row_off=0, width=src.width, height=src.height))
            
            # Si la ventana resulta ser inválida (ej. no hay intersección)
            if window.width <= 0 or window.height <= 0:
                print("La región de interés está fuera de los límites del raster.")
                return # Salir de la función o manejar el error

            # Leer solo la parte del raster dentro de la ventana
            data = src.read(window=window)
            transform = src.window_transform(window) # Obtener la nueva transformación para la ventana
                        
            # Si src.nodata existe y se desea usarlo para rellenar fuera del área del raster:
            if src.nodata is not None:
                # En algunas visualizaciones, la lectura de una ventana puede producir píxeles
                # "transparentes" o "NoData" fuera de los límites reales del raster.
                # Aquí, nos aseguramos de que el `show` de rasterio los maneje correctamente.
                pass # rasterio.plot.show generalmente hace esto bien con nodata
            
            # Mostrar el mapa base recortado
            fig, ax = plt.subplots(figsize=(10, 10)) # Aseguramos que la figura sea cuadrada
            
            # Manejar el número de bandas para la visualización
            if data.shape[0] >= 3:
                # Transponer para obtener (alto, ancho, bandas) para matplotlib si no es así
                if data.shape[0] > 3: # Si tiene más de 3 bandas, toma las 3 primeras
                    display_data = data[:3]
                else:
                    display_data = data
                
                show(display_data, transform=transform, ax=ax)
            else:
                # Si es una imagen en escala de grises (1 banda) o menos de 3 bandas
                show(data, transform=transform, ax=ax, cmap='gray')

            # Dibujar líneas y puntos sobre el mapa base
            lines_proj.plot(ax=ax, color=lines_proj['color'], linewidth=3, zorder=2) # zorder para asegurar que estén encima

            # Agregar marcadores al final de las líneas (estaciones)
            for idx, station_data_item in enumerate(stations_data):
                coord_x, coord_y = lines_proj.geometry[idx].coords[1] # Coordenadas finales de la línea
                ax.plot(
                    coord_x,
                    coord_y,
                    marker='*',
                    color=lines_proj['color'][idx],
                    markersize=15,
                    label=station_data_item['station_name'], # Usar station_name del diccionario
                    zorder=3 # Asegurar que los marcadores estén encima
                )

            # Dibujar el punto de medición
            coord_x_1, coord_y_1 = point_proj.geometry[0].coords[0]

            ax.plot(
                coord_x_1,
                coord_y_1,
                marker='o', # Usar 'o' para punto, ' . ' es muy pequeño
                color='white',
                markeredgecolor='black', # Borde negro para visibilidad en fondo blanco
                markersize=10,
                label='Punto de medición',
                zorder=3
            )

            ax.legend(loc='upper right', fontsize=10, title_fontsize=12)
            ax.axis('off') # Quitar ejes para una imagen limpia

            # Guardar el mapa como archivo de imagen
            plt.savefig(rpath(f'{path}/profile_images/distances.png'), dpi=300, bbox_inches='tight', pad_inches=0)
            plt.close()

    def fill_register_sheet(self, site_dictionary: dict, analog_measurement_dictionary: dict):
        # Información general
        self.register_sheet.range("E3").value = site_dictionary['municipality']
        self.register_sheet.range("E4").value = site_dictionary['department']
        self.register_sheet.range("E5").value = site_dictionary['address']
        self.register_sheet.range("E6").value = site_dictionary['latitude_dms']
        self.register_sheet.range("E7").value = site_dictionary['longitude_dms']
        self.register_sheet.range("E8").value = site_dictionary['altitude']
        self.register_sheet.range("E9").value = site_dictionary['point']
        self.register_sheet.range("E10").value = site_dictionary['around']
        self.register_sheet.range("E11").value = site_dictionary['terrain']
        self.register_sheet.range("E12").value = site_dictionary['signal_path']
        self.register_sheet.range("E13").value = site_dictionary['signal_obstruction']

        # Servidores públicos responsables de las mediciones
        now = datetime.datetime.now()  # Obetención de fecha y hora de la medida
        date = f'{str(now.day).zfill(2)}/{str(now.month).zfill(2)}/{str(now.year)}'
        self.register_sheet.range("L3").value = date
        self.register_sheet.range("H6").value = site_dictionary['engineer_1']
        self.register_sheet.range("H10").value = site_dictionary['engineer_2']

        # Equipo utilizado para la realización del monitoreo in situ
        self.register_sheet.range("R4").value = site_dictionary['instrument_type']
        self.register_sheet.range("T4").value = site_dictionary['instrument_brand']
        self.register_sheet.range("U4").value = site_dictionary['instrument_model']
        self.register_sheet.range("U4").value = site_dictionary['instrument_serial']

        self.register_sheet.range("R5").value = 'Antena'
        self.register_sheet.range("T5").value = site_dictionary['a_antenna_brand']
        self.register_sheet.range("U5").value = site_dictionary['a_antenna_model']

        # Mediciones de niveles de servicio e interferencias para televisión analógica
        for row, (channel, dic) in enumerate(analog_measurement_dictionary.items(), start=21):
            self.register_sheet.range(f"C{row}").value = dic['hour']
            self.register_sheet.range(f"D{row}").value = channel

            service_name = dic['service_name']
            if service_name == 'Caracol':
                fix_service_name = 'Caracol TV S.A.'
            elif service_name == 'RCN':
                fix_service_name = 'RCN TV S.A.'
            elif service_name == 'Canal 1':
                fix_service_name = 'Canal Uno'
            elif service_name == 'Canal Institucional':
                fix_service_name = 'C.Institucional'
            else:
                fix_service_name = service_name

            self.register_sheet.range(f"E{row}").value = fix_service_name
            self.register_sheet.range(f"F{row}").value = dic['station']
            self.register_sheet.range(f"J{row}").value = dic['power_video']
            self.register_sheet.range(f"K{row}").value = dic['power_audio']
            self.register_sheet.range(f"L{row}").value = dic['frequency_video']
            self.register_sheet.range(f"M{row}").value = dic['frequency_audio']

    def fill_graphical_support_sheet(self, site_dictionary: dict, analog_measurement_dictionary: dict, path: str):
        rows_for_images = [3, 22, 41, 63, 82, 101, 123, 142, 161, 183, 202, 221, 243, 262, 281, 303, 322, 341, 363, 382, 401]

        for index, (channel, dic) in enumerate(analog_measurement_dictionary.items()):
            if index < len(rows_for_images):  # Asegurarse de tener filas disponibles para las imágenes
                municipality = site_dictionary['municipality']
                point = str(site_dictionary['point']).zfill(2)
                station = dic['station']
                service_name = dic['service_name']
                img_path = rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_A_{service_name}/CH_{channel}.png')

                # Verificar que la imagen existe
                if os.path.exists(img_path):
                    # Calcular posición (celda A + fila)
                    left = self.graphical_supports_sheet.range(f"A{rows_for_images[index]}").left
                    top = self.graphical_supports_sheet.range(f"A{rows_for_images[index]}").top

                    # Insertar imagen usando win32com
                    img = self.graphical_supports_sheet.pictures.add(
                        image=img_path,             # Ruta absoluta a la imagen
                        left=left,                  # Posición izquierda
                        top=top,                    # Posición superior
                        scale=0.255
                    )

    def fill_general_info_sheet(self, site_dictionary: dict, digital_measurement_dictionary: dict, path: str):
        # Características del punto de medición
        self.general_info_sheet.range("A8").value = site_dictionary['point']
        self.general_info_sheet.range("C8").value = site_dictionary['municipality']
        self.general_info_sheet.range("M8").value = site_dictionary['department']
        self.general_info_sheet.range("V8").value = site_dictionary['latitude_dms']
        self.general_info_sheet.range("AB8").value = site_dictionary['longitude_dms']
        self.general_info_sheet.range("AH8").value = site_dictionary['altitude']
        self.general_info_sheet.range("A11").value = site_dictionary['around']
        self.general_info_sheet.range("H11").value = site_dictionary['terrain']
        self.general_info_sheet.range("M11").value = site_dictionary['signal_path']
        self.general_info_sheet.range("S11").value = site_dictionary['signal_obstruction']
        self.general_info_sheet.range("Z11").value = site_dictionary['address']

        # Equipo utilizado para la medición
        self.general_info_sheet.range("A15").value = site_dictionary['instrument_type']
        self.general_info_sheet.range("I15").value = site_dictionary['instrument_brand']
        self.general_info_sheet.range("Q15").value = site_dictionary['instrument_model']
        self.general_info_sheet.range("W15").value = site_dictionary['instrument_serial']

        self.general_info_sheet.range("A16").value = 'Antena'
        self.general_info_sheet.range("I16").value = site_dictionary['d_antenna_brand']
        self.general_info_sheet.range("Q16").value = site_dictionary['d_antenna_model']

        # Características de las estaciones que se reciben en el punto donde se realiza la medición
        station_list = self.get_station_list(digital_measurement_dictionary)
        for row, station in enumerate(station_list, start=20):
            self.general_info_sheet.range(f"A{row}").value = station

        # Características de la medición
        now = datetime.datetime.now()  # Obetención de fecha y hora de la medida
        date = f'{str(now.day).zfill(2)}/{str(now.month).zfill(2)}/{str(now.year)}'
        self.general_info_sheet.range("AE31").value = date

        # Perfiles de terreno desde el punto donde se realiza la medición a las estaciones monitoreadas
        rows_for_images = [73, 77, 81, 85, 89, 93, 98, 102]
        for index, station in enumerate(station_list):
            if index < len(rows_for_images):
                img_path = rpath(f'{path}/profile_images/elevation_profile-{station}.png')
                
                if os.path.exists(img_path):
                    # Calcular posición
                    left = self.general_info_sheet.range(f"A{rows_for_images[index]}").left
                    top = self.general_info_sheet.range(f"A{rows_for_images[index]}").top
                    
                    # Insertar imagen
                    img = self.general_info_sheet.pictures.add(
                        image=img_path,
                        left=left,
                        top=top,
                        scale=0.96
                    )

        # Distancias desde el punto de medición a las estaciones monitoreadas
        distance_img_path = rpath(f'{path}/profile_images/distances.png')
        if os.path.exists(distance_img_path):
            # Calcular posición para la imagen de distancias
            left = self.general_info_sheet.range("A106").left
            top = self.general_info_sheet.range("A106").top
            
            # Insertar imagen
            distance_img = self.general_info_sheet.pictures.add(
                image=distance_img_path,
                left=left,
                top=top,
                scale=1.39
            )

    def copy_template_sheet(self, new_sheet_name):
        copied_sheet = self.channel_template_sheet.copy(before=self.wb_digital.sheets['SEGUIMIENTO_ANE'])
        copied_sheet.name = new_sheet_name
        
        return copied_sheet
    
    def fill_channel_sheet(self, site_dictionary: dict, digital_measurement_dictionary: dict, sfn_dictionary: dict, path: str):
        # Se crea una hoja y se llena para cada canal digital medido
        for channel, dic in digital_measurement_dictionary.items():
            channel_name = f'CH{channel}'
            channel_sheet = self.copy_template_sheet(channel_name)

            # Características de la medición
            station_name = f"{dic['station'].upper()} - CCNP" if dic['service_name'] in ['Caracol', 'RCN'] else f"{dic['station'].upper()} - RTVC"
            station = self._get_closest_station_name(station_name)
            channel_sheet.range("A8").value = station
            channel_sheet.range("I8").value = dic['service_name'].upper()
            channel_sheet.range("Q8").value = TV_TABLE[channel]
            channel_sheet.range("W8").value = dic['channel_type']
            channel_sheet.range("AF8").value = dic['hour']

            # Características de las estaciones en red SFN que se reciben en el punto donde se realiza la medición
            try:
                sfn_stations_list = list(sfn_dictionary[channel].keys())
                sfn_stations_list.remove(dic['station'])

                for row, sfn_station in enumerate(sfn_stations_list, start=12):
                    renamed_sfn_station = f'{sfn_station.upper()} - CCNP' if dic['service_name'] in ['Caracol', 'RCN'] else f'{sfn_station.upper()} - RTVC'
                    renamed_sfn_station = self._get_closest_station_name(renamed_sfn_station)
                    channel_sheet.range(f"A{row}").value = renamed_sfn_station
            except KeyError:
                pass

            # Registro de verificación de cobertura y parámetros de calidad del servicio TDT. Resultados de medición por PLP.
            plps = PLP_SERVICES[dic['service_name']]
            ts_array = []
            for row, plp in enumerate(plps, start=19):
                key_plp = f'PLP_{plp}'
                if row == 19:
                    channel_sheet.range(f"L{row}").value = dic['channel_power']
                channel_sheet.range(f"P{row}").value = dic[key_plp]['MRPLp']
                channel_sheet.range(f"R{row}").value = dic[key_plp]['BERLdpc']
                channel_sheet.range(f"V{row}").value = dic[key_plp]['cons']
                channel_sheet.range(f"Y{row}").value = dic[key_plp]['PLPCodeRate']
                channel_sheet.range(f"AA{row}").value = dic[key_plp]['FFTMode']
                channel_sheet.range(f"AD{row}").value = dic[key_plp]['GINTerval']
                channel_sheet.range(f"AG{row}").value = dic[key_plp]['PPATtern']
                
                try:
                    # Arreglo completo de Transport Stream
                    for i in range(len(dic[key_plp]['TS'])):
                        ts_array.append(dic[key_plp]['TS'][i])
                except KeyError:
                    pass
                
            # Análisis de Transport Stream
            if ts_array:
                for row, ts_service_result in enumerate(ts_array, start=24):
                    channel_sheet.range(f"M{row}").value = ts_service_result[1]
                    channel_sheet.range(f"O{row}").value = ts_service_result[2]
                    channel_sheet.range(f"Q{row}").value = ts_service_result[3]
                    channel_sheet.range(f"T{row}").value = ts_service_result[4]
                    channel_sheet.range(f"X{row}").value = ts_service_result[5]
                    channel_sheet.range(f"AA{row}").value = ts_service_result[6]
                    channel_sheet.range(f"AD{row}").value = 'No Falla'

            # Soportes de medición
            municipality = site_dictionary['municipality']
            point = str(site_dictionary['point']).zfill(2)
            station = dic['station']
            service_name = dic['service_name']
            plp = PLP_SERVICES[service_name][0]
            
            # Rutas de las imágenes
            img_paths = {
                'Channel Power': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/{TV_TABLE[channel]}.png'),
                'MER': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/PLP_{plp}/{TV_TABLE[channel]}_002.png'),
                'BER': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/PLP_{plp}/{TV_TABLE[channel]}_003.png'),
                'Constelation': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/PLP_{plp}/{TV_TABLE[channel]}_001.png'),
                'Echo Pattern': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/PLP_{plp}/{TV_TABLE[channel]}_008.png'),
                'Shoulders': rpath(f'{path}//Soportes punto de medición/{station}/CH_{channel}_D_{service_name}/PLP_{plp}/{TV_TABLE[channel]}_010.png')
            }
            
            # Celdas donde se insertarán las imágenes
            img_cells = {
                'Channel Power': 'A44',
                'MER': 'M44',
                'BER': 'Z44',
                'Constelation': 'A63',
                'Echo Pattern': 'M63',
                'Shoulders': 'Z63'
            }
            
            # Insertar imágenes
            for img_name, img_path in img_paths.items():
                # Verificar que la imagen existe
                if os.path.exists(os.path.abspath(img_path)):
                    left = channel_sheet.range(img_cells[img_name]).left
                    top = channel_sheet.range(img_cells[img_name]).top
                    
                    # Insertar imagen y ajustar tamaño
                    img = channel_sheet.pictures.add(
                        image=img_path,
                        left=left,
                        top=top,
                        scale=0.32
                    )

    @staticmethod
    def sort_dictionary(dictionary: dict):
        # Se obtiene la lista ordenada de las claves (canales) en el diccionario
        list_of_channels = list(dictionary.keys())
        list_of_channels.sort()
        sorted_dictionary = {}
        
        for channel in list_of_channels:
            sorted_dictionary[channel] = dictionary[channel]
        
        return sorted_dictionary

    def fill_reports(self, site_dictionary: dict, analog_measurement_dictionary: dict, digital_measurement_dictionary: dict, sfn_dictionary, path: str):
        # Se ordenan los diccionarios de entrada por canal
        analog_measurement_dictionary = self.sort_dictionary(analog_measurement_dictionary)
        digital_measurement_dictionary = self.sort_dictionary(digital_measurement_dictionary)

        # Coordenadas del punto, para las funciones de graficado
        lat_point = site_dictionary['latitude_dec']
        lon_point = site_dictionary['longitude_dec']

        # Lista de estaciones de digital
        digital_station_list = self.get_station_list(digital_measurement_dictionary)

        # Generación de gráficas de perfiles y distancias
        os.makedirs(rpath(f'{path}/profile_images'), exist_ok=True)
        self.plot_elevation_profile(lat_point, lon_point, digital_station_list, path)
        self.plot_distances_image(lat_point, lon_point, digital_station_list, path)

        # Llenado de postprocesamiento anaógico
        self.fill_register_sheet(site_dictionary, analog_measurement_dictionary)
        self.fill_graphical_support_sheet(site_dictionary, analog_measurement_dictionary, path)

        # Llenado del postprocesamiento tdt.
        self.fill_general_info_sheet(site_dictionary, digital_measurement_dictionary, path)
        self.fill_channel_sheet(site_dictionary, digital_measurement_dictionary, sfn_dictionary, path)

        # Eliminar hoja de plantilla del formato TDT
        self.wb_digital.sheets["Template"].delete()

        # Guardado de archivos.
        municipality = site_dictionary['municipality']
        point = str(site_dictionary['point']).zfill(2)
        
        # Obtener la ruta de la carpeta 'path'
        path_completo = rpath(path)

        # Obtener la ruta de la carpeta que contiene a 'path'
        base_save_path = os.path.dirname(path_completo)

        # Construir la ruta final usando os.path.join()
        analog_save_path = os.path.join(base_save_path, f'FOR_Registro Monitoreo In Situ TV Analógica_V0_P{point}.xlsm')
        digital_save_path = os.path.join(base_save_path, f'FOR_Registro Monitoreo In Situ TDT_V0_P{point}.xlsm')
        
        if os.path.exists(os.path.abspath(analog_save_path)):
            os.remove(analog_save_path)

        if os.path.exists(os.path.abspath(analog_save_path)):
            os.remove(digital_save_path)

        # Guardar los workbooks
        self.wb_analog.save(analog_save_path)
        self.wb_digital.save(digital_save_path)
        
        # Opcional: cerrar los archivos si ya no se van a utilizar
        self.wb_analog.close()
        self.wb_digital.close()
        self.app.quit()

if __name__ == "__main__":
    # Ejemplo de uso
    report = ExcelReport()
    
    site_dictionary = {'municipality': 'Tenjo', 'department': 'Cundinamarca', 'point': 4, 'a_antenna_brand': 'Aaronia', 'a_antenna_model': 'Bicolog', 'instrument_type': 'Analizador de Televisión', 'instrument_brand': 'Rohde&Schwarz', 'instrument_model': 'ETL', 'instrument_serial': '103982', 'd_antenna_brand': 'Televes', 'd_antenna_model': 'DAT BOSS', 'latitude_dec': 4.67759367, 'longitude_dec': -74.054271, 'latitude_dms': '4° 40\' 39,34" N', 'longitude_dms': '74° 3\' 15,38" W', 'altitude': 2604, 'around': 'Urbano', 'terrain': 'Plano', 'signal_path': 'LOS', 'signal_obstruction': 'Ninguna', 'engineer_1': 'Sebastian Chavez Martinez', 'engineer_2': 'German Leonardo Vargas Gutierrrez', 'address': 'asdfasdf'}
    analog_measurement_dictionary = {3: {'frequency_video': 61.25, 'frequency_audio': 65.75, 'power_video': -10.49, 'power_audio': -10.86, 'hour': '14:51', 'station': 'Tibitóc', 'service_name': 'Canal 1'}, 6: {'frequency_video': 83.25, 'frequency_audio': 87.75, 'power_video': -12.43, 'power_audio': -11.86, 'hour': '14:51', 'station': 'Tibitóc', 'service_name': 'Canal Institucional'}, 12: {'frequency_video': 205.25, 'frequency_audio': 209.75, 'power_video': -11.38, 'power_audio': -11.65, 'hour': '14:52', 'station': 'Tibitóc', 'service_name': 'Señal Colombia'}, 8: {'frequency_video': 181.25, 'frequency_audio': 185.75, 'power_video': -11.74, 'power_audio': -11.77, 'hour': '14:52', 'station': 'Suba', 'service_name': 'RCN'}, 10: {'frequency_video': 193.25, 'frequency_audio': 197.75, 'power_video': -11.17, 'power_audio': -11.47, 'hour': '14:52', 'station': 'Suba', 'service_name': 'Caracol'}, 21: {'frequency_video': 513.25, 'frequency_audio': 517.75, 'power_video': -10.6, 'power_audio': -10.02, 'hour': '14:52', 'station': 'Suba', 'service_name': 'CityTV'}, 23: {'frequency_video': 525.25, 'frequency_audio': 529.75, 'power_video': -11.1, 'power_audio': -10.92, 'hour': '15:08', 'station': 'Calatrava', 'service_name': 'Teveandina'}, 25: {'frequency_video': 537.25, 'frequency_audio': 541.75, 'power_video': -10.75, 'power_audio': -10.41, 'hour': '15:08', 'station': 'Calatrava', 'service_name': 'Señal Colombia'}, 32: {'frequency_video': 579.25, 'frequency_audio': 583.75, 'power_video': -10.46, 'power_audio': -10.51, 'hour': '15:08', 'station': 'Calatrava', 'service_name': 'Canal Capital'}, 36: {'frequency_video': 603.25, 'frequency_audio': 607.75, 'power_video': -10.58, 'power_audio': -9.63, 'hour': '15:09', 'station': 'Calatrava', 'service_name': 'Canal 1'}, 38: {'frequency_video': 615.25, 'frequency_audio': 619.75, 'power_video': -11.06, 'power_audio': -10.52, 'hour': '15:09', 'station': 'Calatrava', 'service_name': 'Canal Institucional'}, 2: {'frequency_video': 55.25, 'frequency_audio': 59.75, 'power_video': -11.42, 'power_audio': -11.29, 'hour': '15:25', 'station': 'Manjui', 'service_name': 'Canal Capital'}, 4: {'frequency_video': 67.25, 'frequency_audio': 71.75, 'power_video': -11.5, 'power_audio': -10.71, 'hour': '15:26', 'station': 'Manjui', 'service_name': 'RCN'}, 5: {'frequency_video': 77.25, 'frequency_audio': 81.75, 'power_video': -11.97, 'power_audio': -12.39, 'hour': '15:26', 'station': 'Manjui', 'service_name': 'Caracol'}, 7: {'frequency_video': 175.25, 'frequency_audio': 179.75, 'power_video': -11.33, 'power_audio': -11.28, 'hour': '15:26', 'station': 'Manjui', 'service_name': 'Canal 1'}, 9: {'frequency_video': 187.25, 'frequency_audio': 191.75, 'power_video': -11.58, 'power_audio': -11.09, 'hour': '15:26', 'station': 'Manjui', 'service_name': 'Canal Institucional'}, 11: {'frequency_video': 199.25, 'frequency_audio': 203.75, 'power_video': -12.1, 'power_audio': -10.63, 'hour': '15:26', 'station': 'Manjui', 'service_name': 'Señal Colombia'}}
    digital_measurement_dictionary = {14: {'date': '05/05/2025', 'hour': '14:53:17', 'channel_power': 66.88, 'channel_type': 'Gauss (σ = 1.0 dB)', 'PLP_0': {'SALower': 32.2113342285, 'SAUPper': 3.22442626953, 'LEVel': 56.0706126339, 'CFOFfset': -47.2, 'BROFfset': -0.1, 'PERatio': 0.0, 'BERLdpc': 0.00064, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '3/4', 'IMBalance': -6.74, 'QERRor': -1.43, 'CSUPpression': 3.2, 'MRLO': 28.665, 'MPLO': -1.416, 'MRPLp': 30.1, 'MPPLp': 3.68921748752, 'ERPLp': 2.04746810336, 'EPPLp': 42.8105424169, 'cons': '64QAM', 'FFTMode': '16k ext', 'AMPLitude': 47.624994278, 'PHASe': 2696.04211426, 'GDELay': 4.310742e-05, 'PPATtern': 'PP3', 'TS': [['CARACOL TV HD', '21', '12290', '20', 'PID', 'PCR Jitter, PCR Repetition', 'EITActual'], ['CARACOL HD 2', '22', '12290', '20', 'PID', 'PCR Jitter, PCR Repetition', 'EITActual'], ['LA KALLE', '25', '12290', '20', 'PID', 'PCR Jitter, PCR Repetition', 'EITActual']]}, 'PLP_1': {'SALower': 35.3370437622, 'SAUPper': 1.49951171875, 'LEVel': 56.1906126339, 'CFOFfset': -46.8, 'BROFfset': -0.1, 'PERatio': 0.0, 'BERLdpc': 4.1e-05, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '3/5', 'IMBalance': -14.42, 'QERRor': -2.67, 'CSUPpression': 7.3, 'MRLO': 28.397, 'MPLO': -3.201, 'MRPLp': 29.0, 'MPPLp': -2.34954610693, 'ERPLp': 3.72146490389, 'EPPLp': 131.06214613, 'cons': 'QPSK', 'FFTMode': '16k ext', 'AMPLitude': 37.6625132561, 'PHASe': 2271.78778076, 'GDELay': 6.924414e-05, 'PPATtern': 'PP3', 'TS': [['Caracol Movil', '30', 'ND', '21', 'NA', 'PCR Jitter, PCR Repetition', 'NA']]}, 'station': 'Suba', 'service_name': 'Caracol'}, 15: {'date': '05/05/2025', 'hour': '14:59:20', 'channel_power': 68.46, 'channel_type': 'Rice (σ = 1.01 dB)', 'PLP_0': {'SALower': 3.35957336426, 'SAUPper': -8.83447265625, 'LEVel': 57.5206126339, 'CFOFfset': -38.2, 'BROFfset': -0.08, 'PERatio': 0.0, 'BERLdpc': 2.4e-05, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '3/4', 'IMBalance': -1.84, 'QERRor': -0.14, 'CSUPpression': 6.7, 'MRLO': 32.771, 'MPLO': 17.828, 'MRPLp': 33.2, 'MPPLp': 9.31137625018, 'ERPLp': 1.43934437307, 'EPPLp': 22.4100490113, 'cons': '64QAM', 'FFTMode': '16k ext', 'AMPLitude': 12.4321904182, 'PHASe': 153.501724243, 'GDELay': 3.46289e-06, 'PPATtern': 'PP2', 'TS': [['RCN HD', '2', '12289', '10', 'PID', 'PCR Jitter, PCR Repetition', 'NA'], ['RCN HD 2', '3', '12289', '10', 'PID', 'PCR Jitter, PCR Repetition', 'NA']]}, 'PLP_1': {'SALower': 2.26934814453, 'SAUPper': -9.1859703064, 'LEVel': 57.5306126339, 'CFOFfset': -37.8, 'BROFfset': -0.08, 'PERatio': 0.0, 'BERLdpc': 0.0, 'BBCH': 0.0, 'FERatio': '---', 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '1/2', 'IMBalance': -1.77, 'QERRor': -0.04, 'CSUPpression': 5.7, 'MRLO': 32.62, 'MPLO': 18.39, 'MRPLp': 33.0, 'MPPLp': 10.8778140409, 'ERPLp': 2.2085386747, 'EPPLp': 28.5830966719, 'cons': 'QPSK', 'FFTMode': '16k ext', 'AMPLitude': 14.0440087318, 'PHASe': 147.775810242, 'GDELay': 2.45496e-06, 'PPATtern': 'PP2', 'TS': [['RCN MOVIL', '1', '12289', '11', 'PID', 'PCR Jitter, PCR Repetition', 'NA']]}, 'station': 'Suba', 'service_name': 'RCN'}, 27: {'date': '05/05/2025', 'hour': '15:05:10', 'channel_power': 72.06, 'channel_type': 'Rice (σ = 1.28 dB)', 'PLP_1': {'SALower': 44.6417541504, 'SAUPper': -5.85527420044, 'LEVel': 62.8106126339, 'CFOFfset': -60.5, 'BROFfset': -0.13, 'PERatio': 0.0, 'BERLdpc': 0.0, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/16', 'PLPCodeRate': '1/2', 'IMBalance': -0.04, 'QERRor': 0.03, 'CSUPpression': 34.9, 'MRLO': 34.69, 'MPLO': 24.128, 'MRPLp': 32.6, 'MPPLp': 20.0208953643, 'ERPLp': 1.73833322581, 'EPPLp': 7.43565065218, 'cons': '16QAM', 'FFTMode': '8k ext', 'AMPLitude': 6.58491492271, 'PHASe': 31.4276857376, 'GDELay': 1.89581e-06, 'PPATtern': 'PP4', 'TS': [['Citytv', '16001', '12481', '10001', 'PID', 'PCR Jitter, PCR Repetition', 'TDT'], ['El Tiempo Television', '16002', '12481', '10001', 'PID', 'PCR Jitter, PCR Repetition', 'TDT']]}, 'station': 'Suba', 'service_name': 'CityTV'}, 16: {'date': '05/05/2025', 'hour': '15:09:47', 'channel_power': 74.98, 'channel_type': 'Rayleigh (σ = 4.85 dB)', 'PLP_101': {'SALower': 12.1987838745, 'SAUPper': 1.57965087891, 'LEVel': 64.6306126339, 'CFOFfset': -44.4, 'BROFfset': -0.09, 'PERatio': 0.0, 'BERLdpc': 0.001, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '2/3', 'IMBalance': 0.01, 'QERRor': 0.01, 'CSUPpression': 32.2, 'MRLO': 30.025, 'MPLO': 2.923, 'MRPLp': 29.9, 'MPPLp': 4.9905462419, 'ERPLp': 2.10437365804, 'EPPLp': 36.8539714661, 'cons': '64QAM', 'FFTMode': '16k ext', 'AMPLitude': 45.5582485199, 'PHASe': 503.609481812, 'GDELay': 2.246997e-05, 'PPATtern': 'PP3', 'TS': [['SEÑALCOLOMBIA', '41', '12291', '30', 'PID', 'PCR Jitter, PCR Repetition', 'NA'], ['CANAL INSTITUCIONAL', '42', '12291', '30', 'PID', 'PCR Jitter, PCR Repetition', 'NA'], ['CANAL 1', '43', '12291', '30', 'PID', 'PCR Jitter, PCR Repetition', 'NA']]}, 'station': 'Calatrava', 'service_name': 'RTVC'}, 28: {'date': '05/05/2025', 'hour': '15:12:50', 'channel_power': 74.14, 'channel_type': 'Rayleigh (σ = 5.01 dB)', 'PLP_102': {'SALower': 9.34312438965, 'SAUPper': 46.9619979858, 'LEVel': 64.4306126339, 'CFOFfset': -49.5, 'BROFfset': -0.09, 'PERatio': 0.0, 'BERLdpc': 0.011, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '2/3', 'IMBalance': -0.38, 'QERRor': 0.11, 'CSUPpression': 28.3, 'MRLO': 19.212, 'MPLO': -5.263, 'MRPLp': 23.4, 'MPPLp': 2.1124245358, 'ERPLp': 4.39765074592, 'EPPLp': 51.3322602672, 'cons': '64QAM', 'FFTMode': '16k ext', 'AMPLitude': 55.9540462494, 'PHASe': 770.477355957, 'GDELay': 1.967993e-05, 'PPATtern': 'PP3', 'TS': [['CANAL CAPITAL', '6141', '12291', '3071', 'PID', 'PCR Jitter, PCR Repetition', 'NA'], ['EUREKA - CAPITAL', '6142', '12291', '3071', 'PID', 'PCR Jitter, PCR Repetition', 'NA']]}, 'station': 'Calatrava', 'service_name': 'Canal Capital'}, 17: {'date': '05/05/2025', 'hour': '15:16:27', 'channel_power': 75.62, 'channel_type': 'Rayleigh (σ = 5.33 dB)', 'PLP_103': {'SALower': 8.71215820313, 'SAUPper': 48.4190444946, 'LEVel': 66.2106126339, 'CFOFfset': -44.8, 'BROFfset': -0.09, 'PERatio': 0.0, 'BERLdpc': 0.00093, 'BBCH': 0.0, 'FERatio': 0.0, 'ESRatio': 0.0, 'GINTerval': '1/8', 'PLPCodeRate': '2/3', 'IMBalance': -0.05, 'QERRor': -0.07, 'CSUPpression': 30.3, 'MRLO': 31.735, 'MPLO': 3.466, 'MRPLp': 30.3, 'MPPLp': 2.1124245358, 'ERPLp': 1.99251640283, 'EPPLp': 51.3322602672, 'cons': '64QAM', 'FFTMode': '16k ext', 'AMPLitude': 30.5255041122, 'PHASe': 534.927825928, 'GDELay': 1.769252e-05, 'PPATtern': 'PP3', 'TS': []}, 'station': 'Manjuí', 'service_name': 'Teveandina'}}
    sfn_dictionary = {16: {'Tibitóc': 61, 'Calatrava': 162, 'Manjuí': 252}, 17: {'Tibitóc': 61, 'Calatrava': 162, 'Manjuí': 252}, 14: {'Suba': 158, 'Manjuí': 252}, 15: {'Suba': 158, 'Manjuí': 252}}

    report.fill_reports(site_dictionary, analog_measurement_dictionary, digital_measurement_dictionary, sfn_dictionary, './results')

    # excel = win32.Dispatch("Excel.Application")
    # excel.Visible = True
    # excel.DisplayAlerts = False

    # wb = excel.Workbooks.Open(os.path.abspath("./templates/FOR_Registro Monitoreo In Situ TDT_V0.xlsm"))
    # excel.Workbooks(wb.Name).Activate()
    # template_sheet = wb.Sheets("Template")
    # template_sheet.Copy(After=wb.Sheets(wb.Sheets.Count))
    # new_sheet = excel.ActiveSheet
    # new_sheet.Name = "CH_TEST"
