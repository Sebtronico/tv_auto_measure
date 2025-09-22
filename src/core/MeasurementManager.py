from .InstrumentController import EtlManager, FPHManager, MSDManager, ViaviManager
from src.utils.constants import *
from src.utils.utils import rpath
from collections import defaultdict
import json
import os

class MeasurementManager:
    def __init__(self, dtv: EtlManager|None, atv: EtlManager|None, mbk: EtlManager|FPHManager|ViaviManager|None, rtr: MSDManager|None):
        self.dtv = dtv
        self.atv = atv
        self.mbk = mbk
        self.rtr = rtr


    def _rotate(self, park_acimuth: int, station_acimuth: int):
        if self.rtr is not None:
            self.rtr.move_rotor(park_acimuth, station_acimuth)
            return True
        else:
            return False, station_acimuth
            

    @staticmethod
    def _group_sfn_dictionary(dictionary: dict):
        groups = defaultdict(list)

        for key, value in dictionary.items():
            tuple_value = tuple(value.items())
            groups[tuple_value].append(key)

        result = {tuple(keys): dict(value) for value, keys in groups.items()}

        return result
    

    @staticmethod
    def _get_max_trace(dictionary):
        result = {}

        # Obtener todas las claves de los canales
        channels = set()
        for trace in dictionary.values():
            channels.update(trace.keys())

        # Determinar la traza con el mayor nivel para cada canal
        for channel in channels:
            max_trace = max(dictionary, key=lambda trace: dictionary[trace].get(channel, float('-inf')))
            result[channel] = max_trace

        return result


    @staticmethod
    def _get_central_frequency(frequency_list: list):
        n = len(frequency_list)
        if n % 2 == 1:
            return frequency_list[n // 2]
        else:
            return (frequency_list[n // 2 - 1] + frequency_list[n // 2]) / 2
        

    @staticmethod
    def _get_max_power_station(dictionary: dict):
        result = {}
        for channel, stations in dictionary.items():
            max_station = max(stations, key=stations.get)
            result[channel] = max_station

        return result
    

    @staticmethod
    def save_sfn_progress(path: str, sfn_result: dict):
        # Se crea la carpeta de archivos de guardado
        os.makedirs(f"{path}/savefiles", exist_ok=True)

        # Se guarda el diccionario de resultados en un archivo JSON
        filename = f"{path}/savefiles/sfn_progress.json"
        with open(filename, "w") as f:
            json.dump(sfn_result, f, indent=2)


    @staticmethod
    def load_sfn_progress(path: str):
        filename = f"{path}/savefiles/sfn_progress.json"
        with open(filename, "r") as f:
            progress = json.load(f)

        # Convertir las claves del diccionario a enteros
        progress = {int(k): v for k, v in progress.items()}

        return progress
        

    def sfn_measurement(self, dictionary: dict, path: str, park_acimuth: int, callback_rotate = None):
        """
        Realiza la medición de SFN, devuelve el diccionario de resultados de mayores potencias por canal,
        y guarda la imagen de soporte.

        argumentos de entrada:
        dictionary: Diccionario que contiene los canales objeto de medición, sus respectivas estaciones de
            procedencia y acimuth de cada estación

            ejemplo: {16: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254}, 
                      17: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254},
                      14: {'Suba': 153, 'Manjui': 254},
                      15: {'Suba': 153, 'Manjui': 254}}
        """

        # Se carga el progreso, en caso de existir
        if os.path.exists(f"{path}/savefiles/sfn_progress.json"):    
            return self.load_sfn_progress(path)
        

        supports_path = 'Soportes punto de medición' # Nombre de la carpeta de soportes de la medición
        os.makedirs(f'{path}/{supports_path}', exist_ok=True)

        sfn_dictionary = self._group_sfn_dictionary(dictionary)
        sfn_result = {}
        
        for channels_tuple, stations in sfn_dictionary.items():
            # Envío de comando de reset al principio de cada medición.
            # Esto evita que las trazas queden activas.
            self.dtv.reset()

            # Configuración inicial del analizador.
            self.dtv.sfn_setup()

            # Lista ordenada de canales a medir
            channels = sorted(list(channels_tuple))

            # Creación del archivo .txt
            filename = f"{path}/{supports_path}/SFN_CH{' - '.join(map(str, channels))}"
            with open(f'{filename}.txt', 'w') as file:
                file.write(f"Medición SFN de los canales {' - '.join(map(str, channels))}.\n")

            # Configuración del número de canales que se van a medir
            number_of_channels = len(channels)
            self.dtv.write_str(f'POW:ACH:TXCH:COUN {number_of_channels}')

            # Obtención y configuración de la frecuencia central de la medición
            frequencies  = [TV_TABLE[channel] for channel in channels]
            central_frequency = self._get_central_frequency(frequencies)
            self.dtv.write_str(f'FREQ:CENT {central_frequency} MHz')

            # Configuración de los espaciados entre canales
            distances = [(channels[i] - channels[i - 1]) * 6 for i in range(1, len(channels))]
            for i, distance in enumerate(distances, start=1):
                self.dtv.write_str(f'POW:ACH:SPAC:CHAN{i} {distance} MHz')

            # Medición de potencias para cada grupo de canales
            powers = {}
            for trace, (station, acimuth) in enumerate(stations.items(), start=1):
                
                # Se añade la descripción de la estación en el documento .txt
                with open(f'{filename}.txt', 'a') as file:
                    file.write(f'Traza {trace} hacia la estación {station}, en el acimuth {acimuth}°.\n')

                # Intenta rotar la antena
                rotate_result = self._rotate(park_acimuth, acimuth)
                
                # Si se necesita rotación manual
                if isinstance(rotate_result, tuple) and rotate_result[0] is False:
                    target_acimuth = rotate_result[1]
                    
                    # Si se proporcionó una función callback, la llamamos
                    if callback_rotate:
                        callback_rotate(f'{station}(acimuth {target_acimuth}°)')

                # Se activa la traza
                self.dtv.write_str(f'DISP:TRAC{trace}:MODE AVER')
                self.dtv.write_str(f'DET{trace} RMS')
                self.dtv.write_str(f'INIT:CONT OFF')
                self.dtv.write_str(f'SWE:COUN 10')
                self.dtv.write_str(f'INIT;*WAI')
                self.dtv.write_str(f'DISP:TRAC{trace}:MODE VIEW')

                # Se obtienen las potencias de canal
                self.dtv.write_str(f'POW:TRAC {trace}')
                trace_powers = self.dtv.query_bin_or_ascii_float_list_with_opc('CALCulate:MARKer:FUNCtion:POWer:RESult? MCACpower')#[:-1]

                if len(trace_powers) > 1:
                    trace_powers = trace_powers[:-1]
                    
                # Se añade item al diccionario de potencias
                powers.update({trace: {}})
                for j, channel in enumerate(channels):
                    powers[trace].update({channel: trace_powers[j]})
                
                # Se añade item al diccionario de resultado
                for channel, power in zip(channels, trace_powers):
                    sfn_result.setdefault(channel, {})
                    sfn_result[channel][station] = power
                    with open(f'{filename}.txt', 'a') as file:
                        file.write(f'   Potencia en el canal {channel}: {round(power, 2)}.\n')

            # Se ponen los marcadores en las trazas correspondientes
            marker_trace_result = self._get_max_trace(powers)
            for marker, channel in enumerate(channels, start=1):
                self.dtv.write_str(f'CALC:MARK{marker} ON')
                self.dtv.write_str(f'CALC:MARK{marker}:X {TV_TABLE[channel]} MHz')
                self.dtv.write_str(f'CALC:MARK{marker}:TRAC {marker_trace_result[channel]}')

            # Comandos para evitar el asterisco rojo en la imagen
            self.dtv.write_str(f'SWE:COUN 0')
            self.dtv.write_str(f'INIT;*WAI')
            self.dtv.get_screenshot(filename)

            # Para escribir en el archivo .txt los resultados
            partial_result = self._get_max_power_station(sfn_result)
            with open(f'{filename}.txt', 'a') as file:
                for channel in channels:
                    file.write(f'Canal {channel} -> {partial_result[channel]}.\n')

        # Se obtiene el diccionario de filtrado con los resultados.
        final_sfn_result = self._get_max_power_station(sfn_result)

        # Se guarda el progreso
        self.save_sfn_progress(path, final_sfn_result)

        return final_sfn_result
    

    @staticmethod
    def save_tv_progress(path: str, atv_result: dict, dtv_result: dict):
        # Se crea la carpeta de archivos de guardado
        os.makedirs(f"{path}/savefiles", exist_ok=True)

        # Se guardan los diccionarios de resultados en archivos JSON
        tv_filename = f"{path}/savefiles/tv_progress.json"
        results = {
            'atv_result': atv_result,
            'dtv_result': dtv_result
        }

        with open(tv_filename, "w") as f:
            json.dump(results, f, indent=2)


    @staticmethod
    def load_tv_progress(path: str, dictionary: dict):
        # Se carga el archivo de progreso
        filename = rpath(f"{path}/savefiles/tv_progress.json")
        with open(filename, "r") as f:
            progress = json.load(f)
            # progress = {int(k): v for k, v in progress.items()}

        # Se extraen los resultados de atv y dtv
        atv_result = {int(k): v for k, v in progress['atv_result'].items()}
        dtv_result = {int(k): v for k, v in progress['dtv_result'].items()}

        alrd_ang = {int(k) for k in atv_result.keys()}
        alrd_dig = {int(k) for k in dtv_result.keys()}

        station_ang = {int(ch): dat['station'] for ch, dat in atv_result.items()}
        station_dig = {int(ch): dat['station'] for ch, dat in dtv_result.items()}

        updated_dictionary = {}

        for station, info in dictionary.items():
            # Copias de los subdiccionarios para ir filtrando
            ang = {}
            # Itera sobre los servicios y sus listas de canales analógicos
            for service, channel_list in info['Analógico'].items():
                pending_channels = []
                for channel in channel_list:
                    # Se añade el canal a la lista de pendientes si no ha sido medido,
                    # o si fue medido desde una estación diferente (caso SFN)
                    if channel not in alrd_ang or station_ang.get(channel) != station:
                        pending_channels.append(channel)
                
                # Si hay canales pendientes para este servicio, se añaden al diccionario 'ang'
                if pending_channels:
                    ang[service] = pending_channels

            dig = {}
            # Itera sobre los servicios y sus listas de canales digitales
            for service, channel_list in info['Digital'].items():
                pending_channels = []
                for channel in channel_list:
                    # Se añade el canal a la lista de pendientes si no ha sido medido,
                    # o si fue medido desde una estación diferente (caso SFN)
                    if channel not in alrd_dig or station_dig.get(channel) != station:
                        pending_channels.append(channel)

                # Si hay canales pendientes para este servicio, se añaden al diccionario 'dig'
                if pending_channels:
                    dig[service] = pending_channels

            # Guardar la estación sólo si queda algo por medir
            if ang or dig:
                updated_dictionary[station] = {
                    'Acimuth': info['Acimuth'],
                    'Analógico': ang,
                    'Digital': dig
                }

        return updated_dictionary, atv_result, dtv_result


    def tv_measurement(self, dictionary: dict, park_acimuth: int, path: str, 
                    callback_rotate=None, callback_confirm=None, callback_progress=None):
        photos_path = 'Fotos y videos punto de medición'
        supports_path = 'Soportes punto de medición'

        # Se crea la carpeta de entorno
        os.makedirs(f'{path}/{photos_path}/Entorno', exist_ok=True)

        # Diccionarios de resultado
        if os.path.exists(f"{path}/savefiles/tv_progress.json"):
            dictionary, atv_result, dtv_result = self.load_tv_progress(path, dictionary)
        else:
            atv_result = {}
            dtv_result = {}
        
        # Calcular número total de operaciones (para la barra de progreso)
        total_operations = 0
        current_operation = 0
        
        # Contamos las operaciones totales: una rotación y una medición por cada canal
        for station in dictionary.keys():
            total_operations += 1  # Por la rotación de la antena
            analog_measurement = dictionary[station]['Analógico']
            digital_measurement = dictionary[station]['Digital']
            total_operations += len(analog_measurement)  # Por las mediciones analógicas
            total_operations += len(digital_measurement)  # Por las mediciones digitales
        
        # Reportar progreso inicial
        if callback_progress:
            callback_progress(current_operation, total_operations, f"Iniciando mediciones...")

        # Se recorre el diccionario de medidas
        for station in dictionary.keys():
            try:
                acimuth = dictionary[station]['Acimuth']
                analog_measurement = dictionary[station]['Analógico']
                digital_measurement = dictionary[station]['Digital']

                # Reportar progreso antes de rotar
                if callback_progress:
                    callback_progress(current_operation, total_operations, 
                                    f"Rotando antena hacia acimuth {acimuth}° para estación {station}...")

                # Intenta rotar la antena
                rotate_result = self._rotate(park_acimuth, acimuth)
                
                # Si se necesita rotación manual
                if isinstance(rotate_result, tuple) and rotate_result[0] is False:
                    target_acimuth = rotate_result[1]
                    
                    # Si se proporcionó una función callback, la llamamos
                    if callback_rotate:
                        callback_rotate(target_acimuth)
                
                # Incrementar operación y reportar progreso
                current_operation += 1
                if callback_progress:
                    callback_progress(current_operation, total_operations, 
                                    f"Rotación completada. Preparando mediciones para {station}...")
                
                # Medición de los canales analógicos para cada estación
                for service_name, channel_list in analog_measurement.items():
                    for channel in channel_list:
                        while True:  # Bucle para repetir la medición si es necesario
                            # self.atv.reconnect()
                            # Reportar progreso antes de medir
                            if callback_progress:
                                callback_progress(current_operation, total_operations, 
                                                f"Midiendo canal analógico {channel} ({service_name})...")
                            
                            # Definición de los nombres de las carpetas que se crean por cada canal
                            photos_channel_path = rpath(f'{path}/{photos_path}/{station}/CH_{channel}_A_{service_name}')
                            supports_channel_path = rpath(f'{path}/{supports_path}/{station}/CH_{channel}_A_{service_name}')

                            # Creación de las carpetas
                            os.makedirs(photos_channel_path, exist_ok=True)
                            os.makedirs(supports_channel_path, exist_ok=True)

                            # Medición
                            atv_channel_result = self.atv.atv_measurement(channel, supports_channel_path)
                            atv_channel_result.update({'station': station, 'service_name': service_name})

                            # Se añade al diccionario de resultado
                            atv_result[channel] = atv_channel_result
                            
                            # Incrementar operación y reportar progreso
                            current_operation += 1
                            if callback_progress:
                                callback_progress(current_operation, total_operations, 
                                                f"Medición de canal analógico {channel} ({service_name}) completada.")
                            
                            # Solicitar confirmación al usuario
                            if callback_confirm:
                                confirm_result = callback_confirm(f"Canal analógico {channel} ({service_name})")
                                if confirm_result:  # Si el usuario confirma, salimos del bucle while
                                    break
                                # Si no confirma, se repite el ciclo pero no incrementamos el contador
                                current_operation -= 1  # Restamos porque vamos a repetir la operación
                                if callback_progress:
                                    callback_progress(current_operation, total_operations, 
                                                    f"Repitiendo medición del canal analógico {channel} ({service_name})...")
                            else:
                                break  # Si no hay callback de confirmación, salimos del bucle

                # Medición de los canales digitales para cada estación
                for service_name, channel_list in digital_measurement.items():
                    for channel in channel_list:
                        while True:  # Bucle para repetir la medición si es necesario
                            # self.dtv.reconnect()
                            # Reportar progreso antes de medir
                            if callback_progress:
                                callback_progress(current_operation, total_operations, 
                                                f"Midiendo canal digital {channel} ({service_name})...")
                            
                            # Creación de carpetas
                            for channel_name in TV_SERVICES[service_name]:
                                photos_channel_path = rpath(f'{path}/{photos_path}/{station}/CH_{channel}_D_{service_name}/{channel_name}')
                                os.makedirs(photos_channel_path, exist_ok=True)

                            supports_channel_path = rpath(f'{path}/{supports_path}/{station}/CH_{channel}_D_{service_name}')
                            os.makedirs(supports_channel_path, exist_ok=True)

                            # Medición
                            dtv_channel_result = self.dtv.dtv_measurement(channel, supports_channel_path, service_name)
                            dtv_channel_result.update({'station': station, 'service_name': service_name})

                            # Se añade al diccionario de resultado
                            dtv_result[channel] = dtv_channel_result
                            
                            # Incrementar operación y reportar progreso
                            current_operation += 1
                            if callback_progress:
                                callback_progress(current_operation, total_operations, 
                                                f"Medición de canal digital {channel} ({service_name}) completada.")
                            
                            # Solicitar confirmación al usuario
                            if callback_confirm:
                                confirm_result = callback_confirm(f"Canal digital {channel} ({service_name})")
                                if confirm_result:  # Si el usuario confirma, salimos del bucle while
                                    break
                                # Si no confirma, se repite el ciclo pero no incrementamos el contador
                                current_operation -= 1  # Restamos porque vamos a repetir la operación
                                if callback_progress:
                                    callback_progress(current_operation, total_operations, 
                                                    f"Repitiendo medición del canal digital {channel} ({service_name})...")
                            else:
                                break  # Si no hay callback de confirmación, salimos del bucle

            except Exception as e:
                # Cualquier error que ocurra durante la medición, el progreso se guarda
                self.save_tv_progress(path, atv_result, dtv_result)
                if callback_progress:
                    callback_progress(current_operation, total_operations,
                                      f"Error en la medición. Progreso guardado en {path}/savefiles/")
                raise e
                                
        # Reportar progreso final (100%)
        if callback_progress:
            callback_progress(total_operations, total_operations, "¡Medición completada con éxito!")

        # Se guardan los resultados, en caso de algún error durante el diligenciamiento del excel
        self.save_tv_progress(path, atv_result, dtv_result)

        return atv_result, dtv_result
    

    def mbk_measurement(self, path, progress_callback = None):
        try:
            self.mbk.reset()

            # Se crea la carpeta de soportes
            os.makedirs(f"{path}", exist_ok=True)

            # El diccionario de medidas depende del modelo del instrumento
            if self.mbk.instrument_model_name == "ETL":
                bank_dict = BANDS_ETL
            elif self.mbk.instrument_model_name == "ONA-800":
                bank_dict = BANDS_VIAVI
            else:
                bank_dict = BANDS_FXH
                
        
            total_bands = len(bank_dict.keys())
            current_band = 0

            # Se obtienen las coordenadas
            progress_callback(0, 1, "Obteniendo coordenadas")
            latitude, longitude = self.mbk.get_coordinates()
            latitude, longitude = self.mbk.decimal_coords_to_dms(latitude, longitude)

            # Medición y almacenamiento de soportes
            for band in bank_dict.keys():
                current_band += 1
                progress_callback(current_band/total_bands, 1, f"Midiendo la banda {band}")
                self.mbk.continuous_measurement_bank(band, path, latitude, longitude)

            progress_callback(1, 1, "¡Medición de banco completada con éxito")
        except Exception as e:
            progress_callback(1, 1, f"Error en la medición de banco.")
            raise e
    

if __name__ == '__main__':
    # atv_instrument = EtlManager('172.23.82.20', 50, ['HE200'])
    # dtv_instrument = EtlManager('172.23.82.20', 75, ['TELEVES', 'CABLE TELEVES'])
    # mbk_instrument = EtlManager('172.23.82.20', 50, ['HL223'])
    # rtr_instrument = None

    measurement_manager = MeasurementManager(atv=None, dtv=None, mbk=None, rtr=None)

    # sfn_dic = {16: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254}, 17: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254}, 14: {'Suba': 153, 'Manjui': 254}, 15: {'Suba': 153, 'Manjui': 254}}

    # sfn_result = measurement_manager.sfn_measurement(sfn_dic, './tests', 0)
    # print(sfn_result)

    diccionario_medicion = {
        'Tibitóc': {'Acimuth': 61, 'Analógico': {'Canal 1': 3, 'Canal Institucional': 6, 'Señal Colombia': 12}, 'Digital': {}},
        'Suba': {'Acimuth': 153, 'Analógico': {'RCN': 8, 'Caracol': 10, 'CityTV': 21}, 'Digital': {'CityTV': 27}},
        'Calatrava': {'Acimuth': 157, 'Analógico': {'Teveandina': 23, 'Señal Colombia': 25, 'Canal Capital': 32, 'Canal 1': 36, 'Canal Institucional': 38}, 'Digital': {'Canal Capital': 28}},
        'Manjui': {'Acimuth': 254, 'Analógico': {'Canal Capital': 2, 'RCN': 4, 'Caracol': 5, 'Canal 1': 7, 'Canal Institucional': 9, 'Señal Colombia': 11}, 'Digital': {'Caracol': 14, 'RCN': 15, 'RTVC': 16, 'Teveandina': 17}}}

    # atv_dic, dtv_dic = measurement_manager.tv_measurement(diccionario_medicion, 0, './tests')
    # print('Diccionario analógico')
    # print(atv_dic)
    # print('\n')
    # print('Diccionario Digital')
    # print(dtv_dic)