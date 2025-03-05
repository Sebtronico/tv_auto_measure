from InstrumentController import *
from src.utils.constants import *
from collections import defaultdict
from tkinter import messagebox
import os

class MeasurementManager:
    def __init__(self, dtv: EtlManager, atv: EtlManager, mbk: EtlManager|FPHManager, rtr: MSDManager|None):
        self.dtv = dtv
        self.atv = atv
        self.mbk = mbk
        self.rtr = rtr


    def _rotate(self, park_acimuth: int, station_acimuth: int):
        if self.rtr is not None:
            self.rtr.move_rotor(park_acimuth, station_acimuth)
        else:
            messagebox.showinfo(message=f'Gire el rotor hacia el acimuth {station_acimuth}.\n Una vez apuntado, haga click en aceptar.')


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
    

    def sfn_measurement(self, dictionary: dict, path: str, park_acimuth: int):
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
        sfn_dictionary = self._group_sfn_dictionary(dictionary)
        sfn_result = {}
        
        for channels_tuple, stations in sfn_dictionary.items():
            # Envío de comando de reset al principio de cada medición.
            # Esto evita que las trazas queden activas.
            self.dtv.reset()

            # Configuración inicial del analizador.
            self.dtv.sfn_setup()

            #
            channels = sorted(list(channels_tuple))

            # Creación del archivo .txt
            filename = f"{path}/SFN_CH{' - '.join(map(str, channels))}"
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


                # Se gira el rotor hacia el acimuth de la estación y se activa la traza
                self._rotate(park_acimuth, acimuth) 
                self.dtv.write_str(f'DISP:TRAC{trace}:MODE AVER')
                self.dtv.write_str(f'DET{trace} RMS')
                self.dtv.write_str(f'INIT:CONT OFF')
                self.dtv.write_str(f'SWE:COUN 10')
                self.dtv.write_str(f'INIT;*WAI')
                self.dtv.write_str(f'DISP:TRAC{trace}:MODE VIEW')

                # Se obtienen las potencias de canal
                self.dtv.write_str(f'POW:TRAC {trace}')
                trace_powers = self.dtv.query_bin_or_ascii_float_list_with_opc('CALCulate:MARKer:FUNCtion:POWer:RESult? MCACpower')[:-1]

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

        return final_sfn_result
    

    def tv_measurement(self, dictionary: dict, park_acimuth: int, path: str):
        photos_path = 'Fotos y videos punto de medición' # Nombre de la carpeta de fotos y videos
        supports_path = 'Soportes punto de medición' # Nombre de la carpeta de soportes de la medición

        # Se crea la carpeta de entorno
        os.makedirs(f'{path}/{photos_path}/Entorno', exist_ok=True)

        # Diccionarios de resultado
        atv_result = {}
        dtv_result = {}

        # Se recorre el diccionario de medidas
        for station in dictionary.keys(): 
            acimuth = dictionary[station]['Acimuth']
            analog_measurement = dictionary[station]['Analógico']
            digital_measurement = dictionary[station]['Digital']

            self._rotate(park_acimuth, acimuth)

            # Medición de los canales analógicos para cada estación
            for service_name, channel in analog_measurement.items():
                # Definición de los nombres de las carpetas que se crean por cada canal
                photos_channel_path = f'{path}/{photos_path}/{station}/CH_{channel}_A_{service_name}'
                supports_channel_path = f'{path}/{supports_path}/{station}/CH_{channel}_A_{service_name}'

                # Creación de las carpetas
                os.makedirs(photos_channel_path, exist_ok=True)
                os.makedirs(supports_channel_path, exist_ok=True)

                # Medición
                atv_channel_result = self.atv.atv_measurement(channel, supports_channel_path)
                atv_channel_result.update({'station': station, 'service_name': service_name})

                # Se añade al diccionario de resultado
                atv_result[channel] = atv_channel_result

            # Medición de los canales digitales para cada estación
            for service_name, channel in digital_measurement.items():
                # Creación de carpetas
                for channel_name in TV_SERVICES[service_name]:
                    photos_channel_path = f'{path}/{photos_path}/{station}/CH_{channel}_D_{service_name}/{channel_name}'
                    os.makedirs(photos_channel_path, exist_ok=True)

                supports_channel_path = f'{path}/{supports_path}/{station}/CH_{channel}_D_{service_name}'
                os.makedirs(supports_channel_path, exist_ok=True)

                # Medición
                dtv_channel_result = self.dtv.dtv_measurement(channel, supports_channel_path, service_name)
                dtv_channel_result.update({'station': station, 'service_name': service_name})

                # Se añade al diccionario de resultado
                dtv_result[channel] = dtv_channel_result

        return atv_result, dtv_result
    

if __name__ == '__main__':
    atv_instrument = EtlManager('172.23.82.39', 50, ['HE200'])
    dtv_instrument = EtlManager('172.23.82.39', 75, ['TELEVES', 'CABLE TELEVES'])
    mbk_instrument = EtlManager('172.23.82.39', 50, ['HL223'])
    rtr_instrument = None

    measurement_manager = MeasurementManager(atv=atv_instrument, dtv=dtv_instrument, mbk=mbk_instrument, rtr=rtr_instrument)

    # sfn_dic = {16: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254}, 17: {'Tibitóc': 61, 'Calatrava': 157, 'Manjui': 254}, 14: {'Suba': 153, 'Manjui': 254}, 15: {'Suba': 153, 'Manjui': 254}}

    # sfn_result = measurement_manager.sfn_measurement(sfn_dic, './tests', 0)
    # print(sfn_result)

    diccionario_medicion = {
        # 'Asignado Sin Estación': {'Acimuth': 0, 'Analógico': {}, 'Digital': {}},
        # 'Tibitóc': {'Acimuth': 61, 'Analógico': {'Canal 1': 3, 'Canal Institucional': 6, 'Señal Colombia': 12}, 'Digital': {}},
        # 'Suba': {'Acimuth': 153, 'Analógico': {'RCN': 8, 'Caracol': 10, 'CityTV': 21}, 'Digital': {'CityTV': 27}},
        # 'Calatrava': {'Acimuth': 157, 'Analógico': {'Teveandina': 23, 'Señal Colombia': 25, 'Canal Capital': 32, 'Canal 1': 36, 'Canal Institucional': 38}, 'Digital': {'Canal Capital': 28}},
        'Manjui': {'Acimuth': 254, 'Analógico': {'Canal Capital': 2, 'RCN': 4, 'Caracol': 5, 'Canal 1': 7, 'Canal Institucional': 9, 'Señal Colombia': 11}, 'Digital': {'Caracol': 14, 'RCN': 15, 'RTVC': 16, 'Teveandina': 17}}}

    measurement_manager.tv_measurement(diccionario_medicion, 0, './tests')