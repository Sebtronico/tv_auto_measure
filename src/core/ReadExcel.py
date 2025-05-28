import pandas as pd
from collections import OrderedDict
from src.utils.constants import *
from utils import rpath

class ReadExcel:
    def __init__(self, filename: str):

        # Cargue de archivo de preingeniería
        self.filename_pre_engineering = filename
        self.main_coordinates    = pd.read_excel(filename, sheet_name = 0, skiprows = 1)
        self.main_channelization = pd.read_excel(filename, sheet_name = 1, skiprows = 1)
        self.addt_channelization = pd.read_excel(filename, sheet_name = 2, skiprows = 1)
        self.addt_coordinates    = pd.read_excel(filename, sheet_name = 3, skiprows = 0)

        # Cargue de archivo de referencias
        self.filename_references = rpath('./src/utils/Referencias.xlsx')
        self.regionals = pd.read_excel(self.filename_references, sheet_name = 0)
        self.engineers = pd.read_excel(self.filename_references, sheet_name = 1)
        self.stations  = pd.read_excel(self.filename_references, sheet_name = 2)

        # Municipios dentro de la preingeniería
        self.municipalities = self.main_channelization['Municipio'].tolist()
        
        # Renombramiento de columnas, para corregir espacios en blanco
        self.main_coordinates.rename(columns={' RTVC':'RTVC'}, inplace=True)

        # Corrección de valores diferentes para estaciones no asignadas.
        self.main_channelization.replace('Asignados sin Estación', 'Asignado Sin Estación', inplace=True)

        # Corrección de lectura de city tv, en caso de existir
        self.addt_channelization.replace('City TV', 'CityTV', inplace=True)

        # Separación de la columna de canal regional
        self.separate_regional_channel(self.main_channelization, 'Estación Regional TV Analógica', 'Operador Regional')
        self.separate_regional_channel(self.addt_channelization, 'Estación Regional TV Analógica', 'Operador Regional')
        self.separate_regional_channel(self.addt_channelization, 'Estación 2do_CR', 'Segundo Operador Regional')
        self.separate_regional_channel(self.addt_channelization, 'Estación TDT 2do_CR', 'Segundo Operador Regional TDT')

        # Borra las estaciones que no tienen obligación de medir
        for municipality in self.municipalities:
            exception_index = self.main_coordinates.index[self.main_coordinates['Municipio'] == municipality].tolist()[0]

            if self.main_coordinates['PWR\ndBm'].isnull().tolist()[exception_index]:
                delete_index = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()[0]
                self.main_channelization.loc[delete_index,'Estación TDT CCNP'] = 'SIN OBLIGACIÓN'
                self.main_channelization.loc[delete_index,'RCND'] = pd.NA
                self.main_channelization.loc[delete_index,'CRCD'] = pd.NA

        # Corrección de valores nulos en la columna de operador regional
        null_indexes_main = self.main_channelization.index[self.main_channelization['Operador Regional'].isnull() == True].tolist()
        for index in null_indexes_main:
            department = self.main_channelization.at[index,'Departamento']
            null_municipality_main = self.main_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == null_municipality_main) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.main_channelization.loc[index,'Operador Regional'] = regional_channel

        # Creación de columna de canal regional 1 para canales digitales
        null_indexes_addt = self.addt_channelization.index[self.addt_channelization['Operador Regional'].isnull() == True].tolist()
        for index in null_indexes_addt:
            department = self.addt_channelization.at[index,'Departamento']
            null_municipality_addt = self.addt_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == null_municipality_addt) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.addt_channelization.loc[index,'Operador Regional'] = regional_channel
            
        # Creación de columna de canal regional 1 para canales analógicos
        self.main_channelization['Regional Municipio'] = None
        all_indexes_main = self.main_channelization.index.tolist()
        for index in all_indexes_main:
            department = self.main_channelization.at[index,'Departamento']
            municipality_main = self.main_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == municipality_main) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.main_channelization.at[index,'Regional Municipio'] = regional_channel

        # Creación de columna de canal regional 1 para canales analógicos
        self.addt_channelization['Regional Municipio'] = None
        all_indexes_addt = self.addt_channelization.index.tolist()
        for index in all_indexes_addt:
            department = self.addt_channelization.at[index,'Departamento']
            municipality_addt = self.addt_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == municipality_addt) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.addt_channelization.at[index,'Regional Municipio'] = regional_channel

        self.main_iterator = list(SEARCH_PRINCIPALS['Analógico'].keys()) + list(SEARCH_PRINCIPALS['Digital'].keys())
        self.addt_iterator = list(SEARCH_ADDITIONALS['Analógico'].keys()) + list(SEARCH_ADDITIONALS['Digital'].keys())

    # Retorna la lista de municipios
    def get_municipalities(self):
        return self.main_channelization['Municipio'].tolist()
    

    # Retorna el departamento correspondiente a un municipio
    def get_department(self, municipality: str):
        index_municipality = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()[0]
        return self.main_channelization.at[index_municipality, 'Departamento']


    def get_dane_code(self, municipality: str):
        index_municipality = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()[0]
        return self.main_channelization.at[index_municipality, 'Cód.\nDANE']

    # Retorna el número de puntos que hay en un municipio
    def get_number_of_points(self, municipality: str):
        return self.main_coordinates[self.main_coordinates['Municipio'] == municipality]['Pto.'].max()
    

    # Retorna la lista de ingenieros
    def get_engineers_list(self):
        return self.engineers['Ingeniero'].tolist()
    

    # Hace la separación de las columnas que contienen 'NombreDeLaEstación - Operador'
    @staticmethod
    def separate_regional_channel(dataframe: pd.DataFrame, base_column: str, regional_column: str):
        # Aplicar una máscara para detectar las filas que contienen " - "
        mask = dataframe[base_column].str.contains(" - ", na=False)

        # Verificar si hay filas que cumplen la condición
        if mask.any():
            # Crear las nuevas columnas con la división
            split_data = dataframe.loc[mask, base_column].str.split(" - ", expand=True)

            # Si split_data tiene solo una columna, agregar una segunda con NaN
            if split_data.shape[1] == 1:
                split_data[1] = float('nan')

            # Asignar los valores a las columnas
            dataframe.loc[mask, base_column] = split_data[0]
            dataframe.loc[mask, regional_column] = split_data[1]

        # Asegurar que 'regional_column' tenga NaN donde no se hizo la división
        dataframe.loc[~mask, regional_column] = float('nan')

    
    # Llena el diccionario con los valores que se encuentran en la preingeniería
    @staticmethod
    def fill_dictionary(dataframe: pd.DataFrame, stations: list, indexes: list, search: dict, dictionary: dict):
        for tec in search.keys(): # Análogo o Digital
            if tec in ['Analógico', 'Digital']:
                for station_column in search[tec].keys(): # Estación Públicos TV Analógica
                    for service in search[tec][station_column].keys(): # C1, CI, SC
                        for index in indexes:
                            station = str(dataframe.at[index, station_column]).title()
                            if search[tec][station_column][service] == 'Regional Municipio':
                                channel = dataframe.at[index, 'Regional Municipio']

                            elif search[tec][station_column][service] == 'Canal Regional':
                                channel = dataframe.at[index, 'Operador Regional']

                            elif search[tec][station_column][service] == 'Regional 2':
                                channel = dataframe.at[index, 'Segundo Operador Regional']

                            elif search[tec][station_column][service] == 'Regional 2 TDT':
                                channel = dataframe.at[index, 'Segundo Operador Regional TDT']

                            elif search[tec][station_column][service] == 'Local_1_Analógica':
                                channel = dataframe.at[index, 'Local_1']

                            elif search[tec][station_column][service] == 'Local TDT':
                                channel = dataframe.at[index, 'Local_TDT']

                            else:
                                channel = search[tec][station_column][service]

                            if station in stations:
                                try:
                                    dictionary[station][tec].update({channel: int(dataframe.at[index, service])})
                                except ValueError:
                                    continue
            else:
                pass

        return dictionary
    

    # Depura el diccionario para eliminar estaciones repetidas
    @staticmethod
    def debug_dictionary(dictionary: dict):
        # Obtener todas las claves
        keys = list(dictionary.keys())
        
        for i, key in enumerate(keys):
            for another_key in keys[i+1:]:
                # Verificar si la clave está contenida dentro de otra clave
                if key in another_key:
                    # Combinar diccionarios
                    for tecnologia in ['Analógico', 'Digital']:
                        if tecnologia in dictionary[another_key]:
                            dictionary[key][tecnologia].update(dictionary[another_key][tecnologia])
                    # Eliminar la clave más larga
                    del dictionary[another_key]
                elif another_key in key:
                    # Combinar diccionarios
                    for tecnologia in ['Analógico', 'Digital']:
                        if tecnologia in dictionary[key]:
                            dictionary[another_key][tecnologia].update(dictionary[key][tecnologia])
                    # Eliminar la clave más larga
                    del dictionary[key]
                    break

        return dictionary 
    

    # Llena el diccionario con los valores de acimut
    @staticmethod
    def fill_acimuth(municipality: str, point: int, stations: list, dataframe: pd.DataFrame, columns: dict, dictionary: dict):
        index_acimuth = dataframe.index[(dataframe['Municipio'] == municipality) & (dataframe['Pto.'] == point)].tolist()[0]
        for station in stations:
            for station_column, acimuth_column in columns.items():
                try:
                    if station in dataframe.loc[index_acimuth, station_column].title():
                        try:
                            dictionary[station]['Acimuth'] = int(dataframe.at[index_acimuth, acimuth_column])
                            break
                        except ValueError:
                            continue
                except AttributeError:
                    continue

        return dictionary

    
    # Función para ordenar el diccionario por acimuth y por número de canal
    @staticmethod
    def sort_dictionary(dictionary):
        # Ordenar las estaciones por acimut de menor a mayor
        sorted_stations = sorted(dictionary.items(), key=lambda x: x[1]['Acimuth'])
        
        sorted_dictionary = {}
        for station, data in sorted_stations:
            # Ordenar los diccionarios 'Analógico' y 'Digital' por número de canal
            analogico_ordenado = dict(sorted(data['Analógico'].items(), key=lambda x: x[1]))
            digital_ordenado = dict(sorted(data['Digital'].items(), key=lambda x: x[1]))
            
            # Construir el nuevo diccionario con el mismo formato
            sorted_dictionary[station] = {
                'Acimuth': data['Acimuth'],
                'Analógico': analogico_ordenado,
                'Digital': digital_ordenado
            }
        
        return sorted_dictionary
    

    # Realiza todo el proceso de creación y llenado del diccionario. 
    def get_dictionary(self, municipality: str, point: int):
        # Filtra las estaciones que contienen al municipio seleccionado
        main_index = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()
        addt_index = self.addt_channelization.index[self.addt_channelization['Municipio'] == municipality].tolist()

        # Obtiene la lista de estaciones en ambas hojas
        main_stations = [str(self.main_channelization.at[main_index[0],key]).title() for key in self.main_iterator if str(self.main_channelization.at[main_index[0],key]).title() not in ['Sin Obligación', 'No Tiene']]
        addt_stations = [str(self.addt_channelization.at[index,key]).title() for index in addt_index for key in self.addt_iterator if str(self.addt_channelization.at[index,key]).title() not in ['No Aplica']]

        # Obtiene la lista total de estaciones
        stations = list(set(main_stations + addt_stations))

        # Creación del diccionario en blanco
        dictionary = {station: {'Acimuth': 0, 'Analógico': {}, 'Digital': {}} for station in stations}

        # Llenado del diccionario con ambas hojas
        dictionary = self.fill_dictionary(self.main_channelization, stations, main_index, SEARCH_PRINCIPALS,  dictionary)
        dictionary = self.fill_dictionary(self.addt_channelization, stations, addt_index, SEARCH_ADDITIONALS, dictionary)

        # Listado final de estaciones
        stations  = list(dictionary.keys())

        # Llena el acimuth de todas las estaciones
        dictionary = self.fill_acimuth(municipality, point, stations, self.main_coordinates, PRINCIPAL,  dictionary)
        dictionary = self.fill_acimuth(municipality, point, stations, self.addt_coordinates, ADDITIONAL, dictionary)

        # Eliminación de estaciones repetidas
        dictionary = self.debug_dictionary(dictionary)

        # Ordena las estaciones, primero por orden de acimuth y luego por orden de canal
        dictionary = self.sort_dictionary(dictionary)
        
        return dictionary
    

    # Retorna el diccionario de canales en los que hay que medir SFN.
    @staticmethod
    def get_sfn(dictionary: dict):
        digital_channels = {}

        # Recorremos el diccionario para buscar los canales digitales
        for station, info in dictionary.items():
            acimuth = info.get('Acimuth')
            technologies = {k: v for k, v in info.items() if k != 'Acimuth'}

            if 'Digital' in technologies:
                for number in technologies['Digital'].values():
                    if number not in digital_channels:
                        digital_channels[number] = {}
                    digital_channels[number][station] = acimuth

        # Ordenar cada subdiccionario por acimuth de menor a mayor
        sorted_channels = {
            number: dict(sorted(stations.items(), key=lambda item: item[1]))
            for number, stations in digital_channels.items() if len(stations) > 1
        }

        return sorted_channels
    

    # Actualiza el diccionario para eliminar los canales de menor potencia (obtenidos de la medición SFN).
    @staticmethod
    def update_sfn(diccionario_original, diccionario_eleccion):
        """
        Limpia el diccionario original eliminando los canales digitales repetidos
        basándose en el diccionario de elección.
        
        Args:
            diccionario_original (dict): Diccionario con la información de todas las estaciones
            diccionario_eleccion (dict): Diccionario que indica qué estación se quiere mantener para cada canal
            
        Returns:
            dict: Diccionario original modificado con los canales repetidos eliminados
        """
        # Creamos una copia del diccionario original para no modificarlo directamente
        diccionario_limpio = diccionario_original.copy()
        
        # Iteramos sobre cada estación en el diccionario
        for estacion in diccionario_limpio:
            # Obtenemos el diccionario Digital de la estación actual
            digital = diccionario_limpio[estacion]['Digital']
            
            # Lista para almacenar los canales a eliminar
            canales_a_eliminar = []
            
            # Revisamos cada canal en el diccionario Digital
            for operador, canal in digital.items():
                # Si el canal está en el diccionario de elección
                if canal in diccionario_eleccion:
                    # Si la estación actual no es la elegida para este canal
                    if diccionario_eleccion[canal] != estacion:
                        # Marcamos el canal para eliminar
                        canales_a_eliminar.append(operador)
            
            # Eliminamos los canales marcados
            for operador in canales_a_eliminar:
                del digital[operador]
                
        return diccionario_limpio


    # Retorna la lista de estaciones que van en la hoja de 'Información Gral' del excel de post procesamiento
    @staticmethod
    def get_excel_station_list(dictionary: dict):
        excel_station_list = []
        for estacion, technology in dictionary.items():
            if 'Digital' in technology:
                for service in technology['Digital']:
                    if service in ['Caracol', 'RCN']:
                        excel_station_list.append(f"{estacion.upper()} - CCNP")
                    else:
                        excel_station_list.append(f"{estacion.upper()} - RTVC")
        excel_station_list = sorted(list(set(excel_station_list)))
        
        return excel_station_list
    
if __name__ == '__main__':
    filename = './tests/Preingenieria Cundinamarca.xlsx'
    excel = ReadExcel(filename)


    municipio = 'Subachoque'
    punto = 1

    print(excel.get_dane_code(municipio))
    # print(type(excel.get_department(municipio)))

    # dictionary = excel.get_dictionary(municipio, punto)

    # sfn = excel.get_sfn(dictionary)
    # print(sfn)

    # seleccion = {16: 'Manjui', 17: 'Manjui', 14: 'Manjui', 15: 'Manjui'}

    # dictionary2 = excel.update_sfn(dictionary, seleccion)
    # print(dictionary2)