import pandas as pd
from src.utils.constants import *

class ReadExcel:
    def __init__(self, filename: str):

        # Cargue de archivo de preingeniería
        self.filename_pre_engineering = filename
        self.main_coordinates    = pd.read_excel(filename, sheet_name = 0, skiprows = 1)
        self.main_channelization = pd.read_excel(filename, sheet_name = 1, skiprows = 1)
        self.addt_channelization = pd.read_excel(filename, sheet_name = 2, skiprows = 1)
        self.addt_coordinates    = pd.read_excel(filename, sheet_name = 3, skiprows = 0)

        # Cargue de archivo de referencias
        self.filename_references = './src/utils/Referencias.xlsx'
        self.regionals = pd.read_excel(self.filename_references, sheet_name = 0)
        self.engineers = pd.read_excel(self.filename_references, sheet_name = 1)
        self.stations  = pd.read_excel(self.filename_references, sheet_name = 2)

        # Municipios dentro de la preingeniería
        self.municipalities = self.main_channelization['Municipio'].tolist()
        
        # Renombramiento de columnas, para corregir espacios en blanco
        self.main_coordinates.rename(columns={' RTVC':'RTVC'}, inplace=True)

        # Corrección de valores diferentes para estaciones no asignadas.
        self.main_channelization.replace('Asignados sin Estación', 'Asignado Sin Estación', inplace=True)

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

        # Corrección de valores nulos en la columna de operador regional
        null_indexes_addt = self.addt_channelization.index[self.addt_channelization['Operador Regional'].isnull() == True].tolist()
        for index in null_indexes_addt:
            department = self.addt_channelization.at[index,'Departamento']
            null_municipality_addt = self.addt_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == null_municipality_addt) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.addt_channelization.loc[index,'Operador Regional'] = regional_channel

        self.main_iterator = list(SEARCH_PRINCIPALS['Analógico'].keys()) + list(SEARCH_PRINCIPALS['Digital'].keys())
        self.addt_iterator = list(SEARCH_ADDITIONALS['Analógico'].keys()) + list(SEARCH_ADDITIONALS['Digital'].keys())

    # Retorna la lista de municipios
    def get_municipalities(self):
        return self.main_channelization['Municipio'].tolist()
    
    # Retorna el número de puntos que hay en un municipio
    def get_number_of_points(self, municipio: str):
        return self.main_coordinates[self.main_coordinates['Municipio'] == municipio]['Pto.'].max()
    
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
                            # Pendiente añadir lo de segundo regional
                            if search[tec][station_column][service] == 'Canal Regional':
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
                                dictionary[station][tec].update({channel: int(dataframe.at[index, service])})
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
    
    @staticmethod
    def clean_asigned_without_station(dictionary: dict):
        if 'Asignado Sin Estación' not in dictionary:
            return dictionary  # Si no existe la clave, devolvemos el diccionario sin cambios

        # Obtener los datos de "Asignado Sin Estación"
        datos_ase = dictionary['Asignado Sin Estación']

        # Recopilar todos los valores de "Analógico" en las demás estaciones
        valores_analogicos_presentes = {}
        for estacion, info in dictionary.items():
            if estacion != 'Asignado Sin Estación' and 'Analógico' in info:
                for canal, numero in info['Analógico'].items():
                    valores_analogicos_presentes[numero] = canal  # Guardamos por número de canal

        # Filtrar las entradas de "Analógico" en 'Asignado Sin Estación' que ya existen en otra estación
        if 'Analógico' in datos_ase:
            datos_ase['Analógico'] = {
                canal: numero for canal, numero in datos_ase['Analógico'].items()
                if numero not in valores_analogicos_presentes
            }

        # Si después de limpiar 'Asignado Sin Estación', no quedan valores en 'Analógico' y 'Digital', eliminarlo
        if not datos_ase['Analógico'] and not datos_ase['Digital']:
            del dictionary['Asignado Sin Estación']

        return dictionary

    # Realiza todo el proceso de creación y llenado del diccionario. 
    def get_dictionary(self, municipality: str, point: int):
        main_index = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()
        addt_index = self.addt_channelization.index[self.addt_channelization['Municipio'] == municipality].tolist()

        main_stations = [str(self.main_channelization.at[main_index[0],key]).title() for key in self.main_iterator if str(self.main_channelization.at[main_index[0],key]).title() not in ['Sin Obligación', 'No Tiene']]
        addt_stations = [str(self.addt_channelization.at[index,key]).title() for index in addt_index for key in self.addt_iterator if str(self.addt_channelization.at[index,key]).title() not in ['No Aplica']]

        stations = list(set(main_stations + addt_stations))

        dictionary = {station: {'Acimuth': 0, 'Analógico': {}, 'Digital': {}} for station in stations}

        dictionary = self.fill_dictionary(self.main_channelization, stations, main_index, SEARCH_PRINCIPALS,  dictionary)
        dictionary = self.fill_dictionary(self.addt_channelization, stations, addt_index, SEARCH_ADDITIONALS, dictionary)

        dictionary = self.debug_dictionary(dictionary)

        stations  = list(dictionary.keys())

        dictionary = self.fill_acimuth(municipality, point, stations, self.main_coordinates, PRINCIPAL,  dictionary)
        dictionary = self.fill_acimuth(municipality, point, stations, self.addt_coordinates, ADDITIONAL, dictionary)

        dictionary = self.clean_asigned_without_station(dictionary)

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
    def update_sfn(dictionary, selection):

        # Selecciones de ejemplo
        # selecciones = {
        #     16: 'Boquerón De Chipaque',
        #     17: 'El Tigre'
        # }

        # Eliminar los canales no seleccionados del diccionario original
        for numero, estacion_seleccionada in selection.items():
            for estacion, info in dictionary.items():
                if 'Digital' in info and numero in info['Digital']:
                    if estacion != estacion_seleccionada:
                        del info['Digital'][numero]

        return dictionary

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
    filename = './templates/Preingenieria Meta.xlsx'
    preingenieria = ReadExcel(filename)
    dictionary = preingenieria.get_dictionary('El Dorado', 1)
    print(dictionary)
    print('\n')
    # print(preingenieria.get_excel_station_list(dictionary))
    # sfn = preingenieria.get_sfn(dictionary)
    # print(sfn)
    # selection = {
    #     16: 'Boquerón De Chipaque',
    #     17: 'El Tigre'
    # }
    # updated_dictionary = preingenieria.update_sfn(dictionary, selection)
    # print(updated_dictionary)