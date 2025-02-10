import pandas as pd
from ..utils.constants import *

class LeerPreingenieria:
    def __init__(self, filename: str):

        # Cargue de archivo de preingeniería
        self.filename_pre_engineering = filename
        self.main_coordinates          = pd.read_excel(filename, sheet_name = 0, skiprows = 1)
        self.main_channelization       = pd.read_excel(filename, sheet_name = 1, skiprows = 1)
        self.addl_channelization = pd.read_excel(filename, sheet_name = 2, skiprows = 1)
        self.addl_coordinates    = pd.read_excel(filename, sheet_name = 3, skiprows = 0)

        # Cargue de archivo de referencias
        self.filename_references = './src/utils/Referencias.xlsx'
        self.regionals = pd.read_excel(self.filename_references, sheet_name = 0)
        self.engineers = pd.read_excel(self.filename_references, sheet_name = 1)
        self.stations  = pd.read_excel(self.filename_references, sheet_name = 2)

        # Municipios dentro de la preingeniería
        self.municipalities = self.main_channelization['Municipio'].tolist()
        
        # Renombramiento de columnas, para corregir espacios en blanco
        self.main_coordinates.rename(columns={' RTVC':'RTVC'}, inplace=True)

        # Separación de la columna de canal regional
        self.main_channelization[['Estación Regional TV Analógica', 'Operador Regional']] = self.main_channelization['Estación Regional TV Analógica'].str.split(" - ", expand=True)
        self.addl_channelization[['Estación Regional TV Analógica', 'Operador Regional']] = self.addl_channelization['Estación Regional TV Analógica'].str.split(" - ", expand=True)

        # Borra las estaciones que no tienen obligación de canalizar
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
        null_indexes_addl = self.addl_channelization.index[self.addl_channelization['Operador Regional'].isnull() == True].tolist()
        for index in null_indexes_addl:
            department = self.addl_channelization.at[index,'Departamento']
            null_municipality_addl = self.addl_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == null_municipality_addl) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.addl_channelization.loc[index,'Operador Regional'] = regional_channel

        self.main_iterator = list(SEARCH_PRINCIPALS['Analógico'].keys()) + list(SEARCH_PRINCIPALS['Digital'].keys())
        self.addl_iterator = list(SEARCH_ADDITIONALS['Analógico'].keys()) + list(SEARCH_ADDITIONALS['Digital'].keys())


    def get_municipalities(self):
        return self.main_channelization['Municipio'].tolist()
    

    def get_number_of_points(self, municipio: str):
        return self.main_coordinates[self.main_coordinates['Municipio'] == municipio]['Pto.'].max()
    

    def get_engineers_list(self):
        return self.engineers['Ingeniero'].tolist()
    

    @staticmethod
    def fill_dictionary(dataframe: pd.DataFrame, stations: list, indexes: int, search: dict, dictionary: dict):
        for tec in search.keys(): # Análogo o Digital
            if tec in ['Analógico', 'Digital']:
                for station_column in search[tec].keys(): # Estación Públicos TV Analógica
                    for service in search[tec][station_column].keys(): # C1, CI, SC
                        for index in indexes:
                            station = dataframe.at[index, station_column].title()
                            if search[tec][station_column][service] == 'Canal Regional':
                                channel = dataframe.at[index, 'Operador Regional']
                            else:
                                channel = search[tec][station_column][service]
                            if station in stations:
                                dictionary[station][tec].update({channel: int(dataframe.at[index, service])})
            else:
                pass

        return dictionary
    

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
    

    def get_dictionary(self, municipality: str, point: int):
        main_index = self.main_channelization.index[self.main_channelization['Municipio'] == municipality].tolist()
        addl_index = self.addl_channelization.index[self.addl_channelization['Municipio'] == municipality].tolist()

        main_stations = [self.main_channelization.at[addl_index[0],key].title() for key in self.main_iterator if self.main_channelization.at[main_index[0],key].title() not in ['Sin Obligación', 'No Tiene']]
        addl_stations = [self.addl_channelization.at[index,key].title() for index in addl_index for key in self.addl_iterator if self.addl_channelization.at[index,key].title() not in ['No Aplica']]

        stations = list(set(main_stations + addl_stations))

        dictionary = {station: {'Acimuth': 0, 'Analógico': {}, 'Digital': {}} for station in stations}

        dictionary = self.fill_dictionary(self.main_channelization, stations, main_index, SEARCH_PRINCIPALS,  dictionary)
        dictionary = self.fill_dictionary(self.addl_channelization, stations, addl_index, SEARCH_ADDITIONALS, dictionary)

        dictionary = self.debug_dictionary(dictionary)

        stations  = list(dictionary.keys())

        dictionary = self.fill_acimuth(municipality, point, stations, self.main_coordinates, PRINCIPAL,  dictionary)
        dictionary = self.fill_acimuth(municipality, point, stations, self.addl_coordinates, ADDITIONAL, dictionary)

        return dictionary
    

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