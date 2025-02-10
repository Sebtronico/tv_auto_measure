import pandas as pd
from ..utils.constants import *

class LeerPreingenieria:
    def __init__(self, filename: str):

        # Cargue de archivo de preingeniería
        self.filename_pre_engineering = filename
        self.main_coordinates          = pd.read_excel(filename, sheet_name = 0, skiprows = 1)
        self.main_channelization       = pd.read_excel(filename, sheet_name = 1, skiprows = 1)
        self.additional_channelization = pd.read_excel(filename, sheet_name = 2, skiprows = 1)
        self.additional_coordinates    = pd.read_excel(filename, sheet_name = 3, skiprows = 0)

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
        self.additional_channelization[['Estación Regional TV Analógica', 'Operador Regional']] = self.additional_channelization['Estación Regional TV Analógica'].str.split(" - ", expand=True)

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
        null_indexes_additional = self.additional_channelization.index[self.additional_channelization['Operador Regional'].isnull() == True].tolist()
        for index in null_indexes_additional:
            department = self.additional_channelization.at[index,'Departamento']
            null_municipality_additional = self.additional_channelization.at[index,'Municipio']
            regional_index = self.regionals.index[(self.regionals['Municipio'] == null_municipality_additional) & (self.regionals['Departamento'] == department)].tolist()[0]
            regional_channel = self.regionals.at[regional_index, 'Operador']
            self.additional_channelization.loc[index,'Operador Regional'] = regional_channel

        self.main_iterator = list(SEARCH_PRINCIPALS['Analógico'].keys()) + list(SEARCH_PRINCIPALS['Digital'].keys())
        self.additional_iterator = list(SEARCH_ADDITIONALS['Analógico'].keys()) + list(SEARCH_ADDITIONALS['Digital'].keys())


    def get_municipios(self):
        return self.municipalities
    

    def get_number_of_points(self, municipio: str):
        number_of_points = self.main_coordinates[self.main_coordinates['Municipio'] == municipio]['Pto.'].max()

        return number_of_points