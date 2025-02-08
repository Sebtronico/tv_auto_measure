import pandas as pd
from ..utils.constants import *

class LeerPreingenieria:
    def __init__(self, filename: str):
        self.filename_preingenieria = filename
        self.coordenadas_principal  = pd.read_excel(filename, sheet_name = 0, skiprows = 1)
        self.canalizacion_principal = pd.read_excel(filename, sheet_name = 1, skiprows = 1)
        self.canalizacion_adicional = pd.read_excel(filename, sheet_name = 2, skiprows = 1)
        self.coordenadas_adicional  = pd.read_excel(filename, sheet_name = 3, skiprows = 0)

        self.filename_referencias = '.\\Formatos\\Referencias.xlsx'
        self.regionales = pd.read_excel(self.filename_referencias, sheet_name = 0)
        self.ingenieros = pd.read_excel(self.filename_referencias, sheet_name = 1)
        self.estaciones = pd.read_excel(self.filename_referencias, sheet_name = 2)

        self.canalizacion_principal[['Estación Regional TV Analógica', 'Operador Regional']] = self.canalizacion_principal['Estación Regional TV Analógica'].str.split(" - ", expand=True)
        index_nulos_principal = self.canalizacion_principal.index[self.canalizacion_principal['Operador Regional'].isnull() == True].tolist()

        self.coordenadas_principal.rename(columns={' RTVC':'RTVC'}, inplace=True)

        for index in index_nulos_principal:
            departamento = self.canalizacion_principal.at[index,'Departamento']
            municipio_nulo_principal = self.canalizacion_principal.at[index,'Municipio']
            index_regionales = self.regionales.index[(self.regionales['Municipio'] == municipio_nulo_principal) & (self.regionales['Departamento'] == departamento)].tolist()[0]
            canal_regional = self.regionales.at[index_regionales, 'Operador']
            self.canalizacion_principal.loc[index,'Operador Regional'] = canal_regional

        self.municipios = self.canalizacion_principal['Municipio'].tolist()

        for municipio in self.municipios:
            indice_excepciones = self.coordenadas_principal.index[self.coordenadas_principal['Municipio'] == municipio].tolist()[0]

            if self.coordenadas_principal['PWR\ndBm'].isnull().tolist()[indice_excepciones] == True:
                indice_borrar = self.canalizacion_principal.index[self.canalizacion_principal['Municipio'] == municipio].tolist()[0]
                self.canalizacion_principal.loc[indice_borrar,'Estación TDT CCNP'] = 'SIN OBLIGACIÓN'
                self.canalizacion_principal.loc[indice_borrar,'RCND'] = pd.NA
                self.canalizacion_principal.loc[indice_borrar,'CRCD'] = pd.NA

        self.canalizacion_adicional[['Estación Regional TV Analógica', 'Operador Regional']] = self.canalizacion_adicional['Estación Regional TV Analógica'].str.split(" - ", expand=True)
        index_nulos_adicional = self.canalizacion_adicional.index[self.canalizacion_adicional['Operador Regional'].isnull() == True].tolist()

        for index in index_nulos_adicional:
            departamento = self.canalizacion_adicional.at[index,'Departamento']
            municipio_nulo_adicional = self.canalizacion_adicional.at[index,'Municipio']
            index_regionales = self.regionales.index[(self.regionales['Municipio'] == municipio_nulo_adicional) & (self.regionales['Departamento'] == departamento)].tolist()[0]
            canal_regional = self.regionales.at[index_regionales, 'Operador']
            self.canalizacion_adicional.loc[index,'Operador Regional'] = canal_regional

        self.iterador_principal = list(SEARCH_PRINCIPALS['Analógico'].keys()) + list(SEARCH_PRINCIPALS['Digital'].keys())
        self.iterador_adicional = list(SEARCH_ADDITIONALS['Analógico'].keys()) + list(SEARCH_ADDITIONALS['Digital'].keys())