from InstrumentManager import InstrumentManager
from src.utils.constants import *
import time
import statistics
import csv
import numpy as np
import os
import matplotlib.pyplot as plt

class EtlManager(InstrumentManager):
    def __init__(self, ip_address: str):
        super().__init__(ip_address)  # Llama al constructor de InstrumentManager

    
    # Generación de captura de pantalla *.png y envío al pc
    def get_screenshot(self, filename: str):
        img_path_instr = 'C:\\R_S\\instr\\user\\screenshot.png' # Ruta y nombre del screenshot en el instrumento
        img_path_pc = f'{filename}.png' # Ruta y nombre del archivo exportado

        self.write_str("HCOP:DEV:LANG PNG") # Definición del formato de la imagen
        self.write_str(f"MMEM:NAME '{img_path_instr}'") # Creación de la imagen en la memoria
        self.write_str("HCOP:IMM") # Captura de la pantalla
        # self.query_opc()

        self.read_file_from_instrument_to_pc(img_path_instr, img_path_pc) # Transferencia del archivo al PC
        self.write_str(f"MMEMory:DELete '{img_path_instr}'")  # Se elimina el archivo de la memoria del instrumento 

    
    # Función de generación del archivo de datos *.dat
    def get_data_file(self, filename: str):
        file_path_instr = 'C:\\R_S\\instr\\user\\datfile.dat' # Ruta y nombre del archivo en el instrumento
        file_path_pc = f'{filename}.dat' # Ruta y nombre del archivo exportado

        self.write_str(f"MMEM:STOR:TRAC 1,'{file_path_instr}'") # Captura del archivo
        # self.query_opc()

        self.read_file_from_instrument_to_pc(file_path_instr, file_path_pc) # Transferencia del archivo al PC
        self.write_str(f"MMEMory:DELete '{file_path_instr}'")  # Se elimina el archivo de la memoria del instrumento


    # Función para leer sobrecarga
    def read_overload(self, seconds: int):
        t = time.time()
        while time.time() - t < seconds:
            overload = self.query_bool_with_opc('STAT:QUES:POW:COND?')
            if overload:
                break
            else:
                continue
        
        return overload
    

    # Función para tratar con la saturación, en caso de existir
    def handle_overload(self):

        overload = self.read_overload(10)

        if overload:
            self.preselector = False
            self.write_bool('INP:PRES:STAT', self.preselector)
            self.write_str('INIT;*WAI')

            overload = self.read_overload(10)

            if overload:
                self.reference_level = 100
                self.write_str(f'DISP:TRAC:Y:RLEV {self.reference_level}')
                self.write_str(f'DISP:TRAC:Y {self.reference_level} dB') # Configura el range log.
                self.write_str('INIT;*WAI')

                overload = self.read_overload(10)

                if overload:
                    while overload and self.attenuation < 20:
                        self.attenuation += 5
                        self.write_str(f'INP:ATT {self.attenuation} dB')
                        self.write_str('INIT;*WAI')

                        overload = self.read_overload(10)

    
    # Función para medición de potencia de canal en tv digital
    def dtv_power_measurement(self, impedance: int, transducers: list, channel: int, path: str):

        # Configuración de parámetros
        self.attenuation = 0
        self.preselector = False
        self.reference_level = 82 if impedance == 50 else 83.75

        self.write_str('INST SAN') # Configura el instrumento al modo "Spectrum Analyzer"
        self.write_str(f'INP:IMP {impedance}') # Selecciona la entrada según la entrada de la función.
        for transducer in transducers:
            self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
            self.write_str('CORR:TRAN ON') # Activa el transductor seleccionado
        self.write_bool(f'INP:PRES:STAT', {self.preselector}) # Enciende el preselector.
        self.write_str('INP:GAIN:STAT OFF') # Apaga el preamplificador.
        self.write_str('DISP:TRAC1:MODE AVER') # Selecciona el modo de traza "average" para la traza 1.
        self.write_str('DET RMS') # Selecciona el detector "RMS"
        self.write_str('CALC:MARK:FUNC:POW:PRES NONE') # Configuración de sin estándar para la medición de potencia en modo ACP 
        self.write_str('CALC:MARK:FUNC:POW:SEL ACP') # Activa la medición de potencia absoluta.
        self.write_str('POW:ACH:ACP 0') # Configura el número de canales adyacentes a 0.
        self.write_str('POW:ACH:BWID 5.830 MHz') # Configuración del ancho de banda a 5.830 MHz.
        self.write_str('CALC:MARK:AOFF') # Desactiva todos los marcadores.
        self.write_str('FREQ:SPAN 8 MHz') # Configuración del Span
        self.write_str(f'DISP:TRAC:Y:RLEV {self.reference_level}') # Configura el expected level.
        self.write_str(f'DISP:TRAC:Y {self.reference_level} dB') # Configura el range log.
        self.write_str('BAND:RES 30 kHz') # Configuración de resolution bandwidth
        self.write_str('BAND:VID 300 kHz') # Configuración de video bandwidth
        self.write_str('SWE:TIME 500 ms') # Configuración de sweeptime
        self.write_str(f'INIT:CONT ON') # Activa la medición contínua
        self.write_str(f'INP:ATT {self.attenuation} dB') # Configuración de la atenuación

        # Configuración de la frecuencia central
        self.write_str(f'FREQ:CENT {TV_TABLE[channel]} MHz') 

        # Lectura de saturación y configuración, en caso de existir
        self.handle_overload()

        self.write_str(f'INIT:CONT OFF') # Desactiva la medición contínua
        self.write_str(f'SWE:COUN 10') # Configuración del conteo de barridos a 10.
        self.write_str(f'INIT;*WAI') # Inicia la medición y aguarda hasta que se complete el número de barridos seleccionado.
        channel_power = self.query_float_with_opc('CALC:MARK:FUNC:POW:RES? ACP') # Lectura del nivel de potencia.
        channel_power = round(channel_power,2) # Redondeo a dos cifras decimales

        filename = f'{path}/{TV_TABLE[channel]}'
        self.get_screenshot(filename) # Toma de la captura de pantalla
        
        self.write_str('FREQ:SPAN 5.830 MHz') # Configuración del Span
        self.write_str(f'INIT;*WAI') # Inicia la medición y aguarda hasta que se complete el número de barridos seleccionado.
        self.get_data_file(filename)

        # Cálculo de tipo de desviación estándar y tipo de canal
        sigma = round((statistics.pstdev(self.query_bin_or_ascii_float_list('TRAC? TRACE1'))),2)
        if sigma >= 0 and sigma <= 1:
            channel_type = f'Gauss (σ = {sigma} dB)'
        elif sigma > 1 and sigma <= 3:
            channel_type = f'Rice (σ = {sigma} dB)'
        elif sigma > 3:
            channel_type = f'Rayleigh (σ = {sigma} dB)'
        
        self.write_str(f'INIT:CONT ON') # Activa la medición contínua

        return {'channel_power': channel_power, 'channel_type': channel_type}
    

    # Función para que el programa espere a que se lean todas las variables dentro de cada modo de medición.
    def wait_for_variables(self, mode: str, seconds: int):
        
        parameters = MODE_PARAMETERS.get(mode, [])  # Evita KeyError si mode es inválido
        if not parameters:
            raise ValueError(f"Modo de medición '{mode}' no válido")

        header = 'CALC:DTV:RES:APGD?' if mode in {'APHase', 'AGRoup'} else 'CALC:DTV:RES?'
        dict_out = {}

        start_time = time.time()

        while time.time() - start_time <= seconds:
            if mode == 'OVER':
                if self.query_str_with_opc(f'{header} PBERldpc') != '10':
                    continue

            for param in parameters:
                try:
                    value = self.query_str_with_opc(f'{header} {param}')
                    if value == '---':  # Si sigue sin valor, espera e intenta en el próximo ciclo
                        continue
                    dict_out[param] = float(value) if value.replace('.', '', 1).isdigit() else value
                except Exception as e:
                    print(f"Error al consultar {param}: {e}")
                    time.sleep(1)

            if all(value != '---' for value in dict_out.values()):
                break

        # Ajuste especial para 'PERatio'
        if dict_out.get('PERatio') == 'n/a (HEM)':
            dict_out['PERatio'] = 0

        return dict_out
    

    # Función para medida del modo spectrum en tv digital
    def dtv_spectrum_measurement(self, channel: int, path: str):
        self.write_str('INST CATV')  # Entrar al modo TV / Radio Analyzer / Receiver.
        self.write_str('CONF:DTV:MEAS DSP')  # Selecciona la ventana Spectrum
        self.write_str('SYST:POS:GPS:DEV PPS2')  # Para que muestre las coordenadas en las imágenes
        self.write_str('DISP:MEAS:OVER:GPS:STAT ON')  # Para que muestre las coordenadas en las imágenes
        self.write_str('DTV:BAND:CHAN B6MHz') # Configura el ancho de banda del canal de TV a 6 MHz.
        self.write_str('DDEM:ISSY TOL') # Configura el modo 'Tolerant' para ISSY processing.
        self.write_str('UNIT:POW DBUV') # Configura la unidad de medida por defecto a dBuV.
        self.write_str('FREQ:SPAN 10 MHz') # Configura el Span.
        self.write_str(f'FREQ:RF {TV_TABLE[channel]} MHz') # Configuración de la frecuencia central.

        self.write_str('CONF:DTV:MEAS:OOB OFF') # Desactiva todos los modos de medida dentro del modo Spectrum.
        self.write_str('CONF:DTV:MEAS:SATT ON') # Activa el modo de función "Shoulders".
        self.write_str(f'INIT:CONT OFF') # Desactiva la medición contínua
        self.write_str(f'SWE:COUN 10') # Configuración del conteo de barridos a 10.
        self.write_str(f'INIT;*WAI') # Inicia la medición y aguarda hasta que se complete el número de barridos seleccionado.

        dict_dsp = self.wait_for_variables('DSP', 30)

        filename = f'{path}/{TV_TABLE[channel]}_010'
        self.get_screenshot(filename)

        return dict_dsp
    

    # Función para medida del modo overview en tv digital
    def dtv_overview_measurement(self, channel: int, path: str):
        self.write_str('CONF:DTV:MEAS OVER')
        self.write_str('DISP:ZOOM:OVER BERLdpc')  # Hacer zoom a la variable BER bef. LDPC.

        # Adquisición de parámetros
        dict_over = self.wait_for_variables('OVER', 30) # Se obtienen todos los valores del modo overview para el txcheck
        BERLdpc = float(dict_over['BERLdpc']) if dict_over['BERLdpc'] != '---' else 'ND'
        
        # Toma de captura de pantalla
        filename = f'{path}/{TV_TABLE[channel]}_003'
        self.get_screenshot(filename)

        # Medición en la ventana L1 pre signalling.
        self.write_str('CONF:DTV:MEAS L1PRe')
        time.sleep(1)
        GINTerval = GINTERVAL_TABLE[self.query_with_opc('CALC:DTV:RES? GINTerval')] # Obtención de la variable intervalo de guardas
        try:
            PLPCodeRate = FEC_TABLE[self.query_with_opc('CALC:DTV:RES:L1Post? DPLP').split(sep=',')[4]] # Obtención de la variable FEC
        except IndexError:
            PLPCodeRate = '---'

        dict_over.update({'BERLdpc': BERLdpc, 'GINTerval': GINTerval, 'PLPCodeRate': PLPCodeRate})

        filename = f'{path}/{TV_TABLE[channel]}_004'
        self.get_screenshot(filename) # Toma de captura de pantalla
        
        # Medición en la ventana L1 post signalling 1.
        self.write_str('CONF:DTV:MEAS L1P1')
        time.sleep(1)
        filename = f'{path}/{TV_TABLE[channel]}_005'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.
        
        # Medición en la ventana L1 post signalling 2.
        self.write_str('CONF:DTV:MEAS L1P2')
        time.sleep(1)
        filename = f'{path}/{TV_TABLE[channel]}_006'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        # Medición en la ventana L1 post signalling 3.
        self.write_str('CONF:DTV:MEAS L1P3')
        time.sleep(1)
        filename = f'{path}/{TV_TABLE[channel]}_011'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.
        
        return dict_over
    

    # Función para la medida del modo modulation analysis en tv digital
    def dtv_modulation_analysis_measurement(self, channel: int, path: str):
        # Medición en la ventana 'Modulation errors'
        self.write_str('CONF:DTV:MEAS MERR')  # Selecciona la ventana Modulation errors
        self.write_str('DISP:ZOOM:MERR MRPLp') # Zoom a la ariable MER (PLP, RMS)
        
        dict_merr = self.wait_for_variables('MERR', 30)
        MRPLp = round(float(dict_merr['MRPLp']), 1) if dict_merr['MRPLp'] != '---' else 'ND'

        filename = f'{path}/{TV_TABLE[channel]}_002'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        # Medición en la ventana de constelación.
        self.write_str('CONF:DTV:MEAS CONS')
        time.sleep(5)
        try:
            cons = MODULATION_TABLE[self.query_with_opc('CALC:DTV:RES:L1Post? DPLP').split(sep=',')[1]] # Obtención de la variable modulación.
        except IndexError:
            cons = '---'

        FFTMode = FFT_MODE_TABLE[self.query_with_opc('CALC:DTV:RES? FFTMode')] # Obtención de la variable FFT.

        dict_merr.update({'MRPLp': MRPLp, 'cons': cons, 'FFTMode': FFTMode})
        
        filename = f'{path}/{TV_TABLE[channel]}_001'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        # Medición en la ventana MER vs Carrier.
        self.write_str('CONF:DTV:MEAS MERFrequency')
        
        t = time.time()
        while True:
            result = etl.query_str_with_opc('CALC:DTV:RES? MERFrms')  
            if result != '---':  
                time.sleep(0.25)
                break  # Sale del bucle si la condición 1 se cumple

            if time.time() - t >= 10:  
                break  # Sale del bucle si han pasado 10 segundos

        filename = f'{path}/{TV_TABLE[channel]}_009'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        self.write_str('CONF:DTV:MEAS CCDF')  # Selecciona la ventana Modulation errors
        time.sleep(2)
        
        filename = f'{path}/{TV_TABLE[channel]}_012'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        return dict_merr
    

    # Función para la medida del modo channel analysis en tv digital
    def dtv_channel_analysis_measurement(self, channel: int, path: str, MRPLp, BERLdpc):
        
        # Medición la ventana Echo Pattern.
        self.write_str('CONF:DTV:MEAS EPATtern')
        self.write_str('DISP:LIST:STATE OFF') # Desactiva la vista de la lista.

        t = time.time()
        while True:
            result = etl.query_str_with_opc('CALC:DTV:RES? MERFrms')  
            if result != '---':
                time.sleep(0.25)
                break  # Sale del bucle si la condición 1 se cumple

            if time.time() - t >= 10:  
                break  # Sale del bucle si han pasado 10 segundos

        PPATtern = self.query_with_opc('CALC:DTV:RES:L1PR? PPATtern') # Obtención de la variable patrón de pilotos.

        filename = f'{path}/{TV_TABLE[channel]}_007'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.
        
        self.write_str('DISP:LIST:STATE ON') # Aciva la vista de la lista.
        filename = f'{path}/{TV_TABLE[channel]}_008'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        time_to_wait = 30 if MRPLp != 'ND' and BERLdpc != 'ND' else 5

        self.write_str('CONF:DTV:MEAS APHase') # Selecciona la ventana Amplitude and Phase.
        dict_aph = self.wait_for_variables('APHase', time_to_wait)
        filename = f'{path}/{TV_TABLE[channel]}_013'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        self.write_str('CONF:DTV:MEAS AGRoup') # Selecciona la ventana Amplitude and Phase.
        dict_agr = self.wait_for_variables('AGRoup', time_to_wait)
        filename = f'{path}/{TV_TABLE[channel]}_014'
        self.get_screenshot(filename) # Toma de captura de pantalla # Toma de captura de pantalla.

        dict_apg = dict_aph | dict_agr

        dict_apg.update({'PPATtern': PPATtern})

        return dict_apg
    

    # Función para medición de televisión analógica
    def atv_measurement(self, impedance: int, transducers: list, channel: int, path: str):
        
        self.write_str('INST SAN') # Switches the instrument to "Spectrum Analyzer" mode.
        for transducer in transducers:
            self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
            self.write_str('CORR:TRAN ON') # Activa el transductor seleccionado
        self.write_str(f'INP:PRES:STAT {self.preselector}') # Enciende el preselector.
        self.write_str('INP:GAIN:STAT OFF') # Apaga el preamplificador.
        self.write_str(f'INP:IMP {impedance}') # Selecciona la entrada según la entrada de la función.
        self.write_str(f'INIT:CONT OFF') # Desactiva la medición contínua
        self.write_str(f'SWE:COUN 10') # Configuración del conteo de barridos a 10.
        self.write_str('DISP:TRAC1:MODE AVER') # Select the average mode for trace 1
        self.write_str('DET RMS') # Select the RMS detector
        self.write_str('CALC:MARK:AOFF') # Switches off all markers.
        self.write_str('FREQ:SPAN 6.5 MHz') # Setting the span
        self.write_str('BAND:RES 10 kHz') # Setting the resolution bandwidth
        self.write_str('BAND:VID 30 kHz') # Setting the video bandwidth
        self.write_str('SWE:TIME 500 ms') # Setting the sweeptime
        self.write_str('INP:ATT 0 dB') # Setting the attenuation

        self.write_str(f'FREQ:CENT {TV_TABLE[channel]} MHz')
        self.write_str(f'CALC:MARK1 ON')
        self.write_str(f'CALC:MARK1:X {TV_TABLE[channel] - 3} MHz')
        self.write_str(f'CALC:MARK2 ON')
        self.write_str(f'CALC:MARK2:X {TV_TABLE[channel] - 1.75} MHz')
        self.write_str(f'CALC:MARK3 ON')
        self.write_str(f'CALC:MARK3:X {TV_TABLE[channel] + 2.75} MHz')
        self.write_str(f'CALC:MARK4 ON')
        self.write_str(f'CALC:MARK4:X {TV_TABLE[channel] + 3} MHz')
        self.write_str(f'INIT;*WAI')

        filename = f'{path}/CH_{channel}'
        self.get_screenshot(filename)
        self.get_data_file(filename)

        atv_dict = {
            'frequency_video': TV_TABLE[channel] - 1.75,
            'frequency_audio': TV_TABLE[channel] + 2.75,
            'power_video': round(self.query_float_with_opc('CALC:MARK2:Y?'), 2),
            'power_audio': round(self.query_float_with_opc('CALC:MARK3:Y?'), 2)
        }

        return atv_dict


    # Función para convertir coordenadas de formato decimal a grados, minutos y segundos
    @staticmethod
    def decimal_coords_to_dms(latitude: float, longitude: float):
        # Conversión de latitud
        lat_deg = int(abs(latitude))
        lat_min_dec = (abs(latitude) - lat_deg) * 60
        lat_min = int(lat_min_dec)
        lat_seg = (lat_min_dec - lat_min) * 60
        lat_direction = "N" if latitude >= 0 else "S"
        lat_dms = f"{lat_deg}° {lat_min}' {lat_seg:.2f}\" {lat_direction}"

        # Conversión de longitud
        lon_deg = int(abs(longitude))
        lon_min_dec = (abs(longitude) - lon_deg) * 60
        lon_min = int(lon_min_dec)
        lon_seg = (lon_min_dec - lon_min) * 60
        lon_direction = "E" if longitude >= 0 else "W"
        lon_dms = f"{lon_deg}° {lon_min}' {lon_seg:.2f}\" {lon_direction}"

        return lat_dms, lon_dms
    

    # Función para obtener las coordenadas del ETL
    def get_coordinates(self):
        self.write_str('INST CATV')  # Entrar al modo TV / Radio Analyzer / Receiver.
        self.write_str('CONF:DTV:MEAS OVER')  # Selecciona la ventana Spectrum
        self.write_str('SYST:POS:GPS:DEV PPS2')  # Para que muestre las coordenadas en las imágenes
        self.write_str('DISP:MEAS:OVER:GPS:STAT ON')  # Para que muestre las coordenadas en las imágenes

        while True:
            try:
                latitude  = self.query_float_with_opc('SYST:POS:LAT?')
                longitude = self.query_float_with_opc('SYST:POS:LONG?')
                break
            except ValueError:
                continue

        return self.decimal_coords_to_dms(latitude, longitude)


    # Función para añadir hora y coordenadas al .dat
    def add_to_dat_file(self, filename: str, latitude: str, longitude: str):
        # Leer el contenido del archivo en codificación ANSI (latin-1)
        with open(filename, "r", encoding="latin-1") as f:
            lines = f.readlines()

        hour = self.query_bin_or_ascii_int_list_with_opc('SYSTem:TIME?')

        # Definir las líneas que quieres insertar
        new_lines = [
            f"Hour; {hour[0]}:{hour[1]}:{hour[2]};\n"
            f"Serial; {self.instrument_serial_number};\n"
            f"Latitude;{latitude};\n",
            f"Longitude;{longitude};\n"
        ]

        # Buscar la línea con que contiene "Date" y agregar después las nuevas líneas
        for i, linea in enumerate(lines):
            if "DATE" in linea.upper():
                lines[i + 1:i + 1] = new_lines  # Insertar líneas después

        # Escribir de nuevo el archivo en codificación ANSI (latin-1)
        with open(filename, "w", encoding="utf-8") as f:
            f.writelines(lines)


    # Función para configuración inicial del banco de mediciones
    def measurement_bank_setup(self, impedance: int, transducers: list, band: str):

        # Configuraciones generales para todas las bandas
        self.write_str('INST SAN') # Configura el instrumento al modo "Spectrum Analyzer"
        self.write_str(f'DET RMS') # Selecciona el detector "RMS"
        self.write_str(f'INP:ATT 0 dB')
        self.write_str(f'INP:GAIN:STAT OFF')

        # Configuración por cada banda
        self.write_str(f'INP:IMP {impedance}') # Selecciona la entrada según la entrada de la función.

        # Activa los transductores seleccionados si la unidad es dBuV/m
        if BANDS_ETL[band][6] == 'DBUVm':
            for transducer in transducers:
                self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
                self.write_str('CORR:TRAN ON') # Activa el transductor seleccionado
        # En caso contrario, los apaga todos
        else:
            for transducer in transducers:
                self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
                self.write_str('CORR:TRAN OFF') # Apaga el transductor seleccionado


        self.write_str(f'UNIT:POW {BANDS_ETL[band][6]}') # Configuración de la unidad
        
        # Configuración del instrumento según la banda
        self.write_str(f'FREQ:STAR {BANDS_ETL[band][0]} MHz') # Configuración de la frecuencia inicial
        self.write_str(f'FREQ:STOP {BANDS_ETL[band][1]} MHz') # Configuración de la frecuencia final
        self.write_str(f'BAND:VID {BANDS_ETL[band][2]} kHz') # Configuración del video bandwidth
        self.write_str(f'BAND:RES {BANDS_ETL[band][3]} kHz') # Configuración del resolution bandwidth
        
        # Definición del reference level, según el puerto seleccionado
        reference_level = 82 if impedance == 50 else 83.75

        # Ajuste del nivel de referencia, según el puerto seleccionado y la unidad de medida.
        if BANDS_ETL[band][6] in ['DBUV', 'DBUVm']:
            self.write_str(f'DISP:TRAC:Y:RLEV {reference_level}') # Configuración del nivel de referencia
        else:
            self.write_str(f'DISP:TRAC:Y:RLEV {BANDS_ETL[band][4]}') # Configuración del nivel de referencia


    # Función para el banco de mediciones en el modo de obtener solo una traza con el .dat
    def measurement_bank_one_trace(self, impedance: int, transducers: list, band: str, path: str, latitude: str, longitude: str):
        # Configuraciones generales para todas las bandas
        self.measurement_bank_setup(impedance, transducers, band)

        # Configuración de la medición
        self.write_str(f'DISP:TRAC1:MODE {BANDS_ETL[band][5]}') # Configuración del modo de traza
        if BANDS_ETL[band][5] == 'AVERage':
            self.write_str(f'INIT:CONT OFF') # Apagado del modo de barrido continuo
            self.write_str(f'SWE:COUN 10') # Configuración del número de trazas
            self.write_str(f'INIT;*WAI') # Inicio del barrido y espera de que se complete el número de trazas
        elif BANDS_ETL[band][5] == 'MAXHold':
            wait = float(self.query('SWE:TIME?')) # Obtención del tiempo de un barrido
            self.write_str(f'INIT:CONT ON') # Encendido del modo de barrido continuo
            self.write_str(f'INIT') # Inicio del barrido
            time.sleep(wait*10) # Espera a que se complete el número de trazas

        filename = f'{path}/{BANDS_ETL[band][0]} - {BANDS_ETL[band][0]}'
        
        self.get_screenshot(filename)
        self.get_data_file(filename)
        self.add_to_dat_file(f'{filename}.dat', latitude, longitude)


    # Función para graficar promedio, máximo y mínimo
    @staticmethod
    def plot_avg_max_min(matrix: np.ndarray, frequency_vector: np.ndarray, filename: str, unit: str):
        average = np.mean(matrix, axis=0)
        maxhold = np.max(matrix, axis=0)
        minhold = np.min(matrix, axis=0)

        ylim_min = min(minhold) * 0.95 if min(minhold) > 0 else min(minhold) * 1.05
        ylim_max = max(maxhold) * 0.95 if max(maxhold) < 0 else max(maxhold) * 1.05

        # Graficar los resultados
        plt.figure(figsize=(12.8, 7.2), dpi=100)

        power_unit = UNITS[unit]
        # Gráfica del promedio
        plt.subplot(3, 1, 1)
        plt.plot(frequency_vector, average, label='Average', color='blue')
        plt.title('Promedio')
        plt.xlabel('Frecuencia [MHz]')
        plt.ylabel(f'Potencia [{power_unit}]')
        plt.legend()
        plt.grid()
        plt.xlim(min(frequency_vector), max(frequency_vector))
        plt.ylim(ylim_min, ylim_max)

        # Gráfica del máximo
        plt.subplot(3, 1, 2)
        plt.plot(frequency_vector, maxhold, label='Max hold', color='green')
        plt.title('Máximo')
        plt.xlabel('Frecuencia [MHz]')
        plt.ylabel(f'Potencia [{power_unit}]')
        plt.legend()
        plt.grid()
        plt.xlim(min(frequency_vector), max(frequency_vector))
        plt.ylim(ylim_min, ylim_max)

        # Gráfica del mínimo
        plt.subplot(3, 1, 3)
        plt.plot(frequency_vector, minhold, label='Min hold', color='red')
        plt.title('Mínimo')
        plt.xlabel('Frecuencia [MHz]')
        plt.ylabel(f'Potencia [{power_unit}]')
        plt.legend()
        plt.grid()
        plt.xlim(min(frequency_vector), max(frequency_vector))
        plt.ylim(ylim_min, ylim_max)

        # Ajustar el layout para que no se solapen las gráficas
        plt.tight_layout()

        # Mostrar las gráficas
        plt.savefig(f"{filename}.png")


    # Función para graficar espectrograma
    @staticmethod
    def plot_spectrogram(matrix: np.ndarray, frequency_vector: np.ndarray, filename: str, unit: str):
        plt.figure(figsize=(12.8, 7.2), dpi=100)

        extent = [min(frequency_vector), max(frequency_vector), 0, len(matrix)]  # Rango de frecuencia en X, número de muestras en Y
        plt.imshow(matrix, aspect='auto', extent=extent, origin='lower', cmap='inferno')

        power_unit = UNITS[unit]
        # Etiquetas y título
        plt.xlabel("Frecuencia [MHz]")
        plt.ylabel("Muestra")
        plt.title("Espectrograma")
        plt.colorbar(label=f"Potencia [{power_unit}]")

        plt.savefig(f"{filename}_E.png")


    # Función para banco de mediciones en el modo de obtener varias trazas con el .csv
    def continuous_measurement_bank(self, impedance: int, transducers: list, band: str, path: str, latitude: str, longitude: str):

        # Configuraciones generales para todas las bandas
        self.measurement_bank_setup(impedance, transducers, band)

        sweep_points = 1000
        self.write_str(f'SWE:POIN {sweep_points}')
        self.write_str(f'SWE:COUN 0') # Configuración del número de trazas
        self.write_str('DISP:TRAC1:MODE WRIT') # Configuración del modo de traza
        self.write_str('INIT:CONT OFF') # Encendido del modo de barrido continuo
        self.write_str('INIT;*WAI')

        # Se crea el archivo .dat, para copiar su estructura inicial
        filename = f'{path}/{BANDS_ETL[band][0]} - {BANDS_ETL[band][1]}'
        self.get_data_file(filename)
        self.add_to_dat_file(f'{filename}.dat', latitude, longitude)

        # Leer el archivo .dat hasta la línea que contiene "Values;" seguido de algún valor
        lineas_filtradas = []
        with open(f'{filename}.dat', "r", encoding="utf-8") as f:
            for linea in f:
                linea_limpia = linea.strip()
                if linea_limpia.lower().startswith("values;") and len(linea_limpia.split(";")) > 1:
                    lineas_filtradas.append(linea_limpia)
                    break
                lineas_filtradas.append(linea_limpia)

        # Escribir las líneas en el archivo .csv
        with open(f'{filename}.csv', "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")  # Usa ";" como separador

            # Escribir las líneas filtradas del .dat en el .csv
            for linea in lineas_filtradas:
                campos = linea.split(";")
                writer.writerow(campos)
            
            # Variables para las gráficas y para csv
            frequency_vector = np.linspace(BANDS_ETL[band][0], BANDS_ETL[band][1], sweep_points)
            traces = []

            # Escribir encabezado
            writer.writerow([f"{i}" for i in frequency_vector.tolist()])  # Ajusta el número de puntos según el instrumento
            
            for _ in range(100):
                self.write_str('INIT;*WAI')
                waveform = self.query_bin_or_ascii_float_list_with_opc('TRAC? TRACE1')
                traces.append(waveform)
                writer.writerow(waveform)

        # Se elimina el archivo .dat
        os.remove(f'{filename}.dat')

        # Generación de gráficas
        matrix_traces = np.array(traces)
        self.plot_avg_max_min(matrix_traces, frequency_vector, filename, BANDS_ETL[band][6])
        self.plot_spectrogram(matrix_traces, frequency_vector, filename, BANDS_ETL[band][6])



class FPHManager(InstrumentManager):
    def __init__(self, ip_address: str):
        super().__init__(ip_address)  # Llama al constructor de InstrumentManager


    # Generación de captura de pantalla *.png y envío al pc
    def get_screenshot(self, filename: str):
        img_path_instr = '\\Public\Screen Shots\\screenshot.png' # Ruta y nombre del screenshot en el instrumento
        img_path_pc = f'{filename}.png' # Ruta y nombre del archivo exportado

        self.write_str("HCOP:DEV:LANG PNG") # Definición del formato de la imagen
        self.write_str(f"MMEM:NAME '{img_path_instr}'") # Creación de la imagen en la memoria
        self.write_str("HCOP:IMM") # Captura de la pantalla

        self.read_file_from_instrument_to_pc(img_path_instr, img_path_pc) # Transferencia del archivo al PC
        self.write_str(f"MMEMory:DELete '{img_path_instr}'")  # Se elimina el archivo de la memoria del instrumento 

    
    # Función de generación del archivo de datos *.dat
    def get_data_file(self, filename: str):
        file_path_instr_set = '\\Public\\Datasets\\datafile.set' # Ruta y nombre del archivo .set en el instrumento
        file_path_instr_csv = '\\Public\\Datasets\\datafile.csv' # Ruta y nombre del archivo .csv en el instrumento
        file_path_pc_set = f'{filename}.set' # Ruta y nombre del archivo exportado
        file_path_pc_csv = f'{filename}.csv' # Ruta y nombre del archivo exportado

        self.write_str(f"MMEM:STOR:STAT 1,'{file_path_instr_set}'")
        self.write_str(f"MMEM:STOR:CSV:STATe 1,'{file_path_instr_csv}'")

        self.read_file_from_instrument_to_pc(file_path_instr_set, file_path_pc_set) # Transferencia del archivo al PC
        self.read_file_from_instrument_to_pc(file_path_instr_set, file_path_pc_csv) # Transferencia del archivo al PC
        
        self.write_str(f"MMEM:DEL '{file_path_instr_set}'")  # Se elimina el archivo de la memoria del instrumento
        self.write_str(f"MMEM:DEL '{file_path_instr_csv}'")  # Se elimina el archivo de la memoria del instrumento

    
    # Función para configuración inicial del banco de mediciones
    def measurement_bank_setup(self, impedance: int, transducers: list, band: str):

        # Configuraciones generales para todas las bandas
        self.write_str(f'INST SAN') # Configura el instrumento al modo "Spectrum Analyzer"
        self.write_str(f'DET RMS') # Selecciona el detector "RMS"
        self.write_str(f'INP:ATT 0 dB') # Atenuación a 0
        self.write_str(f'INP:GAIN:STAT OFF') # Ganancia a 0

        # Configuración por cada banda
        self.write_str(f'INP:IMP {impedance}') # Selecciona la entrada según la entrada de la función.

        # Activa los transductores seleccionados si la unidad es dBuV/m
        if BANDS_FXH[band][6] == 'DUVM':
            for transducer in transducers:
                self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
                self.write_str('CORR:TRAN ON') # Activa el transductor seleccionado
        # En caso contrario, los apaga todos
        else:
            for transducer in transducers:
                self.write_str(f"CORR:TRAN:SEL '{transducer}'") # Selecciona el transductor suministrado por el usuario.
                self.write_str('CORR:TRAN OFF') # Apaga el transductor seleccionado

        self.write_str(f'UNIT:POW {BANDS_FXH[band][6]}') # Configuración de la unidad
        
        # Configuración del instrumento según la banda
        self.write_str(f'FREQ:STAR {BANDS_FXH[band][0]} MHz') # Configuración de la frecuencia inicial
        self.write_str(f'FREQ:STOP {BANDS_FXH[band][1]} MHz') # Configuración de la frecuencia final
        self.write_str(f'BAND:VID {BANDS_FXH[band][2]} kHz') # Configuración del video bandwidth
        self.write_str(f'BAND {BANDS_FXH[band][3]} kHz') # Configuración del resolution bandwidth
        
        # Definición del reference level, según el puerto seleccionado
        reference_level = 82 if impedance == 50 else 83.75

        # Ajuste del nivel de referencia, según el puerto seleccionado y la unidad de medida.
        if BANDS_FXH[band][6] in ['DBUV', 'DUVM']:
            self.write_str(f'DISP:TRAC:Y:RLEV {reference_level}') # Configuración del nivel de referencia
        else:
            self.write_str(f'DISP:TRAC:Y:RLEV {BANDS_FXH[band][4]}') # Configuración del nivel de referencia


    # Función para el banco de mediciones en el modo de obtener solo una traza con el .dat
    def measurement_bank_one_trace(self, impedance: int, transducers: list, band: str, path: str, latitude: str, longitude: str):
        # Configuraciones generales para todas las bandas
        self.measurement_bank_setup(impedance, transducers, band)

        # Configuración de la medición
        self.write_str(f'DISP:TRAC1:MODE {BANDS_ETL[band][5]}') # Configuración del modo de traza
        if BANDS_ETL[band][5] == 'AVERage':
            self.write_str(f'INIT:CONT OFF') # Apagado del modo de barrido continuo
            self.write_str(f'SWE:COUN 10') # Configuración del número de trazas
            self.write_str(f'INIT;*WAI') # Inicio del barrido y espera de que se complete el número de trazas
        elif BANDS_ETL[band][5] == 'MAXHold':
            wait = float(self.query('SWE:TIME?')) # Obtención del tiempo de un barrido
            self.write_str(f'INIT:CONT ON') # Encendido del modo de barrido continuo
            self.write_str(f'INIT') # Inicio del barrido
            time.sleep(wait*10) # Espera a que se complete el número de trazas

        filename = f'{path}/{BANDS_ETL[band][0]} - {BANDS_ETL[band][0]}'
        
        self.get_screenshot(filename)
        self.get_data_file(filename)
        # self.add_to_dat_file(f'{filename}.dat', latitude, longitude)


if __name__ == '__main__':
    # etl = EtlManager('172.23.82.39')
    etl = EtlManager('192.168.1.108')
    etl.reset()
    latitude, longitude = etl.get_coordinates()
    etl.continuous_measurement_bank(75, ['BICOLOG 20300', 'CABLE BICOLOG'], 'Enlace', './tests', latitude, longitude)