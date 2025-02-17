from InstrumentManager import InstrumentManager
from src.utils.constants import *
import time
import statistics

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
    def get_dat_file(self, filename: str):
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
        self.get_dat_file(filename)

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
    

    # Función para que el programa espere a que se lean todas las variables
    # dentro de cada modo de medición.
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
        self.get_dat_file(filename)

        atv_dict = {
            'frequency_video': TV_TABLE[channel] - 1.75,
            'frequency_audio': TV_TABLE[channel] + 2.75,
            'power_video': round(self.query_float_with_opc('CALC:MARK2:Y?'), 2),
            'power_audio': round(self.query_float_with_opc('CALC:MARK3:Y?'), 2)
        }

        return atv_dict


if __name__ == '__main__':
    etl = EtlManager('172.23.82.39')
    # etl.write_str('CONF:DTV:MEAS CCDF')  # Selecciona la ventana Modulation errors
    etl.write_str('CONF:DTV:MEAS EPATtern')
    # print(len(etl.query_bin_or_ascii_float_list('TRAC? TRACE1')))
    # etl.write_str('DISP:LIST:STATE OFF')
    print(etl.query('CALC:DTV:RES:BFIL? EPPV'))