from RsInstrument import * 
import time
import datetime
import csv
import numpy as np

try:
    instr = RsInstrument('TCPIP::172.23.82.51::INSTR')
    instr.write('SYST:DISP:UPD ON')
    print(instr.resource_name)
except ResourceError:
    print('Error')

instr.clear_status()
instr.reset()

# Configuraciones generales para todas las bandas
instr.write_str_with_opc('INST CATV')  # Entrar al modo TV / Radio Analyzer / Receiver.
instr.write_str_with_opc('CONF:DTV:MEAS OVER')  # Selecciona la ventana Spectrum
instr.write_str_with_opc('SYST:POS:GPS:DEV PPS2')  # Para que muestre las coordenadas en las imágenes
instr.write_str_with_opc('DISP:MEAS:OVER:GPS:STAT ON')  # Para que muestre las coordenadas en las imágenes

while True:
    try:
        latitude  = instr.query_float_with_opc('SYST:POS:LAT?')
        longitude = instr.query_float_with_opc('SYST:POS:LONG?')
        break
    except ValueError:
        continue

print(f'Latitud:    {latitude}')
print(f'Longitud:   {longitude}')

instr.write_str_with_opc('INST SAN') # Configura el instrumento al modo "Spectrum Analyzer"
instr.write_str_with_opc('DET RMS') # Selecciona el detector "RMS"
instr.write_str_with_opc('INP:ATT 0 dB')
instr.write_str_with_opc('INP:GAIN:STAT OFF')
instr.write_str_with_opc('INP:IMP 75') # Selecciona la entrada según la entrada de la función.
instr.write_str_with_opc('UNIT:POW DBM') # Configuración de la unidad.write_str_with_opc(f'UNIT:POW {self.bands[band][6]}') # Configuración de la unidad
instr.write_str_with_opc('DISP:TRAC1:MODE WRIT') # Configuración del modo de traza

instr.write_str_with_opc('FREQ:STAR 894 MHz') # Configuración de la frecuencia inicial
instr.write_str_with_opc('FREQ:STOP 960 MHz') # Configuración de la frecuencia final
instr.write_str_with_opc('BAND:VID 10 kHz') # Configuración del video bandwidth
instr.write_str_with_opc('BAND:RES 30 kHz') # Configuración del resolution bandwidth

instr.write_str_with_opc('INIT:CONT OFF') # Encendido del modo de barrido continuo
instr.write_str_with_opc('INIT') # Inicio del barrido

# Nombre del archivo CSV
filename = 'mediciones_spectrum.csv'

# Tiempo total de adquisición (3 minutos)
end_time = time.time() + 180  # 180 segundos (3 minutos)
# interval = 

# Crear y abrir el archivo CSV
with open(filename, mode='w', newline='') as file:
    serial = instr.idn_string.split(sep=',')[2].split(sep='/')[0]
    # latitud = "4.6097"  
    # longitud = "-74.0817"  
    writer = csv.writer(file)
    
    # Escribir información adicional
    writer.writerow(['Serial:'] + [serial])
    writer.writerow(['Latitud:'] + [latitude])
    writer.writerow(['Longitud:'] + [longitude])
    
    # Escribir encabezado
    writer.writerow(["Timestamp"] + [f"{i}" for i in np.linspace(894, 960, 501).tolist()])  # Ajusta el número de puntos según el instrumento
    
    while time.time() < end_time:
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-2]
        instr.write_str_with_opc('INIT;*WAI')
        waveform = instr.query_bin_or_ascii_float_list_with_opc('TRAC? TRACE1')
        writer.writerow([timestamp] + waveform)
        print(f'Medición tomada a {timestamp}')

    # print(waveform)

print(f'Datos guardados en {filename}')


instr.close()