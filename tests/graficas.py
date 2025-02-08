from RsInstrument import * 
import time
import numpy as np
import matplotlib.pyplot as plt

try:
    instr = RsInstrument('TCPIP::172.23.82.51::INSTR')
    instr.write('SYST:DISP:UPD ON')
    print(instr.resource_name)
except ResourceError:
    print('Error')

instr.clear_status()
instr.reset()

# Configuraciones generales para todas las bandas
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
instr.write_str_with_opc('SWE:POIN 1000') # Configuración del resolution bandwidth

instr.write_str_with_opc('INIT:CONT OFF') # Encendido del modo de barrido continuo
instr.write_str_with_opc('INIT') # Inicio del barrido

# Tiempo total de adquisición (3 minutos)
end_time = time.time() + 30  # 180 segundos (3 minutos)
    
traces = []

while time.time() < end_time:
    instr.write_str_with_opc('INIT;*WAI')
    waveform = instr.query_bin_or_ascii_float_list_with_opc('TRAC? TRACE1')
    traces.append(waveform)

trace_matrix = np.array(traces)

average = np.mean(trace_matrix, axis=0)
maxhold = np.max(trace_matrix, axis=0)
minhold = np.min(trace_matrix, axis=0)

freq = np.linspace(894, 960, 1000)

# Graficar los resultados
plt.figure(figsize=(16, 10))

# Gráfica del promedio
plt.subplot(3, 1, 1)
plt.plot(freq, average, label='Average', color='blue')
plt.title('Promedio')
plt.xlabel('Frecuencia [MHz]')
plt.ylabel('Potencia [dBm]')
plt.legend()
plt.grid()
plt.xlim(min(freq), max(freq))
plt.ylim(min(minhold), max(maxhold))

# Gráfica del máximo
plt.subplot(3, 1, 2)
plt.plot(freq, maxhold, label='Max hold', color='green')
plt.title('Máximo')
plt.xlabel('Frecuencia [MHz]')
plt.ylabel('Potencia [dBm]')
plt.legend()
plt.grid()
plt.xlim(min(freq), max(freq))
plt.ylim(min(minhold), max(maxhold))

# Gráfica del mínimo
plt.subplot(3, 1, 3)
plt.plot(freq, minhold, label='Min hold', color='red')
plt.title('Mínimo')
plt.xlabel('Frecuencia [MHz]')
plt.ylabel('Potencia [dBm]')
plt.legend()
plt.grid()
plt.xlim(min(freq), max(freq))
plt.ylim(min(minhold), max(maxhold))

# Ajustar el layout para que no se solapen las gráficas
plt.tight_layout()

# Mostrar las gráficas
plt.savefig("Ejemplo.png")

# Crear el espectrograma
plt.figure(figsize=(10, 6))
extent = [894, 960, 0, len(traces)]  # Rango de frecuencia en X, número de muestras en Y
plt.imshow(trace_matrix, aspect='auto', extent=extent, origin='lower', cmap='inferno')

# Etiquetas y título
plt.xlabel("Frecuencia (MHz)")
plt.ylabel("Tiempo (muestras)")
plt.title("Espectrograma")
plt.colorbar(label="Potencia (dBm)")

plt.savefig("Espectrograma.png")

instr.close()