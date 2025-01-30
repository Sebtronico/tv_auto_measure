from RsInstrument import * 
import time

try:
    instr = RsInstrument('TCPIP::172.23.82.51::INSTR')
    instr.write('SYST:DISP:UPD ON')
    print(instr.resource_name)
except ResourceError:
    print('Error')

instr.write_str_with_opc('SYST:POS:GPS:DEV PPS2')
instr.write_str_with_opc('DISP:MEAS:OVER:GPS:STAT ON')
print(f'La longitud es: {instr.query_str_with_opc('SYST:POS:LONG?')}')
print(f'La latitud es: {instr.query_str_with_opc('SYST:POS:LAT?')}')
instr.close()

#self.write_str_with_opc('SYST:POS:GPS:DEV PPS2')  # Para que muestre las coordenadas en las imágenes