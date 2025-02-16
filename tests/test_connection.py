from RsInstrument import *

try:
    instr = RsInstrument('TCPIP::172.23.82.39::INSTR')
    instr.write('SYST:DISP:UPD ON')
    print(instr.resource_name)
except ResourceError:
    print('Error')

instr.close()