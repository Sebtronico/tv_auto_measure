from RsInstrument import *
from typing import Dict

class InstrumentManager(RsInstrument):
    # Diccionario para almacenar la primera instancia de cada IP
    _first_instances: Dict[str, 'InstrumentManager'] = {}
    
    def __init__(self, ip_address: str):
        """
        Inicializa el instrumento con la configuración personalizada.
        
        Args:
            ip_address: Dirección IP del instrumento
        """
        try:
            resource_string = f'TCPIP::{ip_address}::INSTR'
            if ip_address in self._first_instances:
                # Si ya existe una instancia, reusamos la sesión ya creada con el parámetro direct_session
                existing_instr = self._first_instances[ip_address]
                super().__init__(resource_name = resource_string, direct_session = existing_instr)
            else:
                # Si es la primera conexión
                super().__init__(resource_name = resource_string)
                self._first_instances[ip_address] = self
            
            # Configurar atributos personalizados
            self.instrument_status_checking = True  # Error check after each command
            self.visa_timeout = 10e3 
            self.opc_timeout = 60e3  # Timeout for opc-synchronised operations
            self.data_chunk_size = 100 # Definición del tamaño del buffer
            self.opc_query_after_write = True
            
            if self.idn_string.split(sep=',')[1] == 'ETL-3':
                self.write('SYST:DISP:UPD ON')
        except ResourceError as e:
            print(e.args[0])
            print('Error al conectar con el instrumento. \nVerifique su disponibilidad o que la IP sea correcta.')