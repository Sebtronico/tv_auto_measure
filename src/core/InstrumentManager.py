from RsInstrument import RsInstrument, ResourceError
from typing import Dict

class InstrumentManager(RsInstrument):
    _first_instances: Dict[str, 'InstrumentManager'] = {}

    def __init__(self, ip_address: str):
        """
        Inicializa el instrumento con la configuración personalizada.
        
        Args:
            ip_address: Dirección IP del instrumento
        """
        self.ip_address = ip_address
        resource_options = [f'TCPIP::{ip_address}::INSTR', f'TCPIP::{ip_address}::5025::SOCKET']

        for resource_string in resource_options:
            try:
                if ip_address in self._first_instances:
                    # Si ya existe una instancia, reusamos la sesión ya creada con el parámetro direct_session.
                    existing_instr = self._first_instances[ip_address]
                    super().__init__(resource_name=resource_string, direct_session=existing_instr)
                else:
                    # Si es la primera conexión, creamos una nueva instancia y la guardamos en _first_instances.
                    super().__init__(resource_name=resource_string)
                    self._first_instances[ip_address] = self
                break  # Sale del bucle si la conexión fue exitosa
            except ResourceError as e:
                # print(f'Error al conectar con {resource_string}.')
                pass

        # Si la instancia no está en _first_instances, significa que la conexión falló.
        if ip_address not in self._first_instances:
            print('No se pudo establecer conexión con el instrumento. Verifique la IP o su disponibilidad.')
            return

        # Configurar atributos personalizados
        self.instrument_status_checking = True  # Error check after each command
        self.visa_timeout = 30e3 
        self.opc_timeout = 60e3  # Timeout for opc-synchronised operations
        self.data_chunk_size = 100  # Definición del tamaño del buffer
        self.opc_query_after_write = True

        # Verificar IDN antes de escribir comandos
        try:
            if self.instrument_model_name == 'ETL':
                self.write('SYST:DISP:UPD ON')
        except Exception as e:
            print(f'Error al verificar IDN o configurar el instrumento: {e}')