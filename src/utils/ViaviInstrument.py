import pyvisa

class ViaviInstrument:
    def __init__(self, ip_address: str):
        self.rm = pyvisa.ResourceManager('@py')
        instrument = self.rm.open_resource(f"TCPIP::{ip_address}::5025::SOCKET")
        instrument.read_termination = '\n'
        instrument.write_termination = '\n'
        instrument.timeout = 5000  # 5000 ms = 5 segundos

        prtm_list = instrument.query(':PRTM:LIST?').split(sep=', ')

        for item in prtm_list:
            module = item.split(sep=': ')[0]
            port = item.split(sep=': ')[1]
            
            if module == 'CA5G-SCPI':
                self.instrument = self.rm.open_resource(f"TCPIP::{ip_address}::{port}::SOCKET")
                break
            else:
                continue

        if hasattr (self, 'instrument'):
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.instrument.timeout = 30000
            self._status_checking = True
            self._opc_timeout = 10000  # 10 segundos por defecto
            self._data_chunk_size = 1000000  # 1MB por defecto
            self._opc_query_after_write = False
        else:
            raise pyvisa.Error('No existe puerto correcto para la comunicación SCPI.')
        
    @property
    def instrument_status_checking(self):
        return self._status_checking
    
    @instrument_status_checking.setter
    def instrument_status_checking(self, value):
        self._status_checking = value
    
    def _check_instrument_errors(self):
        if not self._status_checking:
            return
        
        # Para instrumentos SCPI estándar
        try:
            error = self.instrument.query("SYST:ERR?")
            if not error.startswith("+0"):  # No error
                raise Exception(f"Instrument error: {error}")
        except:
            pass  # Si el comando no existe, ignorar

    @property
    def visa_timeout(self):
        return self.instrument.timeout

    @visa_timeout.setter
    def visa_timeout(self, value_ms):
        self.instrument.timeout = value_ms  # PyVISA usa milisegundos

    def query(self, command: str) -> str:
        return self.instrument.query(command).strip()
    
    def query_with_opc(self, command: str) -> str:
        res = self.query(command).strip()
        self.wait_for_opc()

        return res

    def write(self, command: str):
        self.instrument.write(command)

    @property
    def opc_timeout(self):
        return self._opc_timeout

    @opc_timeout.setter
    def opc_timeout(self, value_ms):
        self._opc_timeout = value_ms

    def wait_for_opc(self):
        """Espera hasta que el instrumento complete la operación"""
        old_timeout = self.instrument.timeout
        self.instrument.timeout = self._opc_timeout
        
        try:
            # Método 1: Usar *OPC?
            result = self.instrument.query("*OPC?")
            return result.strip() == "1"
        except:
            try:
                # Método 2: Usar *OPC y polling de *ESR?
                self.instrument.write("*OPC")
                while True:
                    esr = int(self.instrument.query("*ESR?"))
                    if esr & 1:  # Bit 0 de ESR indica OPC
                        break
                return True
            except:
                return False
        finally:
            self.instrument.timeout = old_timeout

    @property
    def data_chunk_size(self):
        return self._data_chunk_size

    @data_chunk_size.setter
    def data_chunk_size(self, value):
        self._data_chunk_size = value
        # Configurar el buffer de PyVISA
        self.instrument.chunk_size = value

    @property
    def opc_query_after_write(self):
        return self._opc_query_after_write

    @opc_query_after_write.setter
    def opc_query_after_write(self, value):
        self._opc_query_after_write = value

    def write(self, command):
        """Write con OPC opcional"""
        self.instrument.write(command)
        self._check_instrument_errors()
        
        if self._opc_query_after_write:
            self.wait_for_opc()

    def write_with_opc(self, command: str):
        """Write que siempre espera OPC"""
        self.instrument.write(command)
        self._check_instrument_errors()
        self.wait_for_opc()

    @property
    def idn_string(self):
        return self.query('*IDN?')
    
    @property
    def full_instrument_model_name(self):
        return self.idn_string.split(sep=',')[1]

    @property
    def instrument_model_name(self):
        return self.full_instrument_model_name
    
    @property
    def instrument_firmware_version(self):
        return self.idn_string.split(sep=',')[3]
    
    @property
    def instrument_serial_number(self):
        return self.idn_string.split(sep=',')[2]
    
    def close(self):
        self.instrument.close()
        self.rm.close()
    
    def query_int_with_opc(self, command: str) -> int:
        return int(self.query_with_opc(command))
    
    def query_float_with_opc(self, command: str) -> float:
        return float(self.query_with_opc(command))
    
    def query_bin_or_ascii_float_list(self, command: str) -> list[float]:
        res = self.query(command).split(sep=',')

        return [float(item) for item in res]
    
    def query_bin_or_ascii_float_list_with_opc(self, command: str) -> list[float]:
        res = self.query_with_opc(command).split(sep=',')

        return [float(item) for item in res]
    
    def query_bin_or_ascii_int_list_with_opc(self, command: str) -> list[int]:
        res = self.query_with_opc(command).split(sep=',')

        return [int(item) for item in res]
    
    def query_bool_with_opc(self, command: str) -> bool:
        res = self.query_with_opc(command).split(sep=',')

        return [bool(item) for item in res]
    
    def query_str_list_with_opc(self, command: str) -> list[str]:
        return self.query_with_opc(command).split(sep=',')

    def query_str_with_opc(self, command: str) -> str:
        return str(self.query_with_opc(command))
    
    def reset(self):
        self.write('*RST')
        self.write('SPECtrum:CONFigure:RESEt')

    def write_bool(self, command: str, value: bool):
        self.write(f"{command} {'ON' if value else 'OFF'}")

    def write_str(self, command: str):
        self.write(command)

    # def read_file_from_instrument_to_pc(file_path_instr, file_path_pc): Pendiente


    # Delegar atributos automáticamente
    def __getattr__(self, name):
        return getattr(self.instrument, name)


if __name__ == "__main__":
    inst = ViaviInstrument('172.23.83.107')
    print('Equipo', inst.full_instrument_model_name)
    print('Versión', inst.instrument_firmware_version)
    # print('Fecha', date
    # print('Hora', hour
    print('Serial', inst.instrument_serial_number)
    # print('Latitud', latitude
    # print('Longitud', longitude
    print('Frecuencia central', inst.query_int_with_opc('SPEC:FREQ:CENT?') * int(1e6)) 
    print('Frecuencia inicial', inst.query_int_with_opc('SPEC:FREQ:STAR?') * 1e6)
    print('Frecuencia final', inst.query_int_with_opc('SPEC:FREQ:STOP?') * 1e6)
    print('Span', inst.query_int_with_opc('SPEC:FREQ:SPAN?') * 1e6)
    print('Nivel de referencia', round(inst.query_float_with_opc('SPEC:AMP:REF?'), 2))
    print('Offset', inst.query_float_with_opc('SPECtrum:AMPlitude:EXTernal?'))
    print('Resolución de ancho de banda', inst.query_float_with_opc('SPEC:RBW?') * 1e6)
    print('Video de ancho de banda', inst.query_float_with_opc('SPEC:VBW?') * 1e6)
    print('Tiempo de barrido', inst.query_float_with_opc('SPEC:SWEE:TIME?'))
    print('Modo de traza', inst.query_str_with_opc('SPEC:TRAce1:MODE?'))
    print('Detector', inst.query_str_with_opc('SPEC:TRAce:DET?'))
    print('Unidades-x', 'Hz')
    print('Unidades-y', inst.query_str_with_opc('SPEC:AMP:UNIT?'))
    print('Preamplificador', False)