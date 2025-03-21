import asyncio
from pysnmp.hlapi.v3arch.asyncio import *
from .RemoteDesktopConnector import *
import pyautogui
import time
from src.utils.constants import *


class SNMPManager:
    def __init__(self, ip_address: str, open_ts_analyzer):
        """Inicializa el SNMP Manager con los parámetros básicos"""
        self.ip_address = ip_address
        self.read_community = 'public'
        self.write_community = 'management'
        self.port = 161
        self.snmpEngine = SnmpEngine()

        # Crea el controlador del escritorio remoto
        self.remote_desktop = RemoteDesktopConnector(self.ip_address)
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        
        # Crea conexión con el controlador SCPI del instrumento
        # from .InstrumentController import EtlManager
        # self.etl = etl
        # self.etl = EtlManager(self.ip_address)
        # self.etl.write_str('INST TSAN')
        self.open_ts_analyzer = open_ts_analyzer
        self.open_ts_analyzer()
        while True:
            try:
                self.get_mac_address()
                self.snmp_set('1.3.6.1.4.1.2566.127.1.1.157.3.1.2.5.1.3.0.0.0.0.0.1.2', 1) # Apaga el puerto 2
                break
            except:
                # self.etl.write_str_with_opc('INST TSAN') # Configura el instrumento al modo "TS Analyzer"
                self.open_ts_analyzer()
                continue


    # Función asíncrona para get
    async def _snmp_get(self, oid: str):
        iterator = await get_cmd(
            self.snmpEngine,
            CommunityData(self.read_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            ObjectType(ObjectIdentity(oid)),
        )
        return self._parse_response(iterator)


    # Función síncrona para get
    def snmp_get(self, oid: str):
        """Versión síncrona de SNMP GET"""
        return self.loop.run_until_complete(self._snmp_get(oid))


    # Función asíncrona para get next
    async def _snmp_get_next(self, oid: str):
        iterator = await next_cmd(
            self.snmpEngine,
            CommunityData(self.read_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            ObjectType(ObjectIdentity(oid)),
        )
        return self._parse_response(iterator)


    # Función síncrona para get next
    def snmp_get_next(self, oid: str):
        """Versión síncrona de SNMP GETNEXT"""
        return self.loop.run_until_complete(self._snmp_get_next(oid))


    # Función asíncrona para bulk
    async def _snmp_bulk(self, oid: str, max_repetitions=10):
        results = {}
        iterator = await bulk_cmd(
            self.snmpEngine,
            CommunityData(self.read_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            0,
            max_repetitions,
            ObjectType(ObjectIdentity(oid)),
        )
        errorIndication, errorStatus, errorIndex, varBindTable = iterator
        
        if errorIndication:
            return f"Error: {errorIndication}"
        elif errorStatus:
            return f"Error: {errorStatus.prettyPrint()}"
        else:
            for varBind in varBindTable:
                results[str(varBind[0])] = str(varBind[1])
            return results


    # Función síncrona para bulk
    def snmp_bulk(self, oid: str, max_repetitions=10):
        """Versión síncrona de SNMP BULK"""
        return self.loop.run_until_complete(self._snmp_bulk(oid, max_repetitions))
    

    # Función asíncrona para bulk-walk
    async def _async_snmp_bulk_walk(self, oid: str):
        """Obtiene todos los valores de un subárbol SNMP sin recorrer más de lo necesario."""
        results = {}
        base_oid = oid.strip(".")  # Normaliza el OID base para comparación

        async for errorIndication, errorStatus, errorIndex, varBinds in bulk_walk_cmd(
            self.snmpEngine,
            CommunityData(self.read_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            0,
            10,
            ObjectType(ObjectIdentity(oid)),
        ):
            if errorIndication:
                return f"Error: {errorIndication}"
            elif errorStatus:
                return f"Error: {errorStatus.prettyPrint()}"
            else:
                for varBind in varBinds:
                    # Verifica que el OID sigue en el mismo subárbol
                    if not str(varBind[0]).startswith(base_oid):
                        return results  # Si ya no pertenece, detenemos la iteración
                    
                    results[str(varBind[0])] = str(varBind[1])

        return results


    # Función síncrona para bulk-walk
    def snmp_bulk_walk(self, oid):
        """Obtiene todos los valores de un subárbol SNMP de forma síncrona."""
        return self.loop.run_until_complete(self._async_snmp_bulk_walk(oid))
    

    # Función para convertir un oid de string a lista
    @staticmethod
    def oid_to_list(oid_str: str):
        """Convierte un OID en string a una lista de enteros para comparación"""
        return list(map(int, oid_str.split('.')))
    

    # Función asíncrona para haber bulk-walk a un rango específico de oids
    async def _async_snmp_bulk_walk_range(self, start_oid, end_oid):
        """ Realiza un recorrido SNMP BULK-WALK dentro de un rango específico de OIDs """
        results = {}

        async for errorIndication, errorStatus, errorIndex, varBinds in bulk_walk_cmd(
            self.snmpEngine,
            CommunityData(self.read_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            0,
            10,
            ObjectType(ObjectIdentity(start_oid))
        ):
            if errorIndication:
                print(f"Error: {errorIndication}")
                break
            elif errorStatus:
                print(f"Error: {errorStatus.prettyPrint()}")
                break
            else:
                for varBind in varBinds:
                    oid = str(varBind[0])
                    value = varBind[1]

                    # Si el OID ya está fuera del rango, detenemos la iteración
                    if self.oid_to_list(str(varBind[0])) > self.oid_to_list(end_oid):
                        return results

                    results[str(varBind[0])] = str(varBind[1])

        return results
    

    # Función síncrona para hacer bulk-walk a un rango específico de oids
    def snmp_bulk_walk_range(self, start_oid, end_oid):
        """ Realiza un recorrido SNMP BULK-WALK dentro de un rango específico de OIDs de forma síncrona"""
        return self.loop.run_until_complete(self._async_snmp_bulk_walk_range(start_oid, end_oid))


    # Función asíncrona para set 
    async def _snmp_set(self, oid, value, data_type=Integer):
        await set_cmd(
            self.snmpEngine,
            CommunityData(self.write_community),
            await UdpTransportTarget.create((self.ip_address, self.port)),
            ContextData(),
            ObjectType(ObjectIdentity(oid), data_type(value)),
        )


    # Función síncrona para set
    def snmp_set(self, oid, value, data_type=Integer):
        """Versión síncrona de SNMP SET"""
        return self.loop.run_until_complete(self._snmp_set(oid, value, data_type))


    # Función común a get y get next para retornar solo el valor solicitado
    def _parse_response(self, iterator):
        """Parsea la respuesta SNMP y maneja errores"""
        errorIndication, errorStatus, errorIndex, varBinds = iterator
        if errorIndication:
            return f"Error: {errorIndication}"
        elif errorStatus:
            return f"Error: {errorStatus.prettyPrint()}"
        else:
            return str(varBinds[0][1])


    # Función para obtener la mac_address
    def get_mac_address(self):
        """Obtiene la dirección MAC de un dispositivo"""
        oid = '1.3.6.1.4.1.2566.127.1.1.157.3.1.2.2.1.2.1'
        response = self.snmp_get(oid)
        self.mac_address = '.'.join(str(int(part, 16)) for part in response.split(':'))


    # Función para enviar comando de refrescar la tabla de TsTree
    def refresh_table(self):
        """Refresca la tabla de TsTree"""
        oid_refresh = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.10.1.1.{self.mac_address}.1'
        self.snmp_set(oid_refresh, 1, data_type=Gauge32)


    # Función para obtener el número de la tabla.
    def get_table_number(self):
        """Obtiene el número de tabla de TsTree"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.10.1.3.{self.mac_address}.1'
        self.table_number = self.snmp_get(oid)


    # Función para enviar comando de selección de vista
    def select_view(self, view: int):
        """Selecciona una vista específica en el dispositivo"""
        oid = '1.3.6.1.4.1.2566.127.1.1.157.3.15.1.0'
        self.snmp_set(oid, view)


    # Función para obtener solo los ids que corresponden a televisión, descartando radio
    def get_service_ids(self):
        radio_services = ['BLU Radio', 'RCN LA RADIO', 'LA FM RADIO', 'RCN RADIO UNO', 'LA MEGA', 'RADIO NACIONAL','RADIONICA', 'EXPLOREMOS']
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.2.1.3.{self.mac_address}.1.{self.table_number}'
        service_names = self.snmp_bulk_walk(oid)

        # Diccionario con solo oids-valor que contengan servicios de tv
        filtered_service_names = {k: v for k, v in service_names.items() if v not in radio_services}

        return [k.split('.')[-1] for k in filtered_service_names.keys()]


    # Función para obtener los nombres de los servicios
    def get_service_names(self):
        """Obtiene los servicios SNMP disponibles en el dispositivo"""
        filtered_ids = self.get_service_ids()
        service_names = []
        for id in filtered_ids:
            oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.2.1.3.{self.mac_address}.1.{self.table_number}.{id}'
            service_name = self.snmp_get(oid)
            service_names.append(service_name)
        
        return service_names
    

    # Función para obtener el actual network
    def get_actual_network(self):
        """Obtiene el ID de red"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.4.1.17.{self.mac_address}.1.{self.table_number}'
        result = self.snmp_bulk_walk(oid)

        for word in result.values():
            if 'Actual Network' in word:
                actual_network = word.split(sep='  ')[1]
                break
            else:
                actual_network = 'ND'

        return actual_network
    

    # Función para obtener el transport stream ID
    def get_ts_id(self):
        """Obtiene el ID de TS"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.5.1.1.4.{self.mac_address}.1.{self.table_number}'

        return self.snmp_get_next(oid)
    

    # Función para enviar comando "clear statistics and log"
    def clear_statistics_and_log(self):
        """Limpia las estadísticas y logs del dispositivo"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.1.3.9.1.4.{self.mac_address}.1'
        self.snmp_set(oid, 1)


    # Función para enviar comando de iniciar monitoreo
    def start_monitoring(self):
        """Inicia la monitorización del dispositivo"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.1.3.9.1.1.{self.mac_address}.1'
        self.snmp_set(oid, 1)


    # Función para enviar comando de parar monitoreo
    def stop_monitoring(self):
        """Detiene la monitorización del dispositivo"""
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.1.3.9.1.1.{self.mac_address}.1'
        self.snmp_set(oid, 2)


    # Función para obtener el tiempo transcurrido desde el inicio del monitoreo
    def get_elapsed_time(self):
        """Obtiene el tiempo transcurrido desde el inicio de la monitorización"""
        oid = f'.1.3.6.1.4.1.2566.127.1.1.157.3.1.3.9.1.5.{self.mac_address}.1'

        return int(self.snmp_get(oid))


    # Función para obtener la tabla de log durante un minuto
    def get_log_table(self):
        """Obtiene la tabla de logs del dispositivo"""
        # self.start_monitoring()
        MONITORING_TIME = 60
        log_table = {}
        oid = f'1.3.6.1.4.1.2566.127.1.1.157.3.13.3.1.5.{self.mac_address}.1' # LogEntryEvent

        self.clear_statistics_and_log()

        while self.snmp_get(f'1.3.6.1.4.1.2566.127.1.1.157.3.13.2.1.2.{self.mac_address}.1') == '0':
            pass
        
        start_oid = f'{oid}.0'
        end_number = self.snmp_get(f'1.3.6.1.4.1.2566.127.1.1.157.3.13.2.1.2.{self.mac_address}.1')
        end_oid = f'{oid}.{end_number}'

        start_time = time.time()
        while self.get_elapsed_time() < MONITORING_TIME or time.time() - start_time < MONITORING_TIME + 10:
            response = self.snmp_bulk_walk_range(start_oid, end_oid)
            log_table.update(response)

            start_oid = end_oid
            end_number = self.snmp_get(f'1.3.6.1.4.1.2566.127.1.1.157.3.13.2.1.2.{self.mac_address}.1')
            end_oid = f'{oid}.{end_number}'

        # self.stop_monitoring()
        
        return log_table


    # Función para obtener los errores de la tabla log leída
    @staticmethod
    def get_errors(log_table: dict):
        """Obtiene los errores de la tabla de logs"""
        PRIORITIES = {
            1: {
                'TS Sync': [3, 4],
                'Sync Byte': [7, 8],
                'PAT': [11, 13, 14],
                'Continuity Count': [17, 18, 19, 20],
                'PMT': [23, 24, 25],
                'PID': [28, 29, 30]
            },
            2: {
                'Transport': [36],
                'CRC': list(range(39, 52)),
                'PCR Repetition': [74],
                'PCR Discontinuity': [70],
                'PCR Jitter': [78],
                'PTS Repetition': [82],
                'CAT': [86, 87]
            },
            3: {
                'SIRepetition': [99, 100, 101, 104, 105, 106, 111, 112, 113, 114, 115, 119, 121, 122, 123],
                'NITActual': [131, 132],
                'SDTActual': [141, 142],
                'EITActual': [151, 152, 153],
                'EITPresent/following': [155, 156],
                'RST': [166],
                'TDT': [170, 171]
            }
        }

        resultados = {1: set(), 2: set(), 3: set()}  # Usamos sets para evitar duplicados

        for log in map(int, log_table.values()):
            for prioridad, dic in PRIORITIES.items():
                if (prioridad == 1 and log <= 30) or (prioridad == 2 and 36 <= log <= 87) or (prioridad == 3 and log >= 131):
                    resultados[prioridad].update(k for k, v in dic.items() if log in v)

        # Convertimos en strings y reemplazamos vacíos con 'NA'
        priority_1 = ", ".join(sorted(resultados[1])) if resultados[1] else "NA"
        priority_2 = ", ".join(sorted(resultados[2])) if resultados[2] else "NA"
        priority_3 = ", ".join(sorted(resultados[3])) if resultados[3] else "NA"

        return priority_1, priority_2, priority_3


    # Función para obtener un array con todos los resultados del transport stream
    @staticmethod
    def get_result(service_names: list, service_ids: list, actual_network: str, ts_id: str, priority_1: str, priority_2: str, priority_3: str):
        """Obtiene los resultados de la monitorización"""
        result = []
        if len(service_names) == len(service_ids):
            for i in range(len(service_names)):
                result.append([service_names[i], service_ids[i], actual_network, ts_id, priority_1, priority_2, priority_3])
        
        return result
    

    # Función para tomar una captura de pantalla
    @staticmethod
    def take_screenshot(path: str, channel: int, image_number: int):
        img = pyautogui.screenshot(region=(60, 25, 640, 480))
        img.save(f'{path}/TS_{TV_TABLE[channel]}_00{image_number}.png')


    # Función para abrir el escritorio remoto y esperar hasta que se encuentre la imagen del ETL
    def open_remote_desktop(self):
        self.remote_desktop.connect()

        # Se selecciona la vista de diagrama de pastel, que es donde buscará la imagen de apertura
        self.select_view(1)

        while True:
            # self.etl.write_str('INST TSAN') # Configura el instrumento al modo "TS Analyzer"
            self.open_ts_analyzer()
            try: 
                if pyautogui.locateOnScreen('./resources/ETL.png', region = (300, 80, 360, 150), grayscale = True, confidence = 0.99) is not None:
                    time.sleep(5)
                    try:
                        if pyautogui.locateOnScreen('./resources/ETL.png', region = (300, 80, 360, 150), grayscale = True, confidence = 0.99) is not None:
                            self.open_ts_analyzer()
                            break
                    except pyautogui.ImageNotFoundException:
                        continue
            except pyautogui.ImageNotFoundException:
                continue


    # Función para le medición SNMP de un PLP
    def tansport_stream_measurement(self, channel: int, path: str):
        """Obtiene un array con todos los resultados de la medición del transport stream para un PLP"""
        # Se selecciona la vista select view
        self.select_view(1)
        time.sleep(1)

        # Se toma la captura de pantalla del diagrama de pastel.
        self.take_screenshot(path, channel, 1)

        # Se hace el procedimiento para el despliegue del arbol}
        pyautogui.moveTo(60, 220)
        time.sleep(1)
        pyautogui.moveTo(257, 220)
        pyautogui.click()
        pyautogui.drag(95, 0, 0.5, button='left')
        pyautogui.click(187, 220)
        time.sleep(1)

        # Se toma la captura de pantalla del árbol de transport Stream
        self.take_screenshot(path, channel, 2)

        # Se actualiza la tabla tsTree
        self.refresh_table()

        # Se obtiene el número de tabla actual
        self.get_table_number()

        # Se obtienen los valores requeridos en el postprocesamiento
        service_ids = self.get_service_ids()
        service_names = self.get_service_names()
        actual_network = self.get_actual_network()
        ts_id = self.get_ts_id()

        # Se envía el comando Clear Statistics and log of all inputs
        self.select_view(2)
        pyautogui.click(450, 350)
        log_table = self.get_log_table()
        priority_1, priority_2, priority_3 = self.get_errors(log_table)
        self.take_screenshot(path, channel, 3)

        # Se selecciona la vista de Bit Rate y se toma la captura de pantalla
        self.select_view(3)
        time.sleep(1)
        self.take_screenshot(path, channel, 4)

        # Se selecciona la vista de Table repetition y se toma la captura de pantalla
        self.select_view(4)
        time.sleep(1)
        self.take_screenshot(path, channel, 5)

        # Se cierra la ventana de escritrio remoto
        self.remote_desktop.disconnect()

        return self.get_result(service_names, service_ids, actual_network, ts_id, priority_1, priority_2, priority_3)



if __name__ == '__main__':
    from InstrumentController import EtlManager
    etl = EtlManager('172.23.82.20', 75, [])
    instrument = SNMPManager('172.23.82.20', etl.open_ts_analyzer)

    instrument.open_remote_desktop()
    plp_result = instrument.tansport_stream_measurement(16, './tests')
    print(plp_result)
    
    # instrument.snmp_measurement()