import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import threading
from src.core.ReadExcel import ReadExcel
from src.core.InstrumentManager import InstrumentManager
from src.core.InstrumentController import EtlManager, FPHManager, MSDManager
from src.core.MeasurementManager import MeasurementManager
from src.core.ExcelReport import ExcelReport
import os

# Configuración global
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("./src/gui/custom_theme.json")  # Themes: "blue" (standard), "green", "dark-blue"

class DatosCompartidos:
    """Clase para almacenar todos los datos que se comparten entre ventanas"""
    def __init__(self):
        # Datos de la primera ventana
        self.object_preengeneering = None
        self.municipality = None
        self.point = None
        self.measurement_dictionary = None
        self.sfn_dictionary = None
        
        # Datos de la segunda ventana (TV Analógica)
        self.atv_instrument = None
        self.ip_tv_analogica = None
        self.puerto_tv_analogica = None
        self.transductores_tv_analogica = []
        
        # Datos de la tercera ventana (TV Digital)
        self.dtv_instrument = None
        self.ip_tv_digital = None
        self.puerto_tv_digital = None
        self.transductores_tv_digital = []
        
        # Datos de la cuarta ventana (banco de mediciones)
        self.mbk_instrument = None
        self.ip_banco = None
        self.puerto_banco = None
        self.transductores_banco = []
        
        # Datos de la quinta ventana (Rotor)
        self.rtr_instrument = None
        self.ip_rotor = None
        self.angulo_actual_rotor = 0
        
        # Datos de la sexta ventana (Formulario)
        self.site_dictionary = {}
        
        # Datos de la medición
        self.medicion_completada = False
        self.resultado_medicion = None

class AsistenteInstalacion(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuración de la ventana principal
        self.title("Automatización de medición")
        self.geometry("510x650")
        
        # Inicializar datos compartidos
        self.datos = DatosCompartidos()
        
        # Crear contenedor para las ventanas
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Diccionario para almacenar las páginas
        self.frames = {}
        
        # Crear todas las páginas
        for F in (VentanaBienvenida, VentanaAnalogica, VentanaDigital, VentanaBandas, 
                  VentanaRotor, VentanaFormulario, VentanaResumen):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        # Mostrar la primera página
        self.mostrar_ventana(VentanaBienvenida)
    
    def mostrar_ventana(self, cont):
        """Trae al frente la ventana especificada"""
        frame = self.frames[cont]
        frame.tkraise()
        # Actualizar la ventana al mostrarla
        if hasattr(frame, 'actualizar'):
            frame.actualizar()

class VentanaBienvenida(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Bienvenido al Asistente", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Descripción
        descripcion = ctk.CTkLabel(self, text="Este asistente le guiará a través del proceso de configuración.\n"
                                  "Por favor, cargue un archivo Excel  de preingeniería para comenzar.")
        descripcion.pack(pady=10, padx=10)
        
        # Frame para el botón de carga
        frame_carga = ctk.CTkFrame(self)
        frame_carga.pack(pady=20, padx=10, fill="x")
        
        # Estado de archivo
        self.lbl_archivo = ctk.CTkLabel(frame_carga, text="Ningún archivo cargado")
        self.lbl_archivo.pack(side="left", padx=10, pady=10)
        
        # Botón de carga
        self.btn_cargar = ctk.CTkButton(frame_carga, text="Cargar Excel", 
                                   command=self.cargar_excel)
        self.btn_cargar.pack(side="right", padx=10, pady=10)
        
        # Frame para listas desplegables
        frame_listas = ctk.CTkFrame(self)
        frame_listas.pack(pady=20, padx=10, fill="x")
        
        # Primera lista desplegable
        lbl_lista1 = ctk.CTkLabel(frame_listas, text="Selección de municipio:")
        lbl_lista1.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.lista1 = ctk.CTkOptionMenu(frame_listas, values=["Cargue el archivo de preingeniería"], command=self.update_points)
        self.lista1.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.lista1.configure(state="disabled")
        
        # Segunda lista desplegable
        lbl_lista2 = ctk.CTkLabel(frame_listas, text="Selección del punto:")
        lbl_lista2.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.lista2 = ctk.CTkOptionMenu(frame_listas, values=["Seleccione un municipio"])
        self.lista2.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.lista2.configure(state="disabled")
        
        # Configurar columnas
        frame_listas.columnconfigure(1, weight=1)
        
        # Botón para avanzar
        self.btn_siguiente = ctk.CTkButton(self, text="Siguiente", 
                                      command=lambda: self.avanzar())
        self.btn_siguiente.pack(side="bottom", padx=10, pady=20)
        self.btn_siguiente.configure(state="disabled")
    
    def cargar_excel(self):
        """Cargar archivo Excel y poblar las listas desplegables"""
        ruta_archivo = filedialog.askopenfilename(
            title="Seleccionar archivo de preingeniería",
            filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*"))
        )
        
        if ruta_archivo:
            try:
                # Guardar en los datos compartidos
                self.controller.datos.object_preengeneering = ReadExcel(ruta_archivo)
                
                # Actualizar la etiqueta
                self.lbl_archivo.configure(text=f"Archivo cargado: {ruta_archivo.split('/')[-1]}")
                
                # Poblar las listas desplegables con las columnas del Excel
                municipalities = self.controller.datos.object_preengeneering.get_municipalities()
                
                # Habilitar y actualizar primera lista
                self.lista1.configure(values=municipalities, state="normal")
                self.lista1.set('Seleccione un municipio')
                
            except Exception as e:
                self.lbl_archivo.configure(text=f"Error al cargar archivo: {str(e)}")

    def update_points(self, municipality):
        self.municipality = municipality
        self.controller.datos.municipality = municipality
        self.controller.datos.site_dictionary['municipality'] = municipality
        department = self.controller.datos.object_preengeneering.get_department(municipality)
        self.controller.datos.site_dictionary['department'] = department
        number_of_points = self.controller.datos.object_preengeneering.get_number_of_points(municipality)
        points = [str(i + 1) for i in range(number_of_points)]
        
        # Habilitar y actualizar segunda lista
        self.lista2.configure(values=points, state="normal")
        self.lista2.set('Seleccione un punto')
        
        # Cuando se cambia el municipio, eliminar el frame de estaciones si existe
        if hasattr(self, 'frame_estaciones_container'):
            self.frame_estaciones_container.destroy()
        
        # Habilitar botón siguiente
        self.btn_siguiente.configure(state="normal")
        
        # Añadir comando a la segunda lista para actualizar información al seleccionar punto
        self.lista2.configure(command=self.mostrar_estaciones)

    def mostrar_estaciones(self, punto):
        """Muestra las estaciones disponibles para el punto seleccionado"""
        self.punto = int(punto)
        self.controller.datos.point = self.punto
        self.controller.datos.site_dictionary['point'] = self.punto
        # Eliminar el frame anterior si existe
        if hasattr(self, 'frame_estaciones_container'):
            self.frame_estaciones_container.destroy()
        
        # Obtener diccionarios
        diccionario_medición = self.controller.datos.object_preengeneering.get_dictionary(self.municipality, self.punto)
        self.controller.datos.measurement_dictionary = diccionario_medición
        diccionario_sfn = self.controller.datos.object_preengeneering.get_sfn(diccionario_medición)
        self.controller.datos.sfn_dictionary = diccionario_sfn

        # Crear contenedor principal con frame desplazable
        self.frame_estaciones_container = ctk.CTkFrame(self)
        self.frame_estaciones_container.pack(pady=(20,0), padx=10, fill="both", expand=True)
        
        # Crear el frame desplazable
        frame_scroll = ctk.CTkScrollableFrame(self.frame_estaciones_container)
        frame_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Título de la sección
        titulo_estaciones = ctk.CTkLabel(frame_scroll, text="Resumen de la medición",
                                    font=ctk.CTkFont(size=16, weight="bold"))
        titulo_estaciones.pack(pady=(10,15), padx=10)
        
        # Iterar sobre cada estación para mostrar su información
        for estacion, datos in diccionario_medición.items():
            # Frame para cada estación
            frame_estacion = ctk.CTkFrame(frame_scroll)
            frame_estacion.pack(pady=5, padx=10, fill="x")
            
            # Título de la estación con acimuth
            lbl_estacion = ctk.CTkLabel(frame_estacion, 
                                    text=f"{estacion} - {datos['Acimuth']}°",
                                    font=ctk.CTkFont(size=14, weight="bold"))
            lbl_estacion.pack(pady=(5,5), padx=10, anchor="w")
            
            # Sección Analógica
            if 'Analógico' in datos and datos['Analógico']:
                frame_analogico = ctk.CTkFrame(frame_estacion)
                frame_analogico.pack(pady=2, padx=10, fill="x")
                
                lbl_analogico = ctk.CTkLabel(frame_analogico, text="Analógico",
                                        font=ctk.CTkFont(weight="bold"))
                lbl_analogico.pack(pady=(2,2), padx=10, anchor="w")
                
                # Mostrar canales analógicos
                for servicio, canal in datos['Analógico'].items():
                    lbl_canal = ctk.CTkLabel(frame_analogico, text=f"{servicio}: Canal {canal}")
                    lbl_canal.pack(pady=1, padx=30, anchor="w")
            
            # Sección Digital
            if 'Digital' in datos and datos['Digital']:
                frame_digital = ctk.CTkFrame(frame_estacion)
                frame_digital.pack(pady=2, padx=10, fill="x")
                
                lbl_digital = ctk.CTkLabel(frame_digital, text="Digital",
                                        font=ctk.CTkFont(weight="bold"))
                lbl_digital.pack(pady=(2,2), padx=10, anchor="w")
                
                # Mostrar canales digitales
                for servicio, canal in datos['Digital'].items():
                    lbl_canal = ctk.CTkLabel(frame_digital, text=f"{servicio}: Canal {canal}")
                    lbl_canal.pack(pady=1, padx=30, anchor="w")

        # Sección de Canales SFN
        if diccionario_sfn:  # Verificamos que el diccionario no esté vacío
            # Separador
            separador = ctk.CTkFrame(frame_scroll, height=2)
            separador.pack(fill="x", padx=10, pady=(15,0))
            
            # Título de la sección SFN
            titulo_sfn = ctk.CTkLabel(frame_scroll, text="Canales en SFN",
                                    font=ctk.CTkFont(size=16, weight="bold"))
            titulo_sfn.pack(pady=(15,10), padx=10)
            
            # Frame para canales SFN
            frame_sfn = ctk.CTkFrame(frame_scroll)
            frame_sfn.pack(pady=10, padx=10, fill="x")
            
            # Mostrar información de canales SFN
            for canal, estaciones in sorted(diccionario_sfn.items()):
                # Combinar las estaciones en un solo texto
                estaciones_texto = ", ".join(estaciones.keys())
                
                # Mostrar canal y sus estaciones
                lbl_canal_sfn = ctk.CTkLabel(frame_sfn, 
                                        text=f"Canal {canal}: {estaciones_texto}",
                                        font=ctk.CTkFont(size=13))
                lbl_canal_sfn.pack(pady=3, padx=10, anchor="w")
    
    def avanzar(self):
        """Guardar selecciones y avanzar a la siguiente ventana"""
        # Guardar selecciones en datos compartidos
        self.controller.datos.municipality = self.lista1.get()
        self.controller.datos.point = self.lista2.get()
        
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaAnalogica)

class VentanaAnalogica(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Instrumento de TV analógica", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Frame para la configuración de IP
        frame_ip = ctk.CTkFrame(self)
        frame_ip.pack(pady=10, padx=10, fill="x")
        
        # Campo de IP
        lbl_ip = ctk.CTkLabel(frame_ip, text="Dirección IP:")
        lbl_ip.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_ip = ctk.CTkEntry(frame_ip, placeholder_text="192.168.1.108")
        self.entry_ip.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Frame para la configuración de puerto
        frame_puerto = ctk.CTkFrame(self)
        frame_puerto.pack(pady=10, padx=10, fill="x")
        
        # Selección de puerto
        lbl_puerto = ctk.CTkLabel(frame_puerto, text="Seleccione el puerto:")
        lbl_puerto.pack(padx=10, pady=5, anchor="w")
        
        # Opciones de puerto (radiobuttons)
        self.var_puerto = ctk.StringVar(value=50)
        
        rb_puerto_50 = ctk.CTkRadioButton(frame_puerto, text="Puerto 50Ω", 
                                      variable=self.var_puerto, value=50)
        rb_puerto_50.pack(padx=30, pady=5, anchor="w", side="left")
        
        rb_puerto_75 = ctk.CTkRadioButton(frame_puerto, text="Puerto 75Ω", 
                                      variable=self.var_puerto, value=75)
        rb_puerto_75.pack(padx=30, pady=5, anchor="w", side="right")
        
        # Frame para selección de transductores
        frame_transductores = ctk.CTkFrame(self)
        frame_transductores.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Título de transductores
        lbl_transductores = ctk.CTkLabel(frame_transductores, 
                                    text="Seleccione los transductores:")
        lbl_transductores.pack(padx=10, pady=5, anchor="w")
        
        # Lista de transductores
        self.frame_lista_transductores = ctk.CTkScrollableFrame(frame_transductores, height=50)
        self.frame_lista_transductores.pack(padx=10, pady=5, fill="both", expand=True)
        
        # Agregar transductores
        self.vars_transductores = []
        transductores = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']
        
        for i, transductor in enumerate(transductores):
            var = ctk.BooleanVar(value=False)
            self.vars_transductores.append(var)
            
            cb = ctk.CTkCheckBox(self.frame_lista_transductores, text=transductor, 
                             variable=var)
            cb.pack(padx=10, pady=5, anchor="w")

        # Frame para la conexión
        frame_conexion = ctk.CTkFrame(self)
        frame_conexion.pack(pady=10, padx=10, fill="x")

        # Botón de conexión
        self.btn_conectar = ctk.CTkButton(frame_conexion, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(side="left", padx=10, pady=10)

        # Estado de conexión
        self.lbl_estado = ctk.CTkLabel(frame_conexion, text="No conectado")
        self.lbl_estado.pack(side="right", padx=10, pady=10)

        # Frame para los botones de navegación
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaBienvenida))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
    
    def conectar(self):
        """Crear la conexión con el instrumento de TV analógica"""
        ip = self.entry_ip.get()
        
        if not ip:
            self.lbl_estado.configure(text="Error: IP no válida")
            return
        
        # Simular proceso de conexión
        self.lbl_estado.configure(text="Conectando...")
        self.btn_conectar.configure(state="disabled")
        
        # Usar hilos para evitar que la interfaz se congele
        def proceso_conexion():
            # Leer transductores seleccionados
            transductores = []
            transductores_ejemplo = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']
            
            for i, var in enumerate(self.vars_transductores):
                if var.get():
                    transductores.append(transductores_ejemplo[i])

            if any("TELEVES" in transductor for transductor in transductores):
                self.controller.datos.site_dictionary['a_antenna_brand'] = "Televes"
                self.controller.datos.site_dictionary['a_antenna_model'] = "DAT BOSS"
            elif any("BICOLOG" in transductor for transductor in transductores):
                self.controller.datos.site_dictionary['a_antenna_brand'] = "Aaronia"
                self.controller.datos.site_dictionary['a_antenna_model'] = "Bicolog"
            else:
                self.controller.datos.site_dictionary['a_antenna_brand'] = "Rohde&Schwarz"
                self.controller.datos.site_dictionary['a_antenna_model'] = transductores[0].title()

            # Leer la impedancia seleccionada
            impedance = self.var_puerto.get()

            # Código para conectar con el dispositivo
            try:
                # Creación del objeto de conexión
                atv_instrument = EtlManager(ip, impedance, transductores)

                # Actualización de los datos compartidos
                self.controller.datos.atv_instrument = atv_instrument
                self.controller.datos.ip_tv_analogica = atv_instrument.ip_address
                self.controller.datos.puerto_tv_analogica = atv_instrument.impedance
                self.controller.datos.transductores_tv_analogica = atv_instrument.transducers
                instrument_model_name = atv_instrument.instrument_model_name

                self.controller.datos.site_dictionary['instrument_type'] = "Analizador de Televisión"
                self.controller.datos.site_dictionary['instrument_brand'] = atv_instrument.manufacturer
                self.controller.datos.site_dictionary['instrument_model'] = atv_instrument.instrument_model_name
                self.controller.datos.site_dictionary['instrument_serial'] = atv_instrument.instrument_serial_number.split("/")[0]

                conexion_exitosa = True
            except:
                conexion_exitosa = False
                instrument_model_name = ''
            
            # Actualizar la UI desde el hilo principal
            self.after(0, lambda: self.actualizar_estado_conexion(ip, conexion_exitosa, instrument_model_name))
        
        # Iniciar el hilo de conexión
        threading.Thread(target=proceso_conexion, daemon=True).start()
    
    def actualizar_estado_conexion(self, ip, exitoso, instrument_model_name):
        """Actualizar la UI después del intento de conexión"""
        if exitoso:
            self.lbl_estado.configure(text=f"Conectado a {instrument_model_name}: {ip}")
        else:
            self.lbl_estado.configure(text=f"Error al conectar con {ip}")
        
        self.btn_conectar.configure(state="normal")
    
    def avanzar(self):
        """Guardar datos y avanzar a la siguiente ventana"""
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaDigital)

class VentanaDigital(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Instrumento de TV Digital", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Frame para la configuración de IP
        frame_ip = ctk.CTkFrame(self)
        frame_ip.pack(pady=10, padx=10, fill="x")
        
        # Campo de IP
        lbl_ip = ctk.CTkLabel(frame_ip, text="Dirección IP:")
        lbl_ip.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_ip = ctk.CTkEntry(frame_ip, placeholder_text="192.168.1.108")
        self.entry_ip.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Frame para la configuración de puerto
        frame_puerto = ctk.CTkFrame(self)
        frame_puerto.pack(pady=10, padx=10, fill="x")
        
        # Selección de puerto
        lbl_puerto = ctk.CTkLabel(frame_puerto, text="Seleccione el puerto:")
        lbl_puerto.pack(padx=10, pady=5, anchor="w")
        
        # Opciones de puerto (radiobuttons)
        self.var_puerto = ctk.StringVar(value=50)
        
        rb_puerto_50 = ctk.CTkRadioButton(frame_puerto, text="Puerto 50Ω", 
                                      variable=self.var_puerto, value=50)
        rb_puerto_50.pack(padx=30, pady=5, anchor="w", side="left")
        
        rb_puerto_75 = ctk.CTkRadioButton(frame_puerto, text="Puerto 75Ω", 
                                      variable=self.var_puerto, value=75)
        rb_puerto_75.pack(padx=30, pady=5, anchor="w", side="right")
        
        # Frame para selección de transductores
        frame_transductores = ctk.CTkFrame(self)
        frame_transductores.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Título de transductores
        lbl_transductores = ctk.CTkLabel(frame_transductores, 
                                    text="Seleccione los transductores:")
        lbl_transductores.pack(padx=10, pady=5, anchor="w")
        
        # Lista de transductores
        self.frame_lista_transductores = ctk.CTkScrollableFrame(frame_transductores, height=50)
        self.frame_lista_transductores.pack(padx=10, pady=5, fill="both", expand=True)
        
        # Agregar transductores
        self.vars_transductores = []
        transductores = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']
        
        for i, transductor in enumerate(transductores):
            var = ctk.BooleanVar(value=False)
            self.vars_transductores.append(var)
            
            cb = ctk.CTkCheckBox(self.frame_lista_transductores, text=transductor, 
                             variable=var)
            cb.pack(padx=10, pady=5, anchor="w")
        
        # Frame para la conexión
        frame_conexion = ctk.CTkFrame(self)
        frame_conexion.pack(pady=10, padx=10, fill="x")

        # Botón de conexión
        self.btn_conectar = ctk.CTkButton(frame_conexion, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(side="left", padx=10, pady=10)
        
        # Estado de conexión
        self.lbl_estado = ctk.CTkLabel(frame_conexion, text="No conectado")
        self.lbl_estado.pack(side="right", padx=10, pady=10)
        
        # Configurar columnas
        # frame_ip.columnconfigure(1, weight=1)

        # Frame para los botones de navegación
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaAnalogica))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
    
    def conectar(self):
        """Simular la conexión con el instrumento de TV digital"""
        ip = self.entry_ip.get()
        
        if not ip:
            self.lbl_estado.configure(text="Error: IP no válida")
            return
        
        # Proceso de conexión
        self.lbl_estado.configure(text="Conectando...")
        self.btn_conectar.configure(state="disabled")
        
        # Usar hilos para evitar que la interfaz se congele
        def proceso_conexion():
            # Leer transductores seleccionados
            transductores = []
            transductores_ejemplo = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']

            for i, var in enumerate(self.vars_transductores):
                if var.get():
                    transductores.append(transductores_ejemplo[i])
            
            if any("TELEVES" in transductor for transductor in transductores):
                self.controller.datos.site_dictionary['d_antenna_brand'] = "Televes"
                self.controller.datos.site_dictionary['d_antenna_model'] = "DAT BOSS"
            elif any("BICOLOG" in transductor for transductor in transductores):
                self.controller.datos.site_dictionary['d_antenna_brand'] = "Aaronia"
                self.controller.datos.site_dictionary['d_antenna_model'] = "Bicolog"
            else:
                self.controller.datos.site_dictionary['d_antenna_brand'] = "Rohde&Schwarz"
                self.controller.datos.site_dictionary['d_antenna_model'] = transductores[0].title()

            # Leer la impedancia seleccionada
            impedance = self.var_puerto.get()
            
            # Código para conectar con el dispositivo
            try:
                # Creación del objeto de conexión
                dtv_instrument = EtlManager(ip, impedance, transductores)

                # Actualización de los datos compartidos
                self.controller.datos.dtv_instrument = dtv_instrument
                self.controller.datos.ip_tv_digital = dtv_instrument.ip_address
                self.controller.datos.puerto_tv_digital = dtv_instrument.impedance
                self.controller.datos.transductores_tv_digital = dtv_instrument.transducers
                instrument_model_name = dtv_instrument.instrument_model_name
                conexion_exitosa = True

                # Llenado del diccionario del sitio
                self.lbl_estado.configure(text="Obteniendo coordenadas...")
                self.btn_conectar.configure(state="disabled")

                latitude, longitude = dtv_instrument.get_coordinates()
                self.controller.datos.site_dictionary['latitude_dec'] = latitude
                self.controller.datos.site_dictionary['longitude_dec'] = longitude

                latitude_dms, longitude_dms = dtv_instrument.decimal_coords_to_dms(latitude, longitude)
                self.controller.datos.site_dictionary['latitude_dms'] = latitude_dms
                self.controller.datos.site_dictionary['longitude_dms'] = longitude_dms

                altitude = dtv_instrument.get_altitude()
                self.controller.datos.site_dictionary['altitude'] = altitude
            except Exception as e:
                print(e)
                conexion_exitosa = False
                instrument_model_name = ''
            
            # Actualizar la UI desde el hilo principal
            self.after(0, lambda: self.actualizar_estado_conexion(ip, conexion_exitosa, instrument_model_name))
        
        # Iniciar el hilo de conexión
        threading.Thread(target=proceso_conexion, daemon=True).start()
    
    def actualizar_estado_conexion(self, ip, exitoso, instrument_model_name):
        """Actualizar la UI después del intento de conexión"""
        if exitoso:
            self.lbl_estado.configure(text=f"Conectado a {instrument_model_name}: {ip}")
        else:
            self.lbl_estado.configure(text=f"Error al conectar con {ip}")
        
        self.btn_conectar.configure(state="normal")
    
    def avanzar(self):
        """Guardar datos y avanzar a la siguiente ventana"""
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaBandas)

class VentanaBandas(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título - ahora está directamente en el frame principal (self)
        titulo = ctk.CTkLabel(self, text="Instrumento de Banco de mediciones", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Contenedor principal con scroll - ahora va después del título
        self.main_container = ctk.CTkScrollableFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame para la pregunta sobre mediciones
        frame_pregunta = ctk.CTkFrame(self.main_container)
        frame_pregunta.pack(pady=10, padx=10, fill="x")
        
        # Pregunta sobre mediciones
        lbl_pregunta = ctk.CTkLabel(frame_pregunta, text="¿Desea realizar banco de mediciones en este punto?")
        lbl_pregunta.pack(padx=10, pady=5, anchor="w")
        
        # Variable para la respuesta
        self.var_realizar_mediciones = ctk.StringVar(value="no")
        
        # Opciones para realizar mediciones
        rb_mediciones_si = ctk.CTkRadioButton(frame_pregunta, text="Sí", 
                                           variable=self.var_realizar_mediciones,
                                           value="si",
                                           command=self.actualizar_estado_widgets)
        rb_mediciones_si.pack(padx=30, pady=5, anchor="w")
        
        rb_mediciones_no = ctk.CTkRadioButton(frame_pregunta, text="No", 
                                           variable=self.var_realizar_mediciones,
                                           value="no",
                                           command=self.actualizar_estado_widgets)
        rb_mediciones_no.pack(padx=30, pady=5, anchor="w")
        
        # Frame para selección de instrumento
        self.frame_instrumento = ctk.CTkFrame(self.main_container)
        self.frame_instrumento.pack(pady=10, padx=10, fill="x")
        
        # Selección de instrumento
        lbl_instrumento = ctk.CTkLabel(self.frame_instrumento, text="Seleccione el tipo de instrumento:")
        lbl_instrumento.pack(padx=10, pady=5, anchor="w")
        
        # Variable para el tipo de instrumento
        self.var_tipo_instrumento = ctk.StringVar(value="ETL")
        
        # Opciones para el tipo de instrumento
        self.rb_instrumento_etl = ctk.CTkRadioButton(self.frame_instrumento, text="ETL", 
                                             variable=self.var_tipo_instrumento,
                                             value="ETL")
        self.rb_instrumento_etl.pack(padx=30, pady=5, anchor="w")
        
        self.rb_instrumento_fph = ctk.CTkRadioButton(self.frame_instrumento, text="FPH", 
                                             variable=self.var_tipo_instrumento,
                                             value="FPH")
        self.rb_instrumento_fph.pack(padx=30, pady=5, anchor="w")
        
        # Frame para la configuración de IP
        self.frame_ip = ctk.CTkFrame(self.main_container)
        self.frame_ip.pack(pady=10, padx=10, fill="x")
        
        # Campo de IP
        lbl_ip = ctk.CTkLabel(self.frame_ip, text="Dirección IP:")
        lbl_ip.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_ip = ctk.CTkEntry(self.frame_ip, placeholder_text="192.168.1.3")
        self.entry_ip.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Frame para la configuración de puerto
        self.frame_puerto = ctk.CTkFrame(self.main_container)
        self.frame_puerto.pack(pady=10, padx=10, fill="x")
        
        # Selección de puerto
        lbl_puerto = ctk.CTkLabel(self.frame_puerto, text="Seleccione el puerto:")
        lbl_puerto.pack(padx=10, pady=5, anchor="w")
        
        # Opciones de puerto (radiobuttons)
        self.var_puerto = ctk.StringVar(value=50)
        
        self.rb_puerto_50 = ctk.CTkRadioButton(self.frame_puerto, text="Puerto 50Ω",
                                      variable=self.var_puerto, value=50)
        self.rb_puerto_50.pack(padx=30, pady=5, anchor="w")
        
        self.rb_puerto_75 = ctk.CTkRadioButton(self.frame_puerto, text="Puerto 75Ω",
                                      variable=self.var_puerto, value=75)
        self.rb_puerto_75.pack(padx=30, pady=5, anchor="w")
        
        # Frame para selección de transductores
        self.frame_transductores = ctk.CTkFrame(self.main_container)
        self.frame_transductores.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Título de transductores
        lbl_transductores = ctk.CTkLabel(self.frame_transductores, 
                                    text="Seleccione los transductores:")
        lbl_transductores.pack(padx=10, pady=5, anchor="w")
        
        # Lista de transductores
        self.frame_lista_transductores = ctk.CTkScrollableFrame(self.frame_transductores, height=150)
        self.frame_lista_transductores.pack(padx=10, pady=5, fill="both", expand=True)
        
        # Agregar transductores de ejemplo
        self.vars_transductores = []
        self.transductores_widgets = []
        transductores_ejemplo = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']
        
        for i, transductor in enumerate(transductores_ejemplo):
            var = ctk.BooleanVar(value=False)
            self.vars_transductores.append(var)
            
            cb = ctk.CTkCheckBox(self.frame_lista_transductores, text=transductor, 
                             variable=var)
            cb.pack(padx=10, pady=5, anchor="w")
            self.transductores_widgets.append(cb)

        # Frame para la conexión
        self.frame_conexion = ctk.CTkFrame(self.main_container)
        self.frame_conexion.pack(pady=10, padx=10, fill="x")

        # Botón de conexión
        self.btn_conectar = ctk.CTkButton(self.frame_conexion, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(side="left", padx=10, pady=10)
        
        # Estado de conexión
        self.lbl_estado = ctk.CTkLabel(self.frame_conexion, text="No conectado")
        self.lbl_estado.pack(side="right", padx=10, pady=10)
        
        # Configurar columnas
        self.frame_conexion.columnconfigure(1, weight=1)
        
        # Frame para los botones de navegación - se coloca directamente en el self
        # para que quede siempre en la parte inferior
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaDigital))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
        
        # Inicializar el estado de los widgets
        self.actualizar_estado_widgets()
    
    def actualizar_estado_widgets(self):
        """Habilitar o deshabilitar widgets según la selección de realizar mediciones"""
        estado = "normal" if self.var_realizar_mediciones.get() == "si" else "disabled"
        
        # En lugar de intentar configurar el estado de los frames, 
        # configuramos los widgets interactivos individuales
        
        # Configurar estado de los radio buttons de instrumento
        self.rb_instrumento_etl.configure(state=estado)
        self.rb_instrumento_fph.configure(state=estado)
        
        # Configurar estado de la entrada de IP y botón conectar
        self.entry_ip.configure(state=estado)
        self.btn_conectar.configure(state=estado)
        
        # Configurar estado de los radio buttons de puerto
        self.rb_puerto_50.configure(state=estado)
        self.rb_puerto_75.configure(state=estado)
        
        # Configurar estado de los checkboxes de transductores
        for cb in self.transductores_widgets:
            cb.configure(state=estado)
            
        # Ajustar la visibilidad de los frames si está deshabilitado
        if estado == "disabled":
            self.frame_instrumento.pack_forget()
            self.frame_ip.pack_forget()
            self.frame_puerto.pack_forget()
            self.frame_transductores.pack_forget()
        else:
            # Asegurarse de que los frames estén visibles en el orden correcto
            frame_pregunta = self.main_container.winfo_children()[0]
            
            self.frame_instrumento.pack(after=frame_pregunta, pady=10, padx=10, fill="x")
            self.frame_ip.pack(after=self.frame_instrumento, pady=10, padx=10, fill="x")
            self.frame_puerto.pack(after=self.frame_ip, pady=10, padx=10, fill="x")
            self.frame_transductores.pack(after=self.frame_puerto, pady=10, padx=10, fill="both", expand=True)
    
    def conectar(self):
        """Simular la conexión con el instrumento para bandas"""
        ip = self.entry_ip.get()
        
        if not ip:
            self.lbl_estado.configure(text="Error: IP no válida")
            return
        
        # Proceso de conexión
        self.lbl_estado.configure(text="Conectando...")
        self.btn_conectar.configure(state="disabled")
        
        # Usar hilos para evitar que la interfaz se congele
        def proceso_conexion():
             # Leer transductores seleccionados
            transductores = []
            transductores_ejemplo = ['TELEVES', 'CABLE TELEVES', 'BICOLOG 20300', 'CABLE BICOLOG', 'HE200', 'HL223', 'CABLE']
            
            for i, var in enumerate(self.vars_transductores):
                if var.get():
                    transductores.append(transductores_ejemplo[i])

            # Leer la impedancia seleccionada
            impedance = self.var_puerto.get()

            # Leer tipo de instrumento
            tipo_instrumento = self.var_tipo_instrumento.get()

            if tipo_instrumento == "ETL":
                # Código para conectar con el dispositivo
                try:
                    # Creación del objeto de conexión
                    mbk_instrument = EtlManager(ip, impedance, transductores)

                    # Actualización de los datos compartidos
                    self.controller.datos.mbk_instrument = mbk_instrument
                    self.controller.datos.ip_bandas = mbk_instrument.ip_address
                    self.controller.datos.puerto_bandas = mbk_instrument.impedance
                    self.controller.datos.transductores_bandas = mbk_instrument.transducers
                    instrument_model_name = mbk_instrument.instrument_model_name

                    date_for_mbk_folder = mbk_instrument.get_date_for_bank_folder()
                    self.controller.datos.site_dictionary['date_for_mbk_folder'] = date_for_mbk_folder

                    conexion_exitosa = True
                except Exception as e:
                    conexion_exitosa = False
                    instrument_model_name = ''
                    print(e)
            elif tipo_instrumento == "FPH":
                # Código para conectar con el dispositivo
                try:
                    # Creación del objeto de conexión
                    mbk_instrument = FPHManager(ip, impedance, [])

                    # Actualización de los datos compartidos
                    self.controller.datos.mbk_instrument = mbk_instrument
                    self.controller.datos.ip_bandas = mbk_instrument.ip_address
                    self.controller.datos.puerto_bandas = mbk_instrument.impedance
                    self.controller.datos.transductores_bandas = mbk_instrument.transducers
                    instrument_model_name = mbk_instrument.instrument_model_name

                    date_for_mbk_folder = mbk_instrument.get_date_for_bank_folder()
                    self.controller.datos.site_dictionary['date_for_mbk_folder'] = date_for_mbk_folder

                    conexion_exitosa = True
                except:
                    conexion_exitosa = False
                    instrument_model_name = ''            
            
            # Actualizar la UI desde el hilo principal
            self.after(0, lambda: self.actualizar_estado_conexion(ip, conexion_exitosa, instrument_model_name))
        
        # Iniciar el hilo de conexión
        threading.Thread(target=proceso_conexion, daemon=True).start()
    
    def actualizar_estado_conexion(self, ip, exitoso, instrument_model_name):
        """Actualizar la UI después del intento de conexión"""
        if exitoso:
            self.lbl_estado.configure(text=f"Conectado a {instrument_model_name}: {ip}")
        else:
            self.lbl_estado.configure(text=f"Error al conectar con {ip}")
        
        # Rehabilitar el botón solo si las mediciones están habilitadas
        if self.var_realizar_mediciones.get() == "si":
            self.btn_conectar.configure(state="normal")
    
    def avanzar(self):
        """Guardar datos y avanzar a la siguiente ventana"""
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaRotor)

class VentanaRotor(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Configuración del Rotor", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Frame para la selección de uso de rotor
        frame_uso = ctk.CTkFrame(self)
        frame_uso.pack(pady=10, padx=10, fill="x")
        
        # Selección de uso de rotor
        lbl_uso = ctk.CTkLabel(frame_uso, text="¿Desea utilizar el rotor?")
        lbl_uso.pack(padx=10, pady=10, anchor="w")
        
        # Variable para la selección
        self.var_uso = ctk.StringVar(value="no")
        
        # Opciones para usar o no el rotor
        rb_si = ctk.CTkRadioButton(frame_uso, text="Sí", 
                               variable=self.var_uso, value="si",
                               command=self.actualizar_estado_campos)
        rb_si.pack(padx=30, pady=5, anchor="w")
        
        rb_no = ctk.CTkRadioButton(frame_uso, text="No", 
                               variable=self.var_uso, value="no",
                               command=self.actualizar_estado_campos)
        rb_no.pack(padx=30, pady=5, anchor="w")
        
        # Frame para la configuración del rotor
        self.frame_config = ctk.CTkFrame(self)
        self.frame_config.pack(pady=10, padx=10, fill="x")
        
        # Campo de IP
        lbl_ip = ctk.CTkLabel(self.frame_config, text="Dirección IP del Rotor:")
        lbl_ip.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_ip = ctk.CTkEntry(self.frame_config, placeholder_text="192.168.1.101")
        self.entry_ip.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Botón de conexión
        self.btn_conectar = ctk.CTkButton(self.frame_config, text="Conectar", 
                                     command=self.conectar)
        self.btn_conectar.grid(row=0, column=2, padx=10, pady=10)
        
        # Estado de conexión
        self.lbl_estado = ctk.CTkLabel(self.frame_config, text="No conectado")
        self.lbl_estado.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="w")
        
        # Ángulo actual
        lbl_angulo = ctk.CTkLabel(self.frame_config, text="Ángulo actual (0-360°):")
        lbl_angulo.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_angulo = ctk.CTkEntry(self.frame_config, placeholder_text="0")
        self.entry_angulo.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        
        # Configurar columnas
        self.frame_config.columnconfigure(1, weight=1)
        
        # Desactivar campos inicialmente
        self.actualizar_estado_campos()
        
        # Frame para los botones de navegación
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaBandas))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
    
    def actualizar_estado_campos(self):
        """Actualizar el estado de los campos según la selección"""
        if self.var_uso.get() == "si":
            # Habilitar campos
            for widget in self.frame_config.winfo_children():
                widget.configure(state="normal")
        else:
            # Deshabilitar campos
            for widget in self.frame_config.winfo_children():
                if isinstance(widget, (ctk.CTkEntry, ctk.CTkButton)):
                    widget.configure(state="disabled")
    
    def conectar(self):
        """Simular la conexión con el rotor"""
        ip = self.entry_ip.get()
        
        if not ip:
            self.lbl_estado.configure(text="Error: IP no válida")
            return
        
        # Simular proceso de conexión
        self.lbl_estado.configure(text="Conectando...")
        self.btn_conectar.configure(state="disabled")
        
        # Usar hilos para evitar que la interfaz se congele
        def proceso_conexion():
            # Código para conectar con el dispositivo
            try:
                # Creación del objeto de conexión
                rtr_instrument = MSDManager(ip)

                # Actualización de los datos compartidos
                self.controller.datos.rtr_instrument = rtr_instrument
                self.controller.datos.ip_rotor = rtr_instrument.ip_address
                instrument_model_name = rtr_instrument.instrument_model_name
                conexion_exitosa = True
            except:
                conexion_exitosa = False
                instrument_model_name = ''
            
            # Actualizar la UI desde el hilo principal
            self.after(0, lambda: self.actualizar_estado_conexion(ip, conexion_exitosa, instrument_model_name))
        
        # Iniciar el hilo de conexión
        threading.Thread(target=proceso_conexion, daemon=True).start()
    
    def actualizar_estado_conexion(self, ip, exitoso, instrument_model_name):
        """Actualizar la UI después del intento de conexión"""
        if exitoso:
            self.lbl_estado.configure(text=f"Conectado a {instrument_model_name}: {ip}")
        else:
            self.lbl_estado.configure(text=f"Error al conectar con {ip}")
        
        self.btn_conectar.configure(state="normal")
    
    def avanzar(self):
        """Guardar datos y avanzar a la siguiente ventana"""
        # Guardar selección de uso de rotor
        usar_rotor = self.var_uso.get() == "si"
        
        if usar_rotor:
            # Guardar ángulo actual
            try:
                angulo = int(self.entry_angulo.get())
                if 0 <= angulo <= 360:
                    self.controller.datos.angulo_actual_rotor = angulo
                else:
                    # Mostrar error si el ángulo está fuera de rango
                    CTkMessagebox(title="Error", message="El ángulo debe estar entre 0 y 360 grados.")
                    return
            except ValueError:
                # Mostrar error si el ángulo no es un número válido
                CTkMessagebox(title="Error", message="Por favor, ingrese un ángulo válido.")
                return
        
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaFormulario)

class VentanaFormulario(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Información del punto de medición", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Frame para el formulario
        frame_formulario = ctk.CTkFrame(self)
        frame_formulario.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Definir opciones para las listas desplegables
        opciones = {
            "Entorno": ["Urbano", "Suburbano", "Rural"],
            "Tipo de terreno": ["Plano", "Montañoso"],
            "Trayecto de la señal": ["LOS", "NLOS", "OT (tráfico sobre vehículos)"],
            "Obstrucción de la señal": ["Colinas", "Montañas", "Lineas de energía", "Ninguna"],
            "Ingeniero responsable 1": [
                "Angela Yulieth Rivera Gomez",
                "Camilo Andrés Velasco Triana",
                "Carolina Villamizar Rivera",
                "Deisy Xiomara Rivera Rozo",
                "Diego Fernando Gutiérrez Junco",
                "Haider Oswaldo Navarro Navarro",
                "Heider Jair Lopez Giraldo",
                "Jennifer Leidy cespedes Menjura",
                "Jesús Alberto Tirano Vargas",
                "German Leonardo Vargas Gutierrrez",
                "Mauricio Quevedo Arias",
                "Miguel Angel Rojas Ruiz",
                "Pamela Vasquez Montoya",
                "Rafael Alexis Berrio Granados",
                "Rodrigo Cardenas Gomez",
                "Sebastian Chavez Martinez",
                "Trady Alexander Avila Vargas"
            ],
            "Ingeniero responsable 2": [
                "Angela Yulieth Rivera Gomez",
                "Camilo Andrés Velasco Triana",
                "Carolina Villamizar Rivera",
                "Deisy Xiomara Rivera Rozo",
                "Diego Fernando Gutiérrez Junco",
                "Haider Oswaldo Navarro Navarro",
                "Heider Jair Lopez Giraldo",
                "Jennifer Leidy cespedes Menjura",
                "Jesús Alberto Tirano Vargas",
                "German Leonardo Vargas Gutierrrez",
                "Mauricio Quevedo Arias",
                "Miguel Angel Rojas Ruiz",
                "Pamela Vasquez Montoya",
                "Rafael Alexis Berrio Granados",
                "Rodrigo Cardenas Gomez",
                "Sebastian Chavez Martinez",
                "Trady Alexander Avila Vargas",
                "NA"
            ]
        }
        
        # Crear listas desplegables con etiquetas
        self.selecciones = {}
        
        for i, (nombre, valores) in enumerate(opciones.items()):
            # Etiqueta
            lbl = ctk.CTkLabel(frame_formulario, text=f"{nombre}:")
            lbl.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            # Lista desplegable
            var = ctk.StringVar(value="Seleccione una opción")
            self.selecciones[nombre] = var
            
            combo = ctk.CTkOptionMenu(frame_formulario, values=valores, variable=var)
            combo.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
            
            # Configurar columnas
            frame_formulario.columnconfigure(1, weight=1)
        
        # Etiqueta para ingresar dirección
        lbl_direccion = ctk.CTkLabel(frame_formulario, text="Dirección:")
        lbl_direccion.grid(row=i+1, column=0, padx=10, pady=5, sticky="w")
        
        # Campo para ingresar dirección
        self.txt_direccion = ctk.CTkEntry(frame_formulario)
        self.txt_direccion.grid(row=i+1, column=1, padx=10, pady=5, sticky="ew")

        # Configurar columnas
        frame_formulario.columnconfigure(1, weight=1)
        
        # Frame para los botones de navegación
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaRotor))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
    
    def avanzar(self):
        """Guardar datos del formulario y avanzar a la siguiente ventana"""
        # Guardar selecciones del formulario
        formulario_datos = {}
        
        for nombre, var in self.selecciones.items():
            formulario_datos[nombre] = var.get()

        # Guardar datos en el diccionario compartido
        self.controller.datos.site_dictionary['around'] = formulario_datos['Entorno']
        self.controller.datos.site_dictionary['terrain'] = formulario_datos['Tipo de terreno']
        self.controller.datos.site_dictionary['signal_path'] = formulario_datos['Trayecto de la señal']
        self.controller.datos.site_dictionary['signal_obstruction'] = formulario_datos['Obstrucción de la señal']
        self.controller.datos.site_dictionary['engineer_1'] = formulario_datos['Ingeniero responsable 1']
        self.controller.datos.site_dictionary['engineer_2'] = formulario_datos['Ingeniero responsable 2']
        self.controller.datos.site_dictionary['address'] = self.txt_direccion.get()
        
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(VentanaResumen)

class VentanaResumen(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Resumen de canales", 
                             font=ctk.CTkFont(size=20, weight="bold"))
        titulo.pack(pady=10, padx=10)
        
        # Frame para el resumen
        self.frame_resumen = ctk.CTkScrollableFrame(self)
        self.frame_resumen.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Frame para los botones de navegación y medición
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(VentanaFormulario))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para iniciar medición
        self.btn_iniciar = ctk.CTkButton(frame_botones, text="Iniciar Medición", 
                                    command=self.iniciar_medicion)
        self.btn_iniciar.pack(side="right", padx=10, pady=10)
        
        # Botón para finalizar, inicialmente oculto
        self.btn_finalizar = ctk.CTkButton(frame_botones, text="Finalizar", 
                                      command=self.finalizar, fg_color="green")
        self.btn_finalizar.pack(side="right", padx=10, pady=10)
        self.btn_finalizar.pack_forget()  # Ocultar inicialmente
        
        # Barra de progreso, inicialmente oculta
        self.progreso = ctk.CTkProgressBar(self)
        self.progreso.set(0)
        self.progreso.pack(pady=10, padx=20, fill="x")
        self.progreso.pack_forget()  # Ocultar inicialmente
        
        # Etiqueta de estado de medición
        self.lbl_estado_medicion = ctk.CTkLabel(self, text="")
        self.lbl_estado_medicion.pack(pady=5)
        self.lbl_estado_medicion.pack_forget()  # Ocultar inicialmente
    
    def actualizar(self):
        """Actualizar el resumen con los datos actuales"""
        # Limpiar el frame de resumen
        for widget in self.frame_resumen.winfo_children():
            widget.destroy()
        
        # Función para agregar secciones al resumen
        def agregar_seccion(titulo, datos):
            # Frame para la sección
            frame_seccion = ctk.CTkFrame(self.frame_resumen)
            frame_seccion.pack(pady=5, padx=5, fill="x")
            
            # Título de la sección
            lbl_titulo = ctk.CTkLabel(frame_seccion, 
                                  text=titulo,
                                  font=ctk.CTkFont(weight="bold"))
            lbl_titulo.pack(padx=10, pady=5, anchor="w")
            
            # Contenido de la sección
            if isinstance(datos, dict):
                for clave, valor in datos.items():
                    # Omitir si el valor es None o lista vacía
                    if valor is None or (isinstance(valor, list) and len(valor) == 0):
                        continue
                    
                    # Frame para el elemento
                    frame_elemento = ctk.CTkFrame(frame_seccion)
                    frame_elemento.pack(pady=2, padx=10, fill="x")
                    
                    # Etiqueta de la clave
                    lbl_clave = ctk.CTkLabel(frame_elemento, text=f"{clave}:")
                    lbl_clave.pack(side="left", padx=5, pady=2)
                    
                    # Valor formateado
                    if isinstance(valor, list):
                        texto_valor = ", ".join(valor)
                    elif isinstance(valor, bool):
                        texto_valor = "Sí" if valor else "No"
                    else:
                        texto_valor = str(valor)
                    
                    lbl_valor = ctk.CTkLabel(frame_elemento, text=texto_valor)
                    lbl_valor.pack(side="right", padx=5, pady=2)
            else:
                # Para casos donde datos no es un diccionario
                lbl_valor = ctk.CTkLabel(frame_seccion, text=str(datos))
                lbl_valor.pack(padx=20, pady=2, anchor="w")
        
        # Obtener datos
        datos = self.controller.datos
        
        # Agregar sección de archivo
        archivo_dict = {
            # "Archivo Excel": datos.archivo_excel.split('/')[-1] if datos.archivo_excel else "No cargado",
            "Municipio": datos.municipality,
            "Punto": datos.point
        }
        agregar_seccion("Archivo de Datos", archivo_dict)
        
        # Agregar sección de TV Analógica
        tv_analogica_dict = {
            "IP": datos.ip_tv_analogica,
            "Puerto": datos.puerto_tv_analogica,
            "Transductores": datos.transductores_tv_analogica
        }
        agregar_seccion("TV Analógica", tv_analogica_dict)
        
        # Agregar sección de TV Digital
        tv_digital_dict = {
            "IP": datos.ip_tv_digital,
            "Puerto": datos.puerto_tv_digital,
            "Transductores": datos.transductores_tv_digital
        }
        agregar_seccion("TV Digital", tv_digital_dict)
        
        # Agregar sección de Bandas
        bandas_dict = {
            "IP": datos.ip_banco,
            "Puerto": datos.puerto_banco,
            "Transductores": datos.transductores_banco
        }
        agregar_seccion("Medidor de banco", bandas_dict)
        
        # Agregar sección de Rotor
        rotor_dict = {
            "Uso de Rotor": "Sí" if datos.rtr_instrument is not None else "No"
        }
        
        if datos.rtr_instrument is not None:
            rotor_dict.update({
                "IP": datos.ip_rotor,
                "Conexión": "Conectado" if datos.rtr_instrument is not None else "No conectado",
                "Ángulo Actual": f"{datos.angulo_actual_rotor}°" if datos.angulo_actual_rotor is not None else "No especificado"
            })
        
        agregar_seccion("Rotor", rotor_dict)
        
        # Agregar sección de Formulario
        # if datos.formulario_datos:
        #     agregar_seccion("Parámetros de Configuración", datos.formulario_datos)
    
    def iniciar_medicion(self):
        """Iniciar el proceso de medición con barra de progreso"""
        # Mostrar barra de progreso y ocultar botón de inicio
        self.progreso.pack(pady=10, padx=20, fill="x")
        self.lbl_estado_medicion.pack(pady=5)
        self.btn_iniciar.pack_forget()
        
        # Define el callback para actualizar la barra de progreso
        def progress_callback(current, total, message=""):
            # Calcular porcentaje
            percent = current / total if total > 0 else 0
            
            # Actualizar la interfaz en el hilo principal
            def update_ui():
                self.progreso.set(percent)
                self.lbl_estado_medicion.configure(
                    text=f"Progreso: {int(percent * 100)}% - {message}"
                )
                # Forzar actualización de la interfaz
                self.update()
            
            # La actualización debe realizarse en el hilo principal
            self.after(0, update_ui)
        
        # Define el callback para la rotación manual
        def manual_rotation_callback(acimuth):
            msg = CTkMessagebox(
                title="Rotación Manual", 
                message=f"Gire el rotor hacia el acimuth {acimuth}°.\nUna vez apuntado, haga click en aceptar.",
                option_1="Aceptar"
            )
            response = msg.get()
            return response
        
        # Define el callback para confirmar o repetir la medición
        def confirm_measurement_callback(canal_info):
            msg = CTkMessagebox(
                title="Verificación de Medición", 
                message=f"Revise los soportes generados para {canal_info}.\nSi están bien, de clic en continuar, en caso contrario, de clic en repetir para repetir la medición.",
                option_1="Continuar",
                option_2="Repetir"
            )
            response = msg.get()
            return response == "Continuar"  # Devuelve True si el usuario selecciona "Continuar"
        
        # Iniciar proceso de medición en un hilo separado
        def start_measurement(self):
            try:
                # Inicializar la barra de progreso
                progress_callback(0, 1, "Preparando instrumentos...")
                
                # Creación del objeto controlador de las mediciones
                measurement_manager = MeasurementManager(
                    atv = self.controller.datos.atv_instrument,
                    dtv = self.controller.datos.dtv_instrument,
                    mbk = self.controller.datos.mbk_instrument,
                    rtr = self.controller.datos.rtr_instrument
                )

                # Definición de la ruta de almacenamiento de soportes para tv
                municipality = self.controller.datos.municipality
                point = self.controller.datos.point
                storage_path_tv = f"./results/{municipality}/P{str(point).zfill(2)}"
                
                # Definición de la ruta de almacenamiento de soportes para tv
                dane_code = self.controller.datos.object_preengeneering.get_dane_code(municipality)
                date = self.controller.datos.site_dictionary['date_for_mbk_folder']
                storage_path_bank = f"./results/{dane_code}_{date}_{municipality.replace(' ', '-').upper()}_P{point}"

                # Verificar si mbk no es None y crear la segunda barra de progreso si es necesario
                mbk_parallel = False
                if measurement_manager.mbk is not None:
                    if (measurement_manager.mbk.ip_address != measurement_manager.dtv.ip_address):
                        mbk_parallel = True

                        # Crear segunda barra de progreso para medición paralela
                        self.progreso_mbk = ctk.CTkProgressBar(self)
                        self.progreso_mbk.set(0)
                        self.progreso_mbk.pack(pady=5, padx=20, fill="x")
                        self.lbl_estado_medicion_mbk = ctk.CTkLabel(self, text="Preparando medición de banco...")
                        self.lbl_estado_medicion_mbk.pack(pady=2)

                # Actualizar progreso
                progress_callback(0, 1, "Realizando mediciones SFN...")

                # Si mbk existe y es paralelo, iniciar medición en hilo separado
                if mbk_parallel:
                    # Callback para actualizar la barra de progreso de mbk
                    def mbk_progress_callback(current, total, message=""):
                        percent = current / total if total > 0 else 0
                        def update_mbk_ui():
                            self.progreso_mbk.set(percent)
                            self.lbl_estado_medicion_mbk.configure(
                                text=f"Progreso de medición de banco: {int(percent * 100)}% - {message}"
                            )
                            self.update()
                        self.after(0, update_mbk_ui)
                    
                    # Iniciar medición de mbk en hilo separado
                    mbk_thread = threading.Thread(
                        target=measurement_manager.mbk_measurement,
                        args=(storage_path_bank, mbk_progress_callback),
                        daemon=True
                    )
                    mbk_thread.start()

                # Hacer medición de SFN en caso de que sea necesario
                if self.controller.datos.sfn_dictionary:
                    sfn_selection = measurement_manager.sfn_measurement(
                        dictionary = self.controller.datos.sfn_dictionary,
                        path = storage_path_tv,
                        park_acimuth = self.controller.datos.angulo_actual_rotor,
                        callback_rotate = manual_rotation_callback
                    )

                    # Actualizar el diccionario de mediciones
                    measurement_dictionary = self.controller.datos.object_preengeneering.update_sfn(
                        self.controller.datos.measurement_dictionary,
                        sfn_selection
                    )
                else:
                    # Si el diccionario de SFN está vacío, el diccionario de medición es el mismo
                    measurement_dictionary = self.controller.datos.measurement_dictionary

                # Para pruebas
                # measurement_dictionary = {'Manjuí': {'Acimuth': 254, 'Analógico': {'Canal Capital': 2}, 'Digital': {'RTVC': 16}}}

                # Medición de TV con la barra de progreso
                atv_result, dtv_result = measurement_manager.tv_measurement(
                    dictionary = measurement_dictionary,
                    park_acimuth = self.controller.datos.angulo_actual_rotor,
                    path = storage_path_tv,
                    callback_rotate = manual_rotation_callback,
                    callback_confirm = confirm_measurement_callback,
                    callback_progress = progress_callback  # Añadimos el callback de progreso
                )

                # Actualizar progreso para post-procesamiento
                progress_callback(0.9, 1, "Generando reportes...")

                # Llenado de postprocesamiento
                report = ExcelReport()
                report.fill_reports(
                    site_dictionary = self.controller.datos.site_dictionary,
                    analog_measurement_dictionary = atv_result,
                    digital_measurement_dictionary = dtv_result,
                    sfn_dictionary = self.controller.datos.sfn_dictionary
                )

                # Ejecutar mbk_measurement secuencialmente si no es paralelo
                if measurement_manager.mbk is not None and not mbk_parallel:
                    print("Se inicia medición de banco secuencial")
                    progress_callback(0.95, 1, "Iniciando medición de banco...")
                    measurement_manager.mbk_measurement(storage_path_bank, progress_callback)

                # Completar medición
                progress_callback(1, 1, "¡Medición completada con éxito!")
                
                # Mostrar botón de finalizar (en el hilo principal)
                def show_finish_button():
                    self.controller.datos.medicion_completada = True
                    self.btn_finalizar.pack(side="right", padx=10, pady=10)
                
                self.after(0, show_finish_button)
            
            except Exception as e:
                # Manejar errores y mostrarlos al usuario
                def show_error():
                    error_msg = f"Error durante la medición: {str(e)}"
                    self.lbl_estado_medicion.configure(text=error_msg)
                    CTkMessagebox(title="Error", message=error_msg, icon="error")
                    # Volver a mostrar el botón de inicio
                    self.btn_iniciar.pack(side="right", padx=10, pady=10)
                
                self.after(0, show_error)
        
        # Iniciar hilo de medición
        threading.Thread(target=start_measurement, args=(self,), daemon=True).start()
    
    def finalizar(self):
        """Finalizar el asistente"""
        # Mostrar mensaje de confirmación
        respuesta = CTkMessagebox(title="Medición Completada", 
                                     message="La medición ha sido completada con éxito.\n"
                                            "¿Desea realizar otra medición?",
                                     icon="info", option_1="Sí", option_2="No")
        
        if respuesta.get() == "Sí":
            # Reiniciar el asistente
            self.reiniciar()
        else:
            # Cerrar la aplicación
            self.controller.quit()
    
    def reiniciar(self):
        """Reiniciar el asistente para una nueva medición"""
        # Reiniciar datos
        self.controller.datos = DatosCompartidos()
        
        # Ocultar componentes de medición
        self.progreso.pack_forget()
        self.lbl_estado_medicion.pack_forget()
        self.btn_finalizar.pack_forget()
        
        # Ocultar segunda barra de progreso si existe
        if hasattr(self, 'progreso_mbk'):
            self.progreso_mbk.pack_forget()
            self.lbl_estado_medicion_mbk.pack_forget()
        
        # Mostrar botón de inicio
        self.btn_iniciar.pack(side="right", padx=10, pady=10)
        
        # Volver a la primera ventana
        self.controller.mostrar_ventana(VentanaBienvenida)

# Función para iniciar la aplicación
def iniciar_aplicacion():
    app = AsistenteInstalacion()
    app.mainloop()

if __name__ == "__main__":
    iniciar_aplicacion()