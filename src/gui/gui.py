import customtkinter as ctk
import tkinter as tk
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import threading
from src.core.ReadExcel import ReadExcel
from src.core.InstrumentManager import InstrumentManager
from src.core.InstrumentController import EtlManager, FPHManager, ViaviManager, MSDManager
from src.core.MeasurementManager import MeasurementManager
from src.core.ExcelReport import ExcelReport
from src.utils.constants import *
import pandas as pd
import pythoncom
from src.utils.utils import rpath

# Configuración global
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme(rpath("./src/gui/custom_theme.json"))  # Themes: "blue" (standard), "green", "dark-blue"

class AutocompleteEntry(ctk.CTkEntry):
    def __init__(self, master, options_list, listbox_width = 300, *args, **kwargs):
        self.var = ctk.StringVar()
        super().__init__(master, textvariable=self.var, *args, **kwargs)

        self.listbox_width = listbox_width

        self.options_list = sorted(options_list)
        self.var.trace_add("write", self.changed)
       
        self.listbox = None
        self.listbox_list = None
        self.current_selection = -1
        self.matches = []
       
        # Bind events para navegación por teclado
        self.bind("<KeyPress>", self.on_key_press)
        self.bind("<FocusOut>", self.on_focus_out)
       
        # Bind para detectar movimiento de ventana
        self.bind("<Configure>", self.on_configure)
       
        # Variables para monitorear cambios de posición
        self.last_x = None
        self.last_y = None
        self.position_check_id = None
        
        # Variables para el manejo de eventos globales
        self.global_bindings = []

        self.global_bindings = []
        self._after_id = None 

    def bind_global_events(self):
        """Vincula eventos globales para detectar clics y scroll fuera del widget"""
        if self.listbox:
            # Obtener la ventana principal
            root = self.winfo_toplevel()
            
            # Bind para detectar clics en toda la ventana
            click_id = root.bind("<Button-1>", self.on_global_click, "+")
            
            # Bind para detectar scroll en toda la ventana
            scroll_id = root.bind("<MouseWheel>", self.on_global_scroll, "+")
            
            # Guardar los IDs para poder desvincularios después
            self.global_bindings = [
                (root, "<Button-1>", click_id),
                (root, "<MouseWheel>", scroll_id)
            ]

    def unbind_global_events(self):
        """Desvincula los eventos globales"""
        for widget, event, binding_id in self.global_bindings:
            try:
                widget.unbind(event, binding_id)
            except:
                pass
        self.global_bindings = []

    def on_global_click(self, event):
        """Maneja clics globales para cerrar el listbox si está fuera del área"""
        if not self.listbox:
            return
            
        # Verificar si el clic fue dentro del entry
        if self.is_point_inside_widget(event.x_root, event.y_root, self):
            return
            
        # Verificar si el clic fue dentro del listbox
        if self.is_point_inside_widget(event.x_root, event.y_root, self.listbox):
            return
            
        # Si llegamos aquí, el clic fue fuera del área, cerrar listbox
        self.close_listbox()

    def on_global_scroll(self, event):
        """Maneja eventos de scroll globales para cerrar el listbox si está fuera del área"""
        if not self.listbox:
            return
            
        # Obtener las coordenadas del mouse
        x_root = event.x_root
        y_root = event.y_root
        
        # Verificar si el scroll fue dentro del entry
        if self.is_point_inside_widget(x_root, y_root, self):
            return
            
        # Verificar si el scroll fue dentro del listbox
        if self.is_point_inside_widget(x_root, y_root, self.listbox):
            return
            
        # Si llegamos aquí, el scroll fue fuera del área, cerrar listbox
        self.close_listbox()

    def is_point_inside_widget(self, x_root, y_root, widget):
        """Verifica si un punto (en coordenadas root) está dentro de un widget"""
        try:
            if not widget or not widget.winfo_exists():
                return False
                
            widget_x = widget.winfo_rootx()
            widget_y = widget.winfo_rooty()
            widget_width = widget.winfo_width()
            widget_height = widget.winfo_height()
            
            return (widget_x <= x_root <= widget_x + widget_width and
                    widget_y <= y_root <= widget_y + widget_height)
        except:
            return False

    def on_key_press(self, event):
        if not self.listbox:
            return
           
        if event.keysym == "Down":
            self.navigate_list(1)
            return "break"
        elif event.keysym == "Up":
            self.navigate_list(-1)
            return "break"
        elif event.keysym == "Return":
            self.select_current()
            return "break"
        elif event.keysym == "Escape":
            self.close_listbox()
            return "break"

    def navigate_list(self, direction):
        if not self.matches:
            return
           
        # Actualizar selección actual
        self.current_selection += direction
       
        # Ajustar límites
        if self.current_selection < 0:
            self.current_selection = len(self.matches) - 1
        elif self.current_selection >= len(self.matches):
            self.current_selection = 0
           
        # Actualizar selección visual en el listbox
        self.listbox_list.selection_clear(0, "end")
        self.listbox_list.selection_set(self.current_selection)
        self.listbox_list.see(self.current_selection)

    def select_current(self):
        if self.listbox and self.current_selection >= 0:
            value = self.matches[self.current_selection]
            self.var.set(value)
            self.close_listbox()

    def on_focus_out(self, event):
        # Pequeño delay para permitir clics en el listbox
        self.after(100, self.check_focus)

    def check_focus(self):
        try:
            focused = self.focus_get()
            if focused != self.listbox_list and focused != self:
                self.close_listbox()
        except:
            self.close_listbox()

    def on_configure(self, event):
        # Reposicionar el listbox cuando el entry se mueve o cambia
        if self.listbox and event.widget == self:
            self.after(1, self.update_listbox_position)

    def start_position_monitoring(self):
        """Inicia el monitoreo de posición para detectar cambios"""
        if self.listbox:
            self.check_position_change()

    def check_position_change(self):
        """Verifica si la posición del entry ha cambiado"""
        if not self.listbox:
            return
           
        try:
            current_x = self.winfo_rootx()
            current_y = self.winfo_rooty()
           
            # Si la posición cambió, actualizar el listbox
            if self.last_x != current_x or self.last_y != current_y:
                self.last_x = current_x
                self.last_y = current_y
                self.update_listbox_position()
           
            # Continuar monitoreando
            self.position_check_id = self.after(10, self.check_position_change)
        except:
            self.close_listbox()

    def update_listbox_position(self):
        if self.listbox:
            try:
                # Obtener posición actual del entry
                x = self.winfo_rootx()
                y = self.winfo_rooty() + self.winfo_height()
               
                # Verificar que el entry esté visible en la pantalla
                if self.winfo_viewable():
                    self.listbox.geometry(f"+{x}+{y}")
                    # Actualizar las posiciones de referencia
                    self.last_x = x
                    self.last_y = y - self.winfo_height()
                else:
                    self.close_listbox()
            except:
                self.close_listbox()

    def changed(self, *args):
        # Cancelar cualquier llamada 'changed' pendiente
        if self._after_id:
            self.after_cancel(self._after_id)
            self._after_id = None

        # Programar la ejecución real de la lógica de autocompletado
        # Ajusta el tiempo (en milisegundos) según sea necesario. 50-100 ms suele ser un buen punto de partida.
        self._after_id = self.after(10, self._process_change) # Llama a un nuevo método privado

    def _process_change(self): # Nuevo método para contener la lógica original de 'changed'
        self._after_id = None # Reinicia el ID de la tarea una vez que se ejecuta
        pattern = self.var.get().lower()
        self.matches = [m for m in self.options_list if pattern in m.lower()]

        if self.matches and len(pattern) > 0:
            # Asegúrate de que el listbox exista y no intentes crearlo si ya está en proceso de ser creado
            # Volvemos a incluir esta validación para mayor robustez
            if not self.listbox or not self.listbox.winfo_exists():
                self.create_listbox()
            # Si el listbox existe pero su toplevel fue destruido, se necesitaría recrear.
            # La verificación anterior `not self.listbox.winfo_exists()` maneja este caso.
            self.update_listbox_content()
        else:
            self.close_listbox()

    def create_listbox(self):
        self.listbox = ctk.CTkToplevel(self)
        self.listbox.overrideredirect(True)
        self.listbox.wm_attributes("-topmost", True)  # Mantener siempre encima
       
        # Obtener colores del tema personalizado
        try:
            appearance_mode = ctk.get_appearance_mode()
            if appearance_mode == "Dark":
                bg_color = "#343638"  # CTkEntry fg_color modo oscuro
                fg_color = "#ffffff"  # CTkLabel text_color modo oscuro
                select_bg = "#E8A046"  # Color principal del tema
                select_fg = "#ffffff"  # CTkButton text_color modo oscuro
                frame_color = "#2b2b2b"  # CTkToplevel fg_color modo oscuro
            else:
                bg_color = "#f9f9f9"  # CTkEntry fg_color modo claro
                fg_color = "#000000"  # CTkLabel text_color modo claro
                select_bg = "#E8A046"  # Color principal del tema
                select_fg = "#000000"  # CTkButton text_color modo claro
                frame_color = "#f2f2f2"  # CTkToplevel fg_color modo claro
        except:
            # Valores por defecto usando colores del tema
            bg_color = "#343638"
            fg_color = "#ffffff"
            select_bg = "#E8A046"
            select_fg = "#ffffff"
            frame_color = "#2b2b2b"
       
        # Configurar apariencia del toplevel
        self.listbox.configure(fg_color=frame_color)
       
        # Obtener el ancho exacto del entry
        self.update_idletasks()  # Asegurar que las dimensiones estén actualizadas
        entry_width = self.winfo_width()
       
        # Posicionar el listbox
        self.update_listbox_position()
       
        # Crear el listbox con estilo apropiado
        self.listbox_list = tk.Listbox(
            self.listbox,
            height=min(5, len(self.matches)) if self.matches else 5,
            bg=bg_color,
            fg=fg_color,
            selectbackground=select_bg,
            selectforeground=select_fg,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            font=("Roboto", 10),  # Tamaño de fuente consistente con el tema
            activestyle="none",  # Elimina el estilo de activación por defecto
            width=1  # Ancho mínimo, se controlará por el contenedor
        )
        self.listbox_list.pack(fill="both", expand=True, padx=1, pady=1)
        self.listbox_list.bind("<<ListboxSelect>>", self.selection)
        self.listbox_list.bind("<Double-Button-1>", self.selection)
       
        # Configurar el ancho del toplevel para que coincida con el entry
        self.listbox.resizable(False, False)
       
        # Iniciar monitoreo de posición
        self.start_position_monitoring()
        
        # Vincular eventos globales
        self.bind_global_events()
       
        # Reset selection
        self.current_selection = -1

    def update_listbox_content(self):
        if not self.listbox_list:
            return
           
        self.listbox_list.delete(0, "end")
        for item in self.matches:
            self.listbox_list.insert("end", item)
           
        # Actualizar altura del listbox
        height = min(5, len(self.matches))
        self.listbox_list.configure(height=height)
       
        # Obtener el ancho exacto del entry
        self.update_idletasks()
        entry_width = self.winfo_width()
        listbox_height = height * 22 + 4  # 22 píxeles por línea aproximadamente + padding
       
        # Configurar el tamaño exacto del toplevel
        self.listbox.geometry(f"{self.listbox_width}x{listbox_height}")
       
        # Actualizar posición para asegurar que esté en el lugar correcto
        self.update_listbox_position()

    def selection(self, event):
        if self.listbox_list.curselection():
            index = self.listbox_list.curselection()[0]
            value = self.listbox_list.get(index)
            self.var.set(value)
        self.close_listbox()

    def close_listbox(self):
        if self.listbox:
            # Cancelar monitoreo de posición
            if self.position_check_id:
                self.after_cancel(self.position_check_id)
                self.position_check_id = None
            
            # Desvincular eventos globales
            self.unbind_global_events()
           
            self.listbox.destroy()
            self.listbox = None
            self.listbox_list = None
            self.current_selection = -1
            self.matches = []
            self.last_x = None
            self.last_y = None

    def destroy(self):
        # Cancelar monitoreo de posición antes de destruir
        if self.position_check_id:
            self.after_cancel(self.position_check_id)
            self.position_check_id = None
        
        # Desvincular eventos globales
        self.unbind_global_events()
        
        self.close_listbox()
        super().destroy()

class AddChannelWindow(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.title("Añadir canal")
        self.geometry("400x255")
        self.after(200, lambda:self.wm_iconbitmap(rpath("./resources/logoAne.ico")))  # Retraso para correcto cargue del logo
        self.grab_set()        

        self.datos = {}

        # Se configura el ancho de la segunda columna de ingreso de datos
        self.grid_columnconfigure(1, weight=1)

        # Selección de tecnología
        label_tec = ctk.CTkLabel(self, text="Tecnología del canal:")
        label_tec.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        selected_tec_var = tk.StringVar(self)
        selected_tec_var.set("Seleccione la tecnología")
        self.tec_list = ctk.CTkOptionMenu(self, values=['Analógico', 'Digital'], variable=selected_tec_var, command=self.tec_selected)
        self.tec_list.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # Selección de canal
        label_channel = ctk.CTkLabel(self, text="Numero del canal:")
        label_channel.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.channel_entry = ctk.CTkEntry(self, placeholder_text="0")
        self.channel_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Selección de servicio
        label_service = ctk.CTkLabel(self, text="Servicio:")
        label_service.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.service_list = ctk.CTkOptionMenu(self, values=['Seleccione la tecnología'], state='disabled', command=self.service_selected)
        self.service_list.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # Selección de la estación de procedencia
        label_station = ctk.CTkLabel(self, text="Estación de procedencia:")
        label_station.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        
        self.list_station = ctk.CTkOptionMenu(self, values=['Seleccione la tecnología'], state='disabled', command=self.station_selected)
        self.list_station.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        # Botón de cancelar
        self.cancel_button = ctk.CTkButton(self, text="Cancelar", command=self.cancel)
        self.cancel_button.grid(row=4, column=0, pady=20)

        # Botón de guardar
        self.save_button = ctk.CTkButton(self, text="Guardar", command=self.save)
        self.save_button.grid(row=4, column=1, pady=20)

    def tec_selected(self, event):
        tec = self.tec_list.get()

        if tec == 'Analógico':
            # Obtención de la lista de estaciones del diccionario
            station_list = []
            measurement_dictionary = self.master.datos.measurement_dictionary
            for station in measurement_dictionary.keys():
                if measurement_dictionary[station]['Analógico']:
                    station_list.append(station)
            
            station_list.append('---Otra')
            
            # Se configura la lista de servicios
            df = pd.read_excel(rpath('./src/utils/Referencias.xlsx'), sheet_name = 4)
            service_list = df.columns.tolist()
            self.service_list.configure(values=service_list, state='enabled')

            # Se configura la lista de estaciones disponibles
            self.list_station.configure(values=station_list)


        elif tec == 'Digital':
            # Obtención de la lista de estaciones del diccionario
            station_list = []
            measurement_dictionary = self.master.datos.measurement_dictionary
            for station in measurement_dictionary.keys():
                if measurement_dictionary[station]['Digital']:
                    station_list.append(station)
            
            station_list.append('---Otra')
            
            # Se configura la lista de servicios
            service_list = list(PLP_SERVICES.keys())
            self.service_list.configure(values=service_list, state='enabled')

            # Se configura la lista de estaciones disponibles
            self.list_station.configure(values=station_list)

    def service_selected(self, event):
        self.list_station.configure(state='enabled')

    def station_selected(self, event):
        selected_station = self.list_station.get()

        # Se revisa la estación seleccionada
        if selected_station == '---Otra':
            # Se redimensiona la ventana para visualizar todos los widgets
            self.geometry("400x345")

            # Se añaden label y entry con autocompletado para elegir la estación
            self.label_another_station = ctk.CTkLabel(self, text="Seleccione la estación:")
            self.label_another_station.grid(row=4, column=0, padx=10, pady=10, sticky="w")
            
            # Cargue de lista de estaciones
            if self.tec_list.get() == 'Analógico':
                service = self.service_list.get()
                df = pd.read_excel(rpath('./src/utils/Referencias.xlsx'), sheet_name = 4)
                available_stations = df[service].dropna().tolist()

            elif self.tec_list.get() == 'Digital':
                df = pd.read_excel(rpath('./src/utils/Referencias.xlsx'), sheet_name = 2)
                available_stations = df['TX_TDT'].tolist()

            # Entry con autocompletado para seleccionar la estación
            self.new_station_entry = AutocompleteEntry(self, options_list=available_stations, listbox_width=210, placeholder_text="Ingrese la estación")
            self.new_station_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

            # Selección de acimut
            self.label_acimuth = ctk.CTkLabel(self, text="Ingrese el acimut:")
            self.label_acimuth.grid(row=5, column=0, padx=10, pady=10, sticky="w")

            self.acimut_entry = ctk.CTkEntry(self, placeholder_text="0")
            self.acimut_entry.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

            # Reposicionar los botones de guardar y cancelar
            self.cancel_button.grid(row=6, column=0, pady=20)
            self.save_button.grid(row=6, column=1, pady=20)

        # Si se selecciona una estación de las que ya están en la lista
        else:
            self.geometry("400x255")
            if hasattr(self, 'new_station_entry') and hasattr(self, 'acimut_entry'):
                self.label_another_station.destroy()
                self.new_station_entry.destroy()
                self.label_acimuth.destroy()
                self.acimut_entry.destroy()

            # Reposicionar los botones de guardar y cancelar
            self.cancel_button.grid(row=4, column=0, pady=20)
            self.save_button.grid(row=4, column=1, pady=20)

    def save(self):
        # Se obtiene el tipo de tecnología
        self.datos['tecnologia'] = self.tec_list.get()

        # Se obtiene el número de canal
        try:
            channel = int(self.channel_entry.get())
            if channel in list(TV_TABLE.keys()):
                self.datos['channel'] = channel
            else:
                CTkMessagebox(title="Error", message="Por favor, ingrese un canal válido. \n El canal ingresado no coincide con la lista de canales disponibles")
                return
        except ValueError:
            CTkMessagebox(title="Error", message="El número de canal debe ser un número entero.")
            return

        # Se obtiene el servicio
        self.datos['service'] = self.service_list.get()

        # Obtención de la estación
        if hasattr(self, 'new_station_entry'):
            station = self.new_station_entry.get()
            station = station.split(sep=' - ')[0].title()
            self.datos['station'] = station
        else:
            self.datos['station'] = self.list_station.get()

        # Obtención el acimut
        if hasattr(self, 'acimut_entry'):
            try:
                acimuth = int(self.acimut_entry.get())
                if 0 <= acimuth <= 360:
                    self.datos['acimuth'] = acimuth
                else:
                    CTkMessagebox(title="Error", message="Por favor, ingrese un acimut válido. \n El acimut debe ser un número entre 0 y 360.")
                    return
            except ValueError:
                CTkMessagebox(title="Error", message="El acimut debe ser un número entero.")
                return
        else:
            self.datos['acimuth'] = self.master.datos.measurement_dictionary[self.datos['station']]['Acimuth']

        # Cierra la ventana emergente
        self.destroy()

    def cancel(self):
        self.destroy()

class DatosCompartidos:
    """Clase para almacenar todos los datos que se comparten entre ventanas"""
    def __init__(self):
        # Datos de la primera ventana
        self.object_preengeneering = None
        self.municipality = None
        self.point = 0
        self.measurement_dictionary = {}
        self.sfn_dictionary = {}
        
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
        self.dane_code = None
        
        # Datos de la quinta ventana (Rotor)
        self.rtr_instrument = None
        self.ip_rotor = None
        self.angulo_actual_rotor = 0
        
        # Datos de la sexta ventana (Formulario)
        self.site_dictionary = {}
        
        # Datos de la medición
        self.medicion_completada = False
        self.resultado_medicion = None
        
        # Tipo de medición seleccionada
        self.tipo_medicion = "banco"

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuración de la ventana principal
        self.title("Automatización de medición")
        self.geometry("510x650")

        # Configuración del logo de la ventana
        self.iconbitmap(rpath("./resources/logoAne.ico"))

        # Impedir el redimensionamiento de la pantalla
        self.resizable(False, False)

        # Inicializar datos compartidos
        self.datos = DatosCompartidos()
        
        # Crear contenedor para las ventanas
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Diccionario para almacenar las páginas
        self.frames = {}
        
        # Crear todas las páginas
        for F in (MeasurementModeWindow, LoadExcelWindow, ATVInstrumentWindow, TDTInstrumentWindow, BankInstrumentWindow, 
                  RotorWindow, SiteInfoWindow, SummaryWindow):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        # Mostrar la primera página
        self.mostrar_ventana(MeasurementModeWindow)
    
    def mostrar_ventana(self, cont):
        """Trae al frente la ventana especificada"""
        frame = self.frames[cont]
        frame.tkraise()
        # Actualizar la ventana al mostrarla
        if hasattr(frame, 'actualizar'):
            frame.actualizar()

class MeasurementModeWindow(ctk.CTkFrame):
    def __init__(self, parent, controller):
        ctk.CTkFrame.__init__(self, parent)
        self.controller = controller
        
        # Título
        titulo = ctk.CTkLabel(self, text="Seleccione el tipo de medición", 
                             font=ctk.CTkFont(size=24, weight="bold"))
        titulo.pack(pady=(50, 30), padx=10)
        
        # Descripción
        descripcion = ctk.CTkLabel(self, text="Por favor, seleccione el tipo de medición que desea realizar.",
                                  font=ctk.CTkFont(size=16))
        descripcion.pack(pady=(0, 40), padx=20)
        
        # Frame para los botones
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=20, padx=10, fill="x")
        
        # Botón Televisión
        self.btn_television = ctk.CTkButton(frame_botones, text="Televisión", 
                                         font=ctk.CTkFont(size=16),
                                         height=50,
                                         command=self.seleccionar_television)
        self.btn_television.pack(pady=15, padx=40, fill="x")
        
        # Botón Banco de mediciones
        self.btn_banco = ctk.CTkButton(frame_botones, text="Banco de mediciones", 
                                     font=ctk.CTkFont(size=16),
                                     height=50,
                                     command=self.seleccionar_banco)
        self.btn_banco.pack(pady=15, padx=40, fill="x")
        
        # Botón Televisión y banco
        self.btn_ambos = ctk.CTkButton(frame_botones, text="Televisión y banco", 
                                     font=ctk.CTkFont(size=16),
                                     height=50,
                                     command=self.seleccionar_ambos)
        self.btn_ambos.pack(pady=15, padx=40, fill="x")
    
    def seleccionar_television(self):
        """Configurar para medición de televisión y avanzar a la ventana de bienvenida"""
        self.controller.datos.tipo_medicion = "television"
        self.controller.mostrar_ventana(LoadExcelWindow)
    
    def seleccionar_banco(self):
        """Configurar para medición de banco y avanzar directamente a la ventana de bandas"""
        self.controller.datos.tipo_medicion = "banco"
        self.controller.mostrar_ventana(BankInstrumentWindow)
    
    def seleccionar_ambos(self):
        """Configurar para ambos tipos de medición y avanzar a la ventana de bienvenida"""
        self.controller.datos.tipo_medicion = "ambos"
        self.controller.mostrar_ventana(LoadExcelWindow)

class LoadExcelWindow(ctk.CTkFrame):
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

        # Frame para los botones de navegación
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver a selección
        self.btn_volver = ctk.CTkButton(frame_botones, text="Volver", 
                                   command=lambda: self.controller.mostrar_ventana(MeasurementModeWindow))
        self.btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        self.btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                      command=lambda: self.avanzar())
        self.btn_siguiente.pack(side="right", padx=10, pady=10)
        self.btn_siguiente.configure(state="normal")
    
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
                nombre_archivo = ruta_archivo.split('/')[-1]
                max_caracteres_nombre = 35 # Límite para el nombre del archivo

                if len(nombre_archivo) > max_caracteres_nombre:
                    # Acorta el nombre del archivo para evitar que la ventana se ensanche
                    # Ej: "nombre_muy_largo.xlsx" -> "nombre_muy_l...go.xlsx"
                    parte_inicial = nombre_archivo[:max_caracteres_nombre - 15]
                    parte_final = nombre_archivo[-12:]
                    nombre_mostrado = f"{parte_inicial}...{parte_final}"
                else:
                    nombre_mostrado = nombre_archivo
                self.lbl_archivo.configure(text=f"Archivo: {nombre_mostrado}")
                
                # Poblar las listas desplegables con las columnas del Excel
                municipalities = self.controller.datos.object_preengeneering.get_municipalities()
                
                # Habilitar y actualizar primera lista
                self.lista1.configure(values=municipalities, state="normal")
                self.lista1.set('Seleccione un municipio')

                # Desabilita el botón siguiente hasta que se elijan municipio y punto
                self.btn_siguiente.configure(state="disabled")
                
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
        
        # Obtener diccionarios
        diccionario_medicion = self.controller.datos.object_preengeneering.get_dictionary(self.municipality, self.punto)
        self.controller.datos.measurement_dictionary = diccionario_medicion
        diccionario_sfn = self.controller.datos.object_preengeneering.get_sfn(diccionario_medicion)
        self.controller.datos.sfn_dictionary = diccionario_sfn

        # Se crea todo el frame para mostrar las estaciones
        self.actualizar_estaciones(diccionario_medicion, diccionario_sfn)

    def actualizar_estaciones(self, diccionario_medicion: dict, diccionario_sfn: dict):
        # Eliminar el frame anterior si existe
        if hasattr(self, 'frame_estaciones_container'):
            self.frame_estaciones_container.destroy()

        # Crear contenedor principal con frame desplazable
        self.frame_estaciones_container = ctk.CTkFrame(self)
        self.frame_estaciones_container.pack(pady=(20,0), padx=10, fill="both", expand=True)
        
        # Crear el frame desplazable
        self.frame_scroll = ctk.CTkScrollableFrame(self.frame_estaciones_container)
        self.frame_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Título de la sección
        titulo_estaciones = ctk.CTkLabel(self.frame_scroll, text="Resumen de la medición",
                                    font=ctk.CTkFont(size=16, weight="bold"))
        titulo_estaciones.pack(pady=(10,15), padx=10)
        
        # Iterar sobre cada estación para mostrar su información
        for estacion, datos in diccionario_medicion.items():
            # Frame para cada estación
            frame_estacion = ctk.CTkFrame(self.frame_scroll)
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
            separador = ctk.CTkFrame(self.frame_scroll, height=2)
            separador.pack(fill="x", padx=10, pady=(15,0))
            
            # Título de la sección SFN
            titulo_sfn = ctk.CTkLabel(self.frame_scroll, text="Canales en SFN",
                                    font=ctk.CTkFont(size=16, weight="bold"))
            titulo_sfn.pack(pady=(15,10), padx=10)
            
            # Frame para canales SFN
            frame_sfn = ctk.CTkFrame(self.frame_scroll)
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

        # Botón "Agregar canal"
        self.btn_agregar_canal = ctk.CTkButton(self.frame_scroll, text="Agregar canal", command=self.add_channel)
        self.btn_agregar_canal.pack(pady=10, padx=10)

    def add_channel(self):
        """Función para agregar un nuevo canal, muestra los campos de entrada"""
        # Crear la ventana emergente 
        ventana_add = AddChannelWindow(master=self.controller)
        self.wait_window(ventana_add)

        # Se obtiene el diccionario de la nueva estación cargada
        new_station_dictionary = ventana_add.datos

        # Se reprocesan los diccionarios
        diccionario_medicion = self.controller.datos.measurement_dictionary
        diccionario_medicion = self.controller.datos.object_preengeneering.add_station(diccionario_medicion, new_station_dictionary)
        self.controller.datos.measurement_dictionary = diccionario_medicion

        diccionario_sfn = self.controller.datos.object_preengeneering.get_sfn(diccionario_medicion)
        self.controller.datos.sfn_dictionary = diccionario_sfn

        self.actualizar_estaciones(diccionario_medicion, diccionario_sfn)
    
    def avanzar(self):
        """Guardar selecciones y avanzar a la siguiente ventana"""
        # Guardar selecciones en datos compartidos
        self.controller.datos.municipality = self.lista1.get()
        self.controller.datos.point = self.lista2.get()
        
        # Avanzar a la siguiente ventana
        self.controller.mostrar_ventana(ATVInstrumentWindow)

class ATVInstrumentWindow(ctk.CTkFrame):
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
                               command=lambda: controller.mostrar_ventana(LoadExcelWindow))
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
        self.controller.mostrar_ventana(TDTInstrumentWindow)

class TDTInstrumentWindow(ctk.CTkFrame):
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
                               command=lambda: controller.mostrar_ventana(ATVInstrumentWindow))
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
        if self.controller.datos.tipo_medicion == "banco" or self.controller.datos.tipo_medicion == "ambos":
            self.controller.mostrar_ventana(BankInstrumentWindow)
        else:
            self.controller.mostrar_ventana(RotorWindow)

class BankInstrumentWindow(ctk.CTkFrame):
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
        
        self.rb_instrumento_fph = ctk.CTkRadioButton(self.frame_instrumento, text="FPH/FSH", 
                                             variable=self.var_tipo_instrumento,
                                             value="FPH")
        self.rb_instrumento_fph.pack(padx=30, pady=5, anchor="w")

        self.rb_instrumento_viavi = ctk.CTkRadioButton(self.frame_instrumento, text="Viavi", 
                                             variable=self.var_tipo_instrumento,
                                             value="Viavi")
        self.rb_instrumento_viavi.pack(padx=30, pady=5, anchor="w")
        
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
        
        # Crear siempre el frame para municipio, pero no mostrarlo aún
        # Frame para el municipio
        self.frame_municipio = ctk.CTkFrame(self.main_container)
        # No hacemos pack aún
        
        lbl_municipio = ctk.CTkLabel(self.frame_municipio, text="Nombre del municipio:")
        lbl_municipio.pack(padx=10, pady=5, anchor="w")
        
        # self.entry_municipio = ctk.CTkEntry(self.frame_municipio, placeholder_text="Ingrese el municipio")
        self.df = pd.read_excel(rpath('./src/utils/Referencias.xlsx'), sheet_name = 3)
        municipalities_list = self.df['Municipio - departamento'].tolist()

        self.entry_municipio = AutocompleteEntry(self.frame_municipio, options_list=municipalities_list, listbox_width=410, placeholder_text="Ingrese el municipio")
        self.entry_municipio.pack(padx=10, pady=5, fill="x")
        
        # Frame para el código DANE
        # self.frame_dane = ctk.CTkFrame(self.main_container)
        # No hacemos pack aún
        
        # lbl_dane = ctk.CTkLabel(self.frame_dane, text="Código DANE:")
        # lbl_dane.pack(padx=10, pady=5, anchor="w")
        
        # self.entry_dane = ctk.CTkEntry(self.frame_dane, placeholder_text="Ingrese el código DANE (números)")
        # self.entry_dane.pack(padx=10, pady=5, fill="x")
        
        # Frame para los botones de navegación - se coloca directamente en el self
        # para que quede siempre en la parte inferior
        frame_botones = ctk.CTkFrame(self)
        frame_botones.pack(pady=10, padx=10, fill="x", side="bottom")
        
        # Botón para volver
        btn_volver = ctk.CTkButton(frame_botones, text="Anterior", 
                               command=lambda: controller.mostrar_ventana(MeasurementModeWindow) if self.controller.datos.tipo_medicion == "banco" else controller.mostrar_ventana(ATVInstrumentWindow))
        btn_volver.pack(side="left", padx=10, pady=10)
        
        # Botón para avanzar
        btn_siguiente = ctk.CTkButton(frame_botones, text="Siguiente", 
                                  command=self.avanzar)
        btn_siguiente.pack(side="right", padx=10, pady=10)
        
        # Inicializar el estado de los widgets
        self.actualizar_estado_widgets()
    
    def actualizar(self):
        """Se llama cuando la ventana se muestra"""        
        # Mostrar u ocultar frames según el tipo de medición
        if self.controller.datos.tipo_medicion == 'banco':
            # Mostrar los frames de municipio y DANE
            self.frame_municipio.pack(after=self.frame_conexion, pady=10, padx=10, fill="x")
            # self.frame_dane.pack(after=self.frame_municipio, pady=10, padx=10, fill="x")
        else:
            # Ocultar los frames
            self.frame_municipio.pack_forget()
            # self.frame_dane.pack_forget()
    
    def actualizar_estado_widgets(self):
        """Habilitar o deshabilitar widgets según la selección de realizar mediciones"""
        estado = "normal" 
        
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
                except:
                    conexion_exitosa = False
                    instrument_model_name = ''

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

            elif tipo_instrumento == 'Viavi':
                # Código para conectar con el dispositivo
                try:
                    # Creación del objeto de conexión
                    mbk_instrument = ViaviManager(ip, impedance, [])

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
        
        # Rehabilitar el botón
        self.btn_conectar.configure(state="normal")
    
    def avanzar(self):
        """Guardar datos y avanzar a la siguiente ventana"""
        # Si tenemos campos de municipio y DANE, validar y guardar
        if self.controller.datos.tipo_medicion == 'banco':
            # Validar municipio (solo letras)
            municipio_departamento = self.entry_municipio.get()
            if not municipio_departamento:
                CTkMessagebox(title="Error", message="Ingrese el nombre del municipio.")
                return
                
            # Guardar los valores en el controlador
            municipio = municipio_departamento.split(" - ")[0].strip().upper()
            dane_code_index = self.df.index[self.df['Municipio - departamento'] == municipio_departamento]
            dane_code = str(self.df.at[dane_code_index[0], 'DANE']).zfill(5)
            self.controller.datos.municipality = municipio
            self.controller.datos.dane_code = dane_code
            
        # Avanzar a la siguiente ventana
        if self.controller.datos.tipo_medicion == "television" or self.controller.datos.tipo_medicion == "ambos":
            self.controller.mostrar_ventana(RotorWindow)
        else:
            self.controller.mostrar_ventana(SummaryWindow)
        
class RotorWindow(ctk.CTkFrame):
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
                               command=lambda: controller.mostrar_ventana(BankInstrumentWindow) if self.controller.datos.tipo_medicion != "television" else controller.mostrar_ventana(TDTInstrumentWindow))
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
        self.controller.mostrar_ventana(SiteInfoWindow)

class SiteInfoWindow(ctk.CTkFrame):
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
                               command=lambda: controller.mostrar_ventana(RotorWindow))
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
        self.controller.mostrar_ventana(SummaryWindow)

class SummaryWindow(ctk.CTkFrame):
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
                               command=lambda: controller.mostrar_ventana(SiteInfoWindow) if self.controller.datos.tipo_medicion != "banco" else controller.mostrar_ventana(BankInstrumentWindow))
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
        
        # Agregar sección de municipio
        municipality_dict = {
            "Municipio": datos.municipality.title(),
        }

        if datos.tipo_medicion == 'television' or datos.tipo_medicion == 'ambos':
            municipality_dict["Punto"] = datos.point

        if datos.tipo_medicion == 'banco' or datos.tipo_medicion == 'ambos':
            municipality_dict["DANE"] = datos.dane_code

        agregar_seccion("Municipio", municipality_dict)
        
        # Agregar sección de TV Analógica
        if datos.atv_instrument is not None:
            tv_analogica_dict = {
                "Instrumento": datos.atv_instrument.instrument_model_name,
                "IP": datos.atv_instrument.ip_address,
                "Puerto": f"{datos.atv_instrument.impedance}Ω",
                "Transductores": datos.atv_instrument.transducers,
            }
            agregar_seccion("TV Analógica", tv_analogica_dict)
        
        # Agregar sección de TV Digital
        if datos.dtv_instrument is not None:
            tv_digital_dict = {
                "Instrumento": datos.dtv_instrument.instrument_model_name,
                "IP": datos.dtv_instrument.ip_address,
                "Puerto": f"{datos.dtv_instrument.impedance}Ω",
                "Transductores": datos.dtv_instrument.transducers,
            }
            agregar_seccion("TV Digital", tv_digital_dict)
        
        # Agregar sección de Bandas
        if datos.mbk_instrument is not None:
            bandas_dict = {
                "Instrumento": datos.mbk_instrument.instrument_model_name,
                "IP": datos.mbk_instrument.ip_address,
                "Puerto": f"{datos.mbk_instrument.impedance}Ω",
                "Transductores": datos.mbk_instrument.transducers,
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
        def manual_rotation_callback(station):
            msg = CTkMessagebox(
                title="Rotación Manual", 
                message=f"Gire el rotor hacia {station}.\nUna vez apuntado, haga click en aceptar.",
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
        
        # Mostrar botón de finalizar (en el hilo principal)
        def show_finish_button():
            self.controller.datos.medicion_completada = True
            self.btn_finalizar.pack(side="right", padx=10, pady=10)

        # Iniciar proceso de medición en un hilo separado
        def start_measurement(self):
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
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
                root_path = filedialog.askdirectory(
                    title="Seleccione la carpeta para guardar resultados",
                    mustexist=False
                )
                storage_path_tv = rpath(f"{root_path}/{municipality}/P{str(point).zfill(2)}")

                # Verificar si mbk no es None
                if measurement_manager.mbk is not None:
                    # Definición de la ruta de almacenamiento de soportes para tv
                    date = self.controller.datos.site_dictionary['date_for_mbk_folder']

                    # Si solo se va a medir banco, se trae el código DANE desde el usuario
                    if self.controller.datos.tipo_medicion == 'banco':
                        dane_code = self.controller.datos.dane_code
                        storage_path_bank = rpath(f"{root_path}/{dane_code}_{date}_{municipality.replace(' ', '-').upper()}")
                        
                    # En caso contrario, se obtiene desde la preingeniería
                    else:
                        dane_code = self.controller.datos.object_preengeneering.get_dane_code(municipality)
                        storage_path_bank = rpath(f"{root_path}/{dane_code}_{date}_{municipality.replace(' ', '-').upper()}_P{point}")

                # Actualizar progreso
                progress_callback(0, 1, "Realizando mediciones SFN...")

                # Hacer medición de SFN en caso de que sea necesario
                if measurement_manager.atv is not None and measurement_manager.dtv is not None:

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
                    try:
                        report = ExcelReport()
                        report.fill_reports(
                            site_dictionary = self.controller.datos.site_dictionary,
                            analog_measurement_dictionary = atv_result,
                            digital_measurement_dictionary = dtv_result,
                            sfn_dictionary = self.controller.datos.sfn_dictionary
                        )
                    except:
                        progress_callback(1, 1, "Error al generar reportes.")
                        self.after(0, show_finish_button)

                # Ejecutar mbk_measurement secuencialmente si no es paralelo
                if measurement_manager.mbk is not None:
                    progress_callback(0.95, 1, "Iniciando medición de banco...")
                    msg = CTkMessagebox(title="Iniciando medición de banco de mediciones",
                                  message="Conecte la antena para medir banco al instrumento seleccionado.\n Cuando esté conectada, de clic en el botón.")
                    msg.get()
                    measurement_manager.mbk_measurement(storage_path_bank, progress_callback)

                # Completar medición
                progress_callback(1, 1, "¡Medición completada con éxito!")
                
                self.after(0, show_finish_button)
                
            except Exception:
                # Manejar errores y mostrarlos al usuario
                error_msg = f"Error durante la medición."
                self.lbl_estado_medicion.configure(text=error_msg)
                CTkMessagebox(title="Error", message=error_msg, icon="error")
                # Volver a mostrar el botón de inicio
                self.btn_iniciar.pack(side="right", padx=10, pady=10)

            finally:
                pythoncom.CoUninitialize()
        
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
        self.controller.mostrar_ventana(LoadExcelWindow)

# Función para iniciar la aplicación
def iniciar_aplicacion():
    app = MainWindow()
    app.mainloop()

if __name__ == "__main__":
    iniciar_aplicacion()