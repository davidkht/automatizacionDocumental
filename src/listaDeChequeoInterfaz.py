import tkinter as tk
from tkinter import ttk
import threading
import pythoncom  # Importa pythoncom para manejar la inicialización de COM
import listaDeChequeo
    

class listaCApp(tk.LabelFrame):
    def __init__(self,parent,textoLabel,controlador,directorio_de_script):
        super().__init__(parent)
        self.config(text=textoLabel)
        self.controlador=controlador
        self.script_directory=directorio_de_script

        # Valores por defecto
        self.default_values = {
            "CLIENTE": "Universidad Nacional Abierta y a Distancia",
            "CONTRATO": "1782",
            "ORDEN": "CMM-2023-000130",
            "COMERCIAL":"LUIS CAÑÓN",
            "RUTA DE ALMACENAMIENTO" : r'\\172.16.0.9\Depto tecnico\2 0 2 3\2. CONTRATOS\1782 UNAD DOTACIÓN\2. LISTAS DE CHEQUEO'
        }

        #Ruta de Trabajo
        # Configura las columnas del contenedor principal
        self.grid_columnconfigure(0, weight=0)  # Columna del label
        self.grid_columnconfigure(1, weight=1)  # Columna del Entry
        self.grid_columnconfigure(2, weight=0)  # Columna del botón
        # Configura la fila del LabelFrame para que se expanda
        self.grid_rowconfigure(0, weight=1)

        self.labelruta=ttk.Label(self,text="Ruta de Trabajo: ")
        self.labelruta.grid(row=0,column=0,padx=(40,10),pady=(40,0),sticky='w')

        self.variable_de_ruta=tk.StringVar()
        self.variable_de_ruta.set(r"")
        self.ruta=tk.Entry(self,width=115,textvariable=self.variable_de_ruta,state='readonly')
        self.ruta.grid(row=1,column=0,padx=(40,5), pady=(0,60),sticky='nsew')

        self.botonRuta=ttk.Button(self,text="Examinar",width=25)
        self.botonRuta.grid(row=1,column=1,padx=(5,40), pady=(0,60),sticky='w',
                            command=self.buscar_ruta_de_trabajo)



        # Crear el LabelFrame
        self.labelframe = tk.LabelFrame(self, text="Información")
        self.labelframe.grid(row=2,column=0,columnspan=2,padx=40, pady=10,sticky='nsew')
        self.labelframe.grid_columnconfigure(0,weight=1)
        self.labelframe.grid_columnconfigure(1,weight=1)


        # Crear los labels y entries
        self.entries = {}
        for idx, (label_text, default_value) in enumerate(self.default_values.items()):
            label = ttk.Label(self.labelframe, text=label_text)
            label.grid(row=idx, column=0, padx=5, pady=15, sticky='w')
            self.labelframe.grid_rowconfigure(idx,weight=1)
            entry = ttk.Entry(self.labelframe,width=105)
            entry.insert(0, default_value)
            entry.grid(row=idx, column=1, padx=5, pady=15,sticky='w')
            
            self.entries[label_text] = entry

        #Boton frame
        self.botonFrame=tk.LabelFrame(self)
        self.botonFrame.grid(row=3,column=0,columnspan=2,padx=40, pady=(20,40),sticky='nsew')
        # Crear el botón
        self.boton = ttk.Button(self.botonFrame, text="Generar Lista de Chequeo",
                                command=self.ejecutar_automatizacion,width=30)
        self.boton.grid(row=0, column=0, columnspan=1,pady=10,padx=(50,50),sticky='nsew',ipady=4)

        self.botondos = ttk.Button(self.botonFrame, text="Unir fotos",
                                command=lambda x=self.script_directory:listaDeChequeo.unir_informe_con_fotos(x),
                                width=30)
        self.botondos.grid(row=0, column=1,columnspan=1, pady=10,padx=(50,50),sticky='nsew',ipady=4)

        self.cancelarButton=ttk.Button(self.botonFrame,text="Cancelar",
                                       command=self.controlador.show_home,width=30)
        self.cancelarButton.grid(row=0, column=2,columnspan=1, pady=10,padx=(50,50),sticky='nsew',ipady=4)
        # self.botonFrame.columnconfigure()

    def ejecutar_automatizacion(self):
        # Obtener los valores actuales de los entries
        cliente = self.entries["CLIENTE"].get()
        contrato = self.entries["CONTRATO"].get()
        orden = self.entries["ORDEN"].get()
        comercial = self.entries["COMERCIAL"].get()
        rutaAlm = self.entries["RUTA DE ALMACENAMIENTO"].get()

        # Inhabilitar el botón
        self.boton.config(state=tk.DISABLED)
        
        # Crear y mostrar ventana de progreso
        self.progreso_ventana = tk.Toplevel(self)
        self.progreso_ventana.title("Procesando...")
        self.progreso_ventana.geometry("300x100")
        # Calcular la posición para centrar la ventana de progreso sobre la ventana principal
        self.update_idletasks()  # Asegurarse de que todas las tareas pendientes se actualicen
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()
        prog_width = 300
        prog_height = 100
        pos_x = main_x + (main_width // 2) - (prog_width // 2)
        pos_y = main_y + (main_height // 2) - (prog_height // 2)

        self.progreso_ventana.geometry(f"{prog_width}x{prog_height}+{pos_x}+{pos_y}")
        self.progreso_label = ttk.Label(self.progreso_ventana, text="Generando LC, por favor espere...")
        self.progreso_label.pack(pady=10)
        self.barra_progreso = ttk.Progressbar(self.progreso_ventana, mode='indeterminate')
        self.barra_progreso.pack(pady=10)
        self.barra_progreso.start()

        # Crear un hilo para ejecutar la función
        thread = threading.Thread(target=self._run_automatizacion, args=(cliente, orden, contrato, rutaAlm, comercial))
        thread.start()

    def _run_automatizacion(self, cliente, orden, contrato, ruta, comercial):
        # Inicializar COM en este hilo
        pythoncom.CoInitialize()
        try:
            # Ejecutar la función del módulo main
            listaDeChequeo.ejecutar_automatizacion_listasC(cliente, orden, contrato, ruta, comercial,self.script_directory)
        finally:
            # Asegurarse de desinicializar COM después de completar la tarea
            pythoncom.CoUninitialize()
        
        # Simular tiempo de ejecución para prueba
        # time.sleep(20) 

        # Una vez completada la ejecución, actualizar la UI en el hilo principal
        self.after(0, self._on_automatizacion_complete)

    def _on_automatizacion_complete(self):
        # Detener y cerrar la ventana de progreso
        self.barra_progreso.stop()
        self.progreso_ventana.destroy()
        
        # Habilitar el botón nuevamente
        self.boton.config(state=tk.NORMAL)

    def buscar_ruta_de_trabajo(self):
        ruta_valida=self.controlador.examinar_buscar_ruta()
        if ruta_valida==0:
            pass
        else:
            self.variable_de_ruta.set(ruta_valida)


