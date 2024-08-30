import tkinter as tk
from tkinter import ttk
import threading
import pythoncom  # Importa pythoncom para manejar la inicialización de COM
import informes
import os


class InformesApp(tk.LabelFrame):
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
            "CIUDAD": "PASTO",
            "DIRECCION": "Calle 14 No. 28-45 Sector Bomboná"
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
        self.variable_de_ruta.set(os.path.join(self.script_directory,'..'))
        self.ruta=tk.Entry(self,width=115,textvariable=self.variable_de_ruta,state='readonly')
        self.ruta.grid(row=1,column=0,padx=(40,10), pady=(0,60),sticky='nsew')

        self.botonRuta=ttk.Button(self,text="Examinar",width=25,command=self.buscar_ruta_de_trabajo)
        self.botonRuta.grid(row=1,column=1,padx=(10,40), pady=(0,60),sticky='w')

        
        

        # Crear el LabelFrame
        self.labelframe = tk.LabelFrame(self, text="Información")
        self.labelframe.grid(row=2,column=0,padx=(40,10), pady=(10,60),sticky='nsew')
        self.labelframe.grid_columnconfigure(0,weight=1)
        self.labelframe.grid_columnconfigure(1,weight=1)


        # Crear los labels y entries
        self.entries = {}
        for idx, (label_text, default_value) in enumerate(self.default_values.items()):
            label = ttk.Label(self.labelframe, text=label_text)
            label.grid(row=idx, column=0, padx=5, pady=15, sticky=tk.W)
            
            entry = ttk.Entry(self.labelframe,width=100)
            entry.insert(0, default_value)
            entry.grid(row=idx, column=1, padx=5, pady=15)
            
            self.entries[label_text] = entry

        #Boton frame
        self.botonFrame=tk.LabelFrame(self)
        self.botonFrame.grid(row=2,column=1,padx=(10,40), pady=(10,60),sticky='nsew')

        self.botonFrame.grid_columnconfigure(0,weight=1)
        self.botonFrame.grid_rowconfigure(1,weight=1)
        self.botonFrame.grid_rowconfigure(0,weight=1)
        # Crear el botón
        self.boton = ttk.Button(self.botonFrame, text="Generar Informe",
                                command=self.ejecutar_automatizacion,width=20)
        self.boton.grid(row=0, column=0, pady=(25,1),padx=10,sticky='ew',ipady=3)

        self.botondos = ttk.Button(self.botonFrame, text="Unir fotos",width=20,
                                command=lambda x=self.script_directory:informes.unir_informe_con_fotos(x))
        self.botondos.grid(row=1, column=0, pady=(1,25),padx=10,sticky='ew',ipady=3)

        self.botonCancelar = ttk.Button(self.botonFrame, text="Cancelar",
                                command=self.controlador.show_home,width=20)
        self.botonCancelar.grid(row=2, column=0, pady=(1,25),padx=10,sticky='ew',ipady=3)


    def ejecutar_automatizacion(self):
        # Obtener los valores actuales de los entries
        cliente = self.entries["CLIENTE"].get()
        contrato = self.entries["CONTRATO"].get()
        orden = self.entries["ORDEN"].get()
        ciudad = self.entries["CIUDAD"].get()
        direccion = self.entries["DIRECCION"].get()

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
        self.progreso_label = ttk.Label(self.progreso_ventana, text="Generando informes, por favor espere...")
        self.progreso_label.pack(pady=10)
        self.barra_progreso = ttk.Progressbar(self.progreso_ventana, mode='indeterminate')
        self.barra_progreso.pack(pady=10)
        self.barra_progreso.start()

        # Crear un hilo para ejecutar la función
        thread = threading.Thread(target=self._run_automatizacion, args=(cliente, orden, contrato, direccion, ciudad))
        thread.start()

    def _run_automatizacion(self, cliente, orden, contrato, direccion, ciudad):
        # Inicializar COM en este hilo
        pythoncom.CoInitialize()
        try:
            # Ejecutar la función del módulo main
            informes.ejecutar_automatizacion_informes(cliente, orden, contrato, direccion, ciudad,
                                                      self.script_directory,self.variable_de_ruta.get())
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