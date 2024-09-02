import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import sp
import os
import sys
import pandas as pd


class CrearSPWindow(tk.Toplevel):
    def __init__(self, master, datos, fsc, carpetas, carpeta_mitad,quantities,directorio,ruta_de_trabajo):
        super().__init__(master)

        self.title(datos['Carpeta'])
        # self.iconbitmap(os.path.join(script_directory,"imagen.ico"))
        imagen_ico=Image.open(os.path.join(directorio,'..','img','imagen.ico'))
        self.mi_imagen= imagen_ico.resize((48,48))
        self.mi_imagen = ImageTk.PhotoImage(self.mi_imagen)
        self.datos = datos
        self.carpetas = carpetas
        self.carpeta_mitad = carpeta_mitad
        self.cantidades=quantities
        self.directorio=directorio
        self.fsc=fsc
        self.ruta_trabajo=ruta_de_trabajo
        self.quantities_final=None

        self.referencias = sp.extraer_referencias_de_base_de_datos(self.referencias_seleccionadas(),self.directorio)
        self.frameIzquierdo = ttk.Frame(self)
        self.frameIzquierdo.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        self.frameDerecho = ttk.Frame(self)
        self.frameDerecho.grid(row=0, column=1, padx=(10, 20))

        self.crear_widgets()
        self.bind_events()

        
 
    def crear_widgets(self):
        self.tree = ttk.Treeview(self.frameIzquierdo)
        self.tree.grid(row=0, column=0, sticky='nsew')
        # Configurar las columnas en el Treeview
        columnas_a_mostrar = ['DESCRIPCION', 'CANTIDAD', 'MONEDA', 'PRECIO']
        columnas_mostrar_treeview = ['REFERENCIA'] + columnas_a_mostrar
        # Ocultar la columna de árbol (tree) para no tener una columna vacía al inicio
        self.tree['columns'] = columnas_mostrar_treeview
        self.tree['show'] = 'headings'# Esto hace que solo se muestren las columnas definidas, sin la columna de árbol

       # Configurar las columnas en el Treeview
        for columna in columnas_mostrar_treeview:
            self.tree.heading(columna, text=columna)
            self.tree.column(columna, anchor=tk.CENTER)

        self.boton_guardar = ttk.Button(self.frameDerecho, text='Guardar Cantidades', command=self.click_cantidades)
        self.boton_guardar.grid(row=3, column=0, sticky='ew', padx=20)

        self.botonFinal = ttk.Button(self.frameDerecho, image=self.mi_imagen, command=self.click_final)
        self.botonFinal.grid(row=5, column=0, sticky='s', pady=(50, 15))

        self.pdfTrue=sp.crear_carpeta_y_archivos(self.datos['Carpeta'],self.fsc,self.carpeta_mitad,self.directorio,self.ruta_trabajo)
        if self.pdfTrue:
            label = ttk.Label(self.frameDerecho, text="Carpeta creada exitósamente",style="Bold.TLabel")
            label.grid(row=0, column=0,pady=20)
        else:
            label = ttk.Label(self.frameDerecho, text="Carpeta creada sin FSC\nni información de comercial",style="Bold.TLabel")
            label.grid(row=0, column=0,pady=20)

        labelInfo=ttk.Label(self.frameDerecho, text="Porfavor confirme cantidades y guardelas.",style="Large.TLabel")
        labelInfo.grid(row=1, column=0)
        labelInfo2=ttk.Label(self.frameDerecho, text="Finalice con el botón 'Electro'",style="Large.TLabel")
        labelInfo2.grid(row=2, column=0,pady=(0,30))
        # Contador de ocurrencias de referencias
        referencia_count = {}

        for indice, fila in self.referencias.iterrows():
            referencia = indice
            descripcion = fila['DESCRIPCION']
            referencia_completa = f"{referencia} - {descripcion}"
            
            # Contar ocurrencias de cada referencia
            if referencia_completa not in referencia_count:
                referencia_count[referencia_completa] = 0
            else:
                referencia_count[referencia_completa] += 1
            
            # Obtener la cantidad correcta de la lista de tuplas
            cantidad = ''
            occurrences = 0
            for ref, qty in self.cantidades:
                if ref.startswith(referencia_completa):
                    if occurrences == referencia_count[referencia_completa]:
                        cantidad = qty
                        break
                    occurrences += 1
            
            valores = [referencia, descripcion, cantidad, fila['MONEDA'], fila['PRECIO']]
            self.tree.insert("", tk.END, values=valores)

        # Configurar las cabeceras y el ancho de las columnas
        self.tree.column('REFERENCIA', width=80)
        self.tree.column('DESCRIPCION', width=500)
        self.tree.column('CANTIDAD', width=75)
        self.tree.column('MONEDA',  width=60)
        self.tree.column('PRECIO',  width=80)

        self.ajustar_altura_treeview()

    def bind_events(self):
        self.tree.bind("<Double-1>", self.on_double_click)

    def referencias_seleccionadas(self):
        total_elementos = self.master.selected_listbox.size()
        elementos = self.master.selected_listbox.get(0, total_elementos)
        referencias = [elemento.split(" - ")[0] for elemento in elementos]
        return referencias

    def on_double_click(self, event):
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        if self.tree.heading(column)['text'] == 'CANTIDAD':
            self.entry_popup(item, column)
    
    def entry_popup(self, item, column):
        x, y, width, height = self.tree.bbox(item, column)
        entry = tk.Entry(self.tree, width=width)
        entry.place(x=x, y=y, width=width, height=height)

        def save_edit(event):
            value = entry.get()
            if value.strip() and value.replace('.', '', 1).isdigit(): # Validate that the value is a number and not empty
                self.tree.set(item, column=self.tree.heading(column)['text'], value=value)
                entry.destroy()
            else:
                messagebox.showerror("Entrada inválida", "Porfavor ingrese un número válido", parent=self)

        entry.bind("<Return>", save_edit)
        entry.focus()

    def guardar_cantidades(self):
        cantidades = []
        for item in self.tree.get_children():
            cantidad = self.tree.item(item, 'values')[2]
            if cantidad.strip() and cantidad.replace('.', '', 1).isdigit():# Validate the quantity is not empty and numeric
                cantidades.append(cantidad)
            else:
                messagebox.showerror("Error", f"Cantidad inválida para la referencia {self.tree.item(item, 'values')[0]}", parent=self)
                return None
        return cantidades

    def click_cantidades(self):
        try:
            if self.master.selected_listbox.size() == 0:
                self.cantidades = True
                messagebox.showinfo("Información", "No hay elementos. Presione el botón 'ELECTRO'", parent=self)
            else:
                self.quantities_final = self.guardar_cantidades()
                if self.quantities_final is not None:
                    messagebox.showinfo("Información", "Cantidades guardadas", parent=self)
        except Exception as e:
            messagebox.showerror("Error", "ERROR", parent=self)

    def click_final(self):
        try:
            if self.quantities_final and all(cantidad.strip() for cantidad in self.quantities_final):
                try:
                    self.quantities_final = [float(cantidad) for cantidad in self.quantities_final if cantidad.strip()]
                    sp.manejar_SP(self.datos, self.referencias, self.quantities_final, self.carpeta_mitad,self.ruta_trabajo)
                    sp.crear_csv_cot(os.path.join(self.ruta_trabajo, self.carpeta_mitad, self.datos['Carpeta']))
                    messagebox.showinfo("Éxito", "Solicitud creada exitósamente\nPresione Aceptar para salir.", parent=self)
                    self.destroy()
                except Exception as e:
                    messagebox.showerror("Error", str(e), parent=self)
            else:
                messagebox.showwarning("Advertencia", "No se olvide de GUARDAR las cantidades!", parent=self)
        except NameError:
            messagebox.showwarning("Advertencia", "No se olvide de guardar las cantidades!", parent=self)
        except TypeError:
            messagebox.showwarning("Advertencia", "La solicitud se guardará sin ítems", parent=self)
            try:
                sp.manejar_SP(self.datos, self.referencias, self.quantities_final, self.carpeta_mitad,self.ruta_trabajo)
                sp.crear_csv_cot(os.path.join(self.ruta_trabajo, self.carpeta_mitad, self.datos['Carpeta']))
                messagebox.showinfo("Éxito", "Solicitud creada exitósamente\nPresione Aceptar para salir.", parent=self)
                self.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(Exception), parent=self)

    def ajustar_altura_treeview(self, min_height=15, max_height=30):
        num_items = len(self.tree.get_children())
        new_height = min(max(num_items, min_height), max_height)
        self.tree.config(height=new_height)

class SPApp(tk.LabelFrame):
    def __init__(self,parent,textoLabel,controlador,directorio):
        super().__init__(parent)
        self.config(text=textoLabel)
        self.controlador=controlador
        self.quantities = {}
        self.directorio=directorio
        self.carpetas={
            'PHY':'PHYWE',
            'ELECTRO':'ELECTRO',
            '3B':'3B',
            'LN':'LN',
            'TER':'TERCEROS',
            'EU':'EUROMEX',
            'PT': 'PT'
        }

        self.vcmd = (self.register(self.on_validate), '%P')
        self.left_frame_creation()
        self.right_frame_creation()
        self.oferta_frame()
        self.gastos_frame()
        self.creacion_frame()
        self.ruta_frame()
        

    def ruta_frame(self):
        self.frame_de_ruta=tk.LabelFrame(self,text='')
        self.frame_de_ruta.grid(row=0,column=1,padx=(10,0),pady=(10,10),sticky='nsew')

        labelruta=ttk.Label(self.frame_de_ruta,text="Ruta de Trabajo: ")
        labelruta.grid(row=0,column=0,padx=(10,10),pady=(0,5),sticky='w')

        self.variable_de_ruta=tk.StringVar()
        self.variable_de_ruta.set(r"\\172.16.0.9\Depto tecnico\2 0 2 4\1. COTIZACIONES")
        ruta=tk.Entry(self.frame_de_ruta,width=56,textvariable=self.variable_de_ruta,state='readonly')
        ruta.grid(row=1,column=0,padx=(10,5), pady=(0,10),sticky='ew')

        botonRuta=ttk.Button(self.frame_de_ruta,text="Examinar",width=20,
                                  command=self.buscar_ruta_de_trabajo)
        botonRuta.grid(row=2,column=0,padx=(10,10), pady=(0,12),sticky='w',ipady=3)

    def left_frame_creation(self):

        # Frame for the left side inputs
        self.left_frame = tk.LabelFrame(self,text='Información Encabezado')
        self.left_frame.grid(row=0, column=0, rowspan=2,sticky="nwes", padx=(10,0), pady=(5,5))
        self.left_frame.grid_rowconfigure(0,weight=1)
        self.left_frame.grid_rowconfigure(1,weight=1)
        self.left_frame.grid_rowconfigure(2,weight=1)
        self.left_frame.grid_rowconfigure(3,weight=1)
        self.left_frame.grid_rowconfigure(4,weight=1)
        self.left_frame.grid_rowconfigure(5,weight=1)
        self.left_frame.grid_rowconfigure(6,weight=1)
        self.left_frame.grid_rowconfigure(7,weight=1)
        self.left_frame.grid_rowconfigure(8,weight=1)
        self.left_frame.grid_rowconfigure(9,weight=1)
        self.left_frame.grid_columnconfigure(0,weight=1)
        self.left_frame.grid_columnconfigure(1,weight=1)

        #######################################################
        #SECCION COMERCIAL#
        #######################################################

        commercial_label = ttk.Label(self.left_frame,text='Comercial/KAM: ')
        commercial_label.grid(row=3, column=0, columnspan=2, sticky="w",pady=(0, 10),padx=(10,0))

        # "Selecciona FSC" label and input
        self.variableControl = tk.StringVar()
        fsc=ttk.Label(self.left_frame, text="Selecciona FSC")
        fsc.grid(row=1, column=0, sticky="ew",pady=(15, 5),padx=(10,0))


        fsc_entry = ttk.Entry(self.left_frame,textvariable=self.variableControl , state='readonly')
        fsc_entry.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10),padx=(10,20))

        self.controlLabel=tk.StringVar()
        com_entry = ttk.Entry(self.left_frame,textvariable=self.controlLabel , state='readonly')
        com_entry.grid(row=3, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        examine_button = ttk.Button(self.left_frame, text="Examinar",command=self.browse_file)
        examine_button.grid(row=1, column=1, sticky='ew',pady=(20, 5),padx=(0,20))

        #########################################################
        #SECCION ENCABEZADO
        #########################################################


        valores_para_listas=[['Q2-2024','Q3-2024','Q4-2024','Q1-2025','Q2-2025','Q3-2025','Q4-2025'],
                            sp.gastos_de_viaje(self.directorio)]
        
        label = ["Improvistos:", "Estampillas:","Institución", "Trimestre Esperado:", 
                "Ciudad:","Presupuesto: ","Moneda Requerida:"]

        self.improvistosVariable=tk.StringVar()
        improvistos=ttk.Label(self.left_frame, text=label[0])
        improvistos.grid(row=4, column=0, sticky="w",pady=(0, 10),padx=(10,0))

        self.estampillasVariable=tk.StringVar()
        estampillas=ttk.Label(self.left_frame, text=label[1])
        estampillas.grid(row=5, column=0, sticky="w",pady=(0, 10),padx=(10,0))

        self.institucionVariable=tk.StringVar()
        institucion=ttk.Label(self.left_frame, text=label[2])
        institucion.grid(row=6, column=0, sticky="w",pady=(0, 10),padx=(10,0))


        trimistre=ttk.Label(self.left_frame, text=label[3])
        trimistre.grid(row=8, column=0, sticky="w",pady=(0, 10),padx=(10,0))

        ciudad=ttk.Label(self.left_frame, text=label[4])
        ciudad.grid(row=9, column=0, sticky="w",pady=(0, 10),padx=(10,0))

        self.presupuestoVar=tk.StringVar()
        presupuestoL=ttk.Label(self.left_frame, text=label[5])
        presupuestoL.grid(row=7, column=0, sticky="w",pady=(0, 10),padx=(10,0))

        presupuestoEntry=ttk.Entry(self.left_frame,validate='key',validatecommand=self.vcmd,
                                   textvariable=self.presupuestoVar)
        presupuestoEntry.grid(row=7, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        improvistosEntry=ttk.Entry(self.left_frame,validate='key',validatecommand=self.vcmd,
                                   textvariable=self.improvistosVariable)
        improvistosEntry.grid(row=4, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        estampillasEntry=ttk.Entry(self.left_frame,validate='key',validatecommand=self.vcmd,
                                   textvariable=self.estampillasVariable)
        estampillasEntry.grid(row=5, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        institucionEntry=ttk.Entry(self.left_frame,textvariable=self.institucionVariable)
        institucionEntry.grid(row=6, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        self.combovar=tk.StringVar()
        self.comboovar=tk.StringVar()

        combobox = ttk.Combobox(self.left_frame,values=valores_para_listas[0],textvariable=self.combovar)
        combobox.grid(row=8, column=1, sticky="ew", pady=(0, 10),padx=(0,20))
        comboobox = ttk.Combobox(self.left_frame,values=valores_para_listas[1],textvariable=self.comboovar)
        comboobox.grid(row=9, column=1, sticky="ew", pady=(0, 10),padx=(0,20))

        combobox.bind('<FocusOut>', 
                    lambda event, a=self.combovar,b=valores_para_listas[0]:self.on_combobox_change(event,a,b))
        combobox.bind('<Return>', 
                    lambda event, a=self.combovar,b=valores_para_listas[0]:self.on_combobox_change(event,a,b))
        comboobox.bind('<FocusOut>', 
                    lambda event, a=self.comboovar,b=valores_para_listas[1]:self.on_combobox_change(event,a,b))
        comboobox.bind('<Return>', 
                    lambda event, a=self.comboovar,b=valores_para_listas[1]:self.on_combobox_change(event,a,b))

        monedaLabel=ttk.Label(self.left_frame, text=label[6],style="Large.TLabel")
        monedaLabel.grid(row=10, column=0, sticky="w",pady=(0, 40),padx=(10,0))
        self.monedacomboVar=tk.StringVar()
        valores_moneda=['COP','USD','EUR']
        monedacombo=ttk.Combobox(self.left_frame,values=valores_moneda,textvariable=self.monedacomboVar)
        monedacombo.grid(row=10, column=1, sticky="ew", pady=(0, 40),padx=(0,20))

        monedacombo.bind('<FocusOut>', 
                    lambda event, a=self.monedacomboVar,b=valores_moneda:self.on_combobox_change(event,a,b))
        monedacombo.bind('<Return>', 
                    lambda event, a=self.monedacomboVar,b=valores_moneda:self.on_combobox_change(event,a,b))

    def right_frame_creation(self):
        #########################################################
        #SECCION REFERENCIAS Y MARCA
        #########################################################

        top_frame= tk.LabelFrame(self,text='Seleccionar Referencias')
        top_frame.grid(row=1,column=1,columnspan=2,sticky='ew',padx=(10,10))
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_columnconfigure(1, weight=1)
        top_frame.grid_rowconfigure(1, weight=1)
        top_frame.grid_rowconfigure(0, weight=1)

        carpetasPosibles=['PHY','TER','3B','ELECTRO','LN','EU','PT']
        self.carpetaVariable=tk.StringVar()
        nombreCarpeta=ttk.Label(top_frame,text="Carpeta")
        nombreCarpeta.grid(row=0, column=2,sticky='w',pady=(20,10))
        self.carpeta=ttk.Combobox(top_frame,values=carpetasPosibles,
                        textvariable=self.carpetaVariable)
        self.carpeta.grid(row=0,column=3,padx=(10,30),pady=(20,10),sticky='w')
        self.carpeta.set(carpetasPosibles[3])
        self.carpeta.bind('<FocusOut>', 
                    lambda event, a=self.carpetaVariable,b=carpetasPosibles:self.on_combobox_change(event,a,b))
        self.carpeta.bind('<Return>', 
                    lambda event, a=self.carpetaVariable,b=carpetasPosibles:self.on_combobox_change(event,a,b))

        marcalabel=ttk.Label(top_frame, text="Base de Datos")
        marcalabel.grid(row=0, column=0,sticky='ns',pady=(30,10))
        reflabel=ttk.Label(top_frame, text="Referencia")
        reflabel.grid(row=1, column=0,sticky='ns',pady=(0,10))

        marcaslista=['PHYWE','EUROMEX']
        marcavariable=tk.StringVar()
        self.marca=ttk.Combobox(top_frame,values=marcaslista,
                        textvariable=marcavariable)
        self.marca.grid(row=0,column=1,padx=(30,0),pady=(30,10),sticky='w')

        self.marca.bind('<FocusOut>', 
                    lambda event, a=marcavariable,b=marcaslista:self.on_combobox_change(event,a,b))
        self.marca.bind('<Return>', 
                    lambda event, a=marcavariable,b=marcaslista:self.on_combobox_change(event,a,b))
        self.marca.bind('<<ComboboxSelected>>', self.actualizar_referencias_por_seleccion)

        self.searchVar=tk.StringVar()
        search_entry = ttk.Entry(top_frame,textvariable=self.searchVar,width=100)
        search_entry.grid(row=1, column=1, columnspan=3, padx=(0,30), pady=(0,10),sticky='w')
        self.lista_completa_referencias = list(sp.nombres_de_basedeDatos('PHYWE',self.directorio))

        ##########################
        # Listboxes with scrollbar
        ##########################
        listbox_frame = ttk.Frame(top_frame,style='Custom.TFrame')
        listbox_frame.grid(row=3, column=0,columnspan=4,  sticky="nsew", pady=(10, 10),padx=20)

        listbox_frame.grid_rowconfigure(0,weight=1)
        listbox_frame.grid_columnconfigure(0,weight=1)
        listbox_frame.grid_columnconfigure(1,weight=0)
        listbox_frame.grid_columnconfigure(2,weight=1)
        listbox_frame.grid_columnconfigure(3,weight=1)
        listbox_frame.grid_columnconfigure(4,weight=0)


        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical")
        scrollbar.grid(row=0, column=1, sticky="ns")

        scrollbar2=ttk.Scrollbar(listbox_frame, orient="vertical")
        scrollbar2.grid(row=0, column=4, sticky="ns")


        self.ref_listbox = tk.Listbox(listbox_frame, exportselection=False, 
                                yscrollcommand=scrollbar.set,width=55)
        self.ref_listbox.grid(row=0, column=0, sticky="w")
        scrollbar.config(command=self.ref_listbox.yview)

        self.selected_listbox = tk.Listbox(listbox_frame, exportselection=False,
                                    yscrollcommand=scrollbar2.set,width=55)
        self.selected_listbox.grid(row=0, column=3, sticky="e")
        scrollbar2.config(command=self.selected_listbox.yview)

        scrollbar3 =ttk.Scrollbar(listbox_frame,orient='horizontal')
        scrollbar3.grid(row=1,column=0,sticky='ew')
        scrollbar3.config(command=self.ref_listbox.xview)

        scrollbar4 =ttk.Scrollbar(listbox_frame,orient='horizontal')
        scrollbar4.grid(row=1,column=3,sticky='ew')
        scrollbar4.config(command=self.selected_listbox.xview)

        #Buttons to move items between listboxes
        btn_frame = ttk.Frame(listbox_frame)
        btn_frame.grid(row=0, column=2, sticky="nsew",padx=(0,0))

        # Configurar las filas del frame de botones para que se expandan igualmente
        btn_frame.grid_rowconfigure(0, weight=1)  # Fila por encima de los botones
        btn_frame.grid_rowconfigure(2, weight=1)  
        btn_frame.grid_rowconfigure(4, weight=1)  
        btn_frame.grid_rowconfigure(1, weight=1)  # Puedes ajustar este peso si necesitas más control sobre la posición de los botones
        btn_frame.grid_rowconfigure(3, weight=1)  # Fila entre los botones (si deseas espacio entre ellos)
        btn_frame.grid_rowconfigure(5, weight=1)  # Fila por debajo de los botones
        btn_frame.grid_columnconfigure(0,weight=1)

        move_to_selected_button = ttk.Button(btn_frame, text=">", 
                                            command=self.move_to_selected, width=3)
        move_to_selected_button.grid(row=2,column=0,sticky='ew')

        move_to_references_button = ttk.Button(btn_frame, text="X",
                                            command=self.move_to_references, width=3)

        move_to_references_button.grid(row=4,column=0,sticky='ew')

        load_csv_button = ttk.Button(listbox_frame, text="Cargar CSV", command=self.cargar_csv)
        load_csv_button.grid(row=2, column=0, padx=(10, 10), pady=(10, 10),sticky='ew')

        clear_button = ttk.Button(listbox_frame, text="Borrar Todo", command=self.clear_listbox)
        clear_button.grid(row=2, column=3, padx=(20, 10), pady=(10, 10),sticky='ew')

        self.actualizar_ref_listbox()
        self.searchVar.trace_add("write", self.on_search_entry_change)

    def clear_listbox(self):
            self.selected_listbox.delete(0, tk.END)
        
    def oferta_frame(self):
        #########################
        #RADIO BUTTONS FOR OFFER 
        #########################
        # Radio button options (independent groups)
        self.radio_values = {
            "Tipo": tk.StringVar(value="Público"),
            "Requerimiento": tk.StringVar(value="Normal"),
            "Canal": tk.StringVar(value="Institucional")
        }

        # Creating the radio buttons for each option
        options_frame = tk.LabelFrame(self,text='Información para Oferta')
        options_frame.grid(row=0, column=2,padx=(5,5),pady=(10,10),sticky='ns')
        # options_frame.grid_columnconfigure(0,weight=1)
        # options_frame.grid_columnconfigure(1,weight=1)
        # options_frame.grid_columnconfigure(2,weight=1)
        # options_frame.grid_columnconfigure(3,weight=1)

        for i, (label, var) in enumerate(self.radio_values.items()):
            ttk.Label(options_frame, text=label).grid(row=i, column=0, padx=(15,10),pady=(1,1),sticky="w")
            options_frame.grid_rowconfigure(i, weight=1)

            if label == "Tipo":
                choices = ["Público", "Privado", "Mixto"]
            elif label == "Requerimiento":
                choices = ["Urgente", "Normal"]
            else: # Canal
                choices = ["Institucional", "Proyectos", "Presidencia"]

            # Creating the radio buttons
            for j, choice in enumerate(choices):
                radio_btn = ttk.Radiobutton(options_frame, text=choice, 
                                            variable=var, value=choice,style='Custom.TRadiobutton')
                radio_btn.grid(row=i, column=j+1,padx=(5,5),pady=(1,1), sticky="w")

    def creacion_frame(self):

        thirdframe=tk.LabelFrame(self,text='Creación del Documento')
        thirdframe.grid(row=2,column=1,columnspan=3,sticky='nsew',pady=(10,10),padx=10)

        nombreFinal=ttk.Label(thirdframe, text="Nombre final de carpeta")
        nombreFinal.grid(row=0, column=0,sticky='ns',pady=(10, 0))

        self.nombreCarpetaFinalVariable=tk.StringVar()

        nombreCarpetaFinal=ttk.Entry(thirdframe,textvariable=self.nombreCarpetaFinalVariable,width=80)
        nombreCarpetaFinal.grid(row=1,column=0,sticky='ew',pady=(0, 10),padx=(30,30))

        thirdframe.grid_columnconfigure(0,weight=1)
        thirdframe.grid_rowconfigure(0,weight=1)        

        dobutton=ttk.Button(thirdframe, text="Crear SP",command=self.manejar_advertencias,width=20)
        dobutton.grid(row=2, column=0, sticky='ew',pady=(0, 10),padx=(30,30),ipady=3)

    def gastos_frame(self):

        firstframe=tk.LabelFrame(self,text='Gastos Operativos')
        firstframe.grid(row=2,column=0,sticky='ns',pady=(10,10),padx=(10,0))

        gastosope=ttk.Label(firstframe, text="¿Hay Gastos?")
        gastosope.grid(row=0, column=0,sticky='e',pady=(10, 10),padx=(10,0))

        self.switch_var = tk.BooleanVar()
        switch_button=ttk.Checkbutton(firstframe,style="Switch.TCheckbutton", text="", 
                                    variable=self.switch_var, 
                                    onvalue=True, offvalue=False, command=self.on_switch)
        switch_button.grid(row=0,column=1,padx=10,pady=(10, 10))

        self.num_pro_var=tk.StringVar()
        num_pro=ttk.Label(firstframe,text='Número de profesionales: ')
        num_pro.grid(row=1,column=0,sticky='ew',pady=(0,10),padx=(10,0))

        self.num_dias_var=tk.StringVar()
        num_dias=ttk.Label(firstframe,text='Días: ')
        num_dias.grid(row=2,column=0,sticky='e',pady=(0,10),padx=(10,0))

        profesionales=ttk.Entry(firstframe,validate='key',validatecommand=self.vcmd,textvariable= self.num_pro_var)
        profesionales.grid(row=1,column=1,sticky='ew',pady=(0,10),padx=(0,20))

        dias=ttk.Entry(firstframe,validate='key',validatecommand=self.vcmd,textvariable=self.num_dias_var)
        dias.grid(row=2,column=1,sticky='ew',pady=(0,10),padx=(0,20))

        self.widgets_to_control=[profesionales,dias]
        self.on_switch()

        self.numero_consecutivo=tk.StringVar()

        self.carpetaVariable.trace_add('write', self.actualizar_entry1)
        self.institucionVariable.trace_add('write', self.actualizar_entry1)
        self.variableControl.trace_add('write', self.actualizar_entry1)
        self.carpeta.bind('<<ComboboxSelected>>', self.actualizar_entry1)

    def buscar_ruta_de_trabajo(self):
        ruta_valida=self.controlador.examinar_buscar_ruta()
        if ruta_valida==0:
            pass
        else:
            self.variable_de_ruta.set(ruta_valida)

    def on_validate(self,P):
        if P == "":
            return True  # Permite el valor vacío (para poder borrar)
        if P == ".":  # Permite que se ingrese un solo punto decimal
            return True
        try:
            float(P)  # Intenta convertir el valor propuesto a float
            return True
        except ValueError:
            return False

    def browse_file(self):
        filename = filedialog.askopenfilename()
        comercial= sp.encontrar_pdf_y_extraer_nombre(filename)
        if comercial:
            self.controlLabel.set(comercial)
            self.variableControl.set(filename)
        else:
            messagebox.showerror("¡Error!","¡Debe seleccionar un FSC!")

    def move_to_selected(self):
        selected = self.ref_listbox.curselection()
        for i in selected[::-1]:  # Revertir para manejar múltiples selecciones correctamente
            item = self.ref_listbox.get(i)
            self.selected_listbox.insert(tk.END, item)
            # No eliminar el ítem de ref_listbox para cumplir con el reque

    def move_to_references(self):
        selected = self.selected_listbox.curselection()
        for i in selected[::-1]:  # Revertir para manejar múltiples selecciones correctamente
            self.selected_listbox.delete(i)

    def on_combobox_change(self,event,combobox_var,values):
        # Obtiene el texto actual del combobox
        current_text = combobox_var.get()
        
        # Encuentra la coincidencia más cercana de la lista de valores
        closest_match = self.find_closest_match(current_text, values)
        
        # Si hay una coincidencia cercana, actualiza el texto del combobox con ese valor
        if closest_match:
            combobox_var.set(closest_match)
        else:
            # Si no hay ninguna coincidencia, limpia el combobox
            combobox_var.set('')

    def find_closest_match(self,text, values_list):
        # Encuentra el valor más cercano que comience con el texto ingresado
        for value in values_list:
            if value.lower().startswith(text.lower()):
                return value
        return None

    def actualizar_referencias_por_seleccion(self,event):
        seleccion = self.marca.get()  # Obtiene el valor actual seleccionado en la Combobox
        self.lista_completa_referencias = list(sp.nombres_de_basedeDatos(seleccion,self.directorio))
        self.actualizar_ref_listbox()  # Actualiza la listbox con las referencias correspondientes

    def extraer_informacion(self):

        datos = {
            "Comercial": self.controlLabel.get(),
            "Imprevistos": self.improvistosVariable.get(),
            "Estampillas": self.estampillasVariable.get(),
            "Institucion": self.institucionVariable.get(),
            "Trimestre": self.combovar.get(),
            "Ciudad": self.comboovar.get(),
            "Carpeta": self.nombreCarpetaFinalVariable.get(),
            "Tipo": self.radio_values["Tipo"].get(),
            "Requerimiento": self.radio_values["Requerimiento"].get(),
            "Canal": self.radio_values["Canal"].get(),
            "Presupuesto": self.presupuestoVar.get(),
            "Consecutivo": self.consecutivo,
            "Profesionales":  self.num_pro_var.get(),
            "Dias": self.num_dias_var.get(),
            "Moneda":self.monedacomboVar.get()
        }
        return datos

    def preparar_valor(self,valor):
        """
        Convierte NaN a cadena vacía y mantiene el resto de los valores.
        """
        if pd.isna(valor):
            return ""
        else:
            return valor
        
    def manejar_advertencias(self):
        if self.nombreCarpetaFinalVariable.get() == '':
            messagebox.showerror("Error","La carpeta debe tener un nombre no-vacio.")
        elif self.selected_listbox.size()==0:
            respuesta=messagebox.askyesno("Advertencia","No hay referencias seleccionadas en la lista. \n¿Continuar?")
            if respuesta:
                respuestass=messagebox.askokcancel("Creación de Solicitud",f"""
                                            Se creara una solicitud con el nombre:\n
                                            {self.nombreCarpetaFinalVariable.get()}\n
                                            Presione 'Ok' para continuar, o 'Cancelar'\n
                                            para corregir alguna información""")
                if respuestass:
                    self.crear_SP()
        else:
            respuestass=messagebox.askokcancel("Creación de Solicitud",f"""
                                            Se creara una solicitud con el nombre:\n
                                            {self.nombreCarpetaFinalVariable.get()}\n
                                            Presione 'Aceptar' para continuar, o 'Cancelar' para corregir alguna información""")
            if respuestass:
                self.crear_SP()

    def actualizar_ref_listbox(self,search_text=''):
        self.ref_listbox.delete(0, tk.END)  # Limpia la listbox antes de actualizarla
        for indice, nombre in self.lista_completa_referencias:
            if search_text.lower() in str(nombre).lower() or search_text.lower() in str(indice).lower():
                self.ref_listbox.insert(tk.END, f"{indice} - {nombre}")

    def on_search_entry_change(self,*args):
        search_text = self.searchVar.get()
        self.actualizar_ref_listbox(search_text)



    def cargar_csv(self):
        
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                df = pd.read_csv(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar el CSV: {str(e)}")
                return

            # Convertir el DataFrame a una lista de tuplas (referencia, cantidad)
            quantities_list = [(str(row[df.columns[0]]).strip(), str(row[df.columns[1]]).strip()) for _, row in df.iterrows()]
            

            # Obtener las referencias válidas desde el listbox, eliminando espacios en blanco
            valid_references = set(item.split(' - ')[0].strip() for item in self.ref_listbox.get(0, tk.END))
            
            added_count = 0  # Contador para referencias agregadas
            not_found_references = []  # Lista para referencias no encontradas

            # Lista para almacenar las cantidades con la referencia completa
            quantities_with_full_ref = []

            for referencia, cantidad in quantities_list:
                if referencia in valid_references:
                    for item in self.ref_listbox.get(0, tk.END):
                        if item.startswith(referencia + ' - '):
                            self.selected_listbox.insert(tk.END, item)
                            # Añadir a la lista con la referencia completa
                            quantities_with_full_ref.append((item, cantidad))
                            added_count += 1
                            break
                else:
                    not_found_references.append(referencia)

            if not_found_references:
                print(f"Referencias no válidas: {', '.join(not_found_references)}")

            messagebox.showinfo("Resultado", f"Referencias agregadas: {added_count}")

            # Asignar la lista actualizada de quantities_with_full_ref a quantities
            self.quantities = quantities_with_full_ref

 
    def on_switch(self):
        # Esta función se llama cada vez que el estado del switch cambia.
        # Puedes añadir aquí la lógica que necesites ejecutar cuando el switch cambia.
        if self.switch_var.get():
            for widget in self.widgets_to_control:
                widget.configure(state='normal')
        else:
            for widget in self.widgets_to_control:
                widget.delete(0,'end')
                widget.configure(state='disabled')

    def actualizar_entry1(self,*args):
        # Concatenar los valores de Entry2, Entry3, Combobox1 y variable_uno

        prefijo = self.carpetaVariable.get()
        nombredelainstitucion=self.institucionVariable.get()
        comercial=self.variableControl.get().split('/')[-1][:-4]
        numeroCons=str(sp.obtener_nuevo_consecutivo(prefijo,self.carpetas[prefijo],self.variable_de_ruta.get()))
        self.numero_consecutivo.set(numeroCons)
        self.consecutivo=prefijo+" "+self.numero_consecutivo.get()+"-24"
        if nombredelainstitucion== '':
            if comercial == '':
                valor_actualizado=self.consecutivo
            else:
                valor_actualizado=self.consecutivo+' '+comercial
        elif comercial== '':
            valor_actualizado=self.consecutivo+' '+nombredelainstitucion
        else:
            valor_actualizado = self.consecutivo+' '+comercial+' '+nombredelainstitucion
        # Actualizar el valor de Entry1
        self.nombreCarpetaFinalVariable.set(valor_actualizado)

    def crear_SP(self):
        datos=self.extraer_informacion()        
        ventana_top_level=CrearSPWindow(self, datos, self.variableControl.get(), self.carpetas,
                                        self.carpetas[self.carpetaVariable.get()],
                                        self.quantities,self.directorio,
                                        self.variable_de_ruta.get())


