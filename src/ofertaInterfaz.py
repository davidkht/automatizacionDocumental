import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import openpyxl
import oferta
import os
import sys
import shutil


def get_resource_path():
    """ Retorna la ruta absoluta al recurso, para uso en desarrollo o en el ejecutable empaquetado. """
    if getattr(sys, 'frozen', False):
        # Si el programa ha sido empaquetado, el directorio base es el que PyInstaller proporciona
        base_path = sys._MEIPASS
    else:
        # Si estamos en desarrollo, utiliza la ubicación del script
        base_path = os.path.dirname(os.path.realpath(__file__))

    return base_path

# Guarda la ruta del script para su uso posterior en la aplicación
script_directory=get_resource_path()

class OfertaApp(tk.LabelFrame):
    def __init__(self,parent,textoLabel,controlador,directorio_de_script):
        super().__init__(parent)
        self.config(text=textoLabel)
        self.controlador=controlador
        self.script_directory=directorio_de_script

        self.frames = {}# Almacena las instancias de los frames

        self.frame_uno=FirstFrame(self)
        self.frame_uno.grid(row=0, column=0, sticky="nsew")
        self.frames[FirstFrame] = self.frame_uno

        self.frame_dos=SecondFrame(self)
        self.frame_dos.grid(row=0, column=0, sticky="nsew")
        self.frames[SecondFrame] = self.frame_dos

        self.navigation_buttons()# Método para agregar los botones de navegación
        self.current_frame = FirstFrame# Frame actual visible
        self.show_frame(FirstFrame)# Muestra el primer frame
       
    def navigation_buttons(self):
        # Configura y coloca los botones de navegación en la ventana

        self.button_frame = ttk.Frame(self)
        self.button_frame.grid(row=1,column=0,sticky='ew',pady=20,padx=20)
        self.button_frame.grid_columnconfigure(0, weight=1)  # Asegura que el frame de botones se expanda horizontalmente

        # Botones para navegar entre pantallas o cerrar la aplicación
        self.back_button = ttk.Button(self.button_frame, text="Atrás",command=self.go_back)
        self.back_button.grid(row=1,column=0,sticky='e',pady=20,padx=(500,0),ipady=3)

        self.next_button = ttk.Button(self.button_frame, text="Siguiente",command=self.go_next)
        self.next_button.grid(row=1,column=1,sticky='e',pady=20,padx=20,ipady=3)

        self.cancel_button = ttk.Button(self.button_frame, text="Cancelar", 
                                        command=self.controlador.show_home)
        self.cancel_button.grid(row=1,column=2,sticky='e',pady=20,padx=40,ipady=3)

    def update_buttons(self):
        """
        Actualiza el estado y el texto de los botones de navegación basándose en el frame actual.
        
        Establece los botones 'Atrás' y 'Siguiente' según el contexto del frame mostrado:
        - Deshabilita el botón 'Atrás' si el frame actual es el primero.
        - Cambia el texto del botón 'Siguiente' a 'Finalizar' si el frame actual es el último.
        - Deshabilita el botón 'Siguiente' si las condiciones del frame actual no permiten avanzar.
        """    
        self.back_button['state'] = 'normal' if self.current_frame != FirstFrame else 'disabled'
        self.next_button['state'] = 'normal' if self.current_frame != SecondFrame else 'disabled'
        self.next_button['text'] = 'Finalizar' if self.current_frame == SecondFrame else 'Siguiente'
        if self.current_frame == FirstFrame and not self.frames[FirstFrame].can_go_to_next_page():
            self.next_button['state'] = 'disabled'
        if self.current_frame == SecondFrame and self.frames[SecondFrame].canFinish():
            self.next_button['state'] = 'normal'

    def configurar_grid(self):
        """
        Configura las propiedades de expansión de las filas y columnas de la ventana principal.
        
        Asegura que la fila y la columna donde se muestra el contenido principal puedan expandirse,
        mientras que la fila que contiene los botones de navegación no se expanda.
        """
        # Configurar la expansión de las filas y columnas
        self.grid_columnconfigure(0, weight=1)  # Hace que la columna se expanda
        self.grid_rowconfigure(0, weight=1)     # Hace que la fila del FirstFrame se expanda
        self.grid_rowconfigure(1, weight=0)     # Mantiene la fila de los botones sin expandir

    def rutas(self):
        """
        Obtiene y maneja las rutas de los archivos de proyecto y PVP, copiando el archivo PVP a la carpeta del proyecto.
            Returns:
            tuple: Devuelve una tupla conteniendo la ruta del archivo SP y la nueva ruta del archivo PVP después de copiarlo.
        """
        self.carpeta_proyecto=self.frame_uno.entryCarpetaVar.get()
        pvp=self.frame_uno.entryPVPVar.get()
        sp=self.frame_uno.ruta_sp
        try:
            destino=shutil.copy(pvp,self.carpeta_proyecto)
        except shutil.SameFileError:
            destino=pvp
        return (sp,destino)

    def show_frame(self, frame_class):
        """
        Cambia el frame visible en la ventana principal a uno especificado.

        Args:
            frame_class (class): La clase del frame que se desea mostrar.
        """
        frame = self.frames[frame_class]
        frame.tkraise()
        self.current_frame = frame_class
        self.update_buttons()

    def go_back(self):
        """
        Regresa al frame anterior si es posible.
        
        Vuelve al primer frame desde el segundo frame, no realiza acción si ya está en el primer frame.
        """
        if self.current_frame == SecondFrame:
            self.show_frame(FirstFrame)


    def go_next(self):
        """
        Avanza al siguiente frame o finaliza la aplicación en el último frame.
        
        Si está en el primer frame, procesa los datos y muestra el segundo frame.
        Si está en el segundo frame, intenta finalizar la oferta y maneja las excepciones mostrando un mensaje de error.
        """
        if self.current_frame == FirstFrame:
            sp,pvp=self.rutas()
            self.tablacomparativa=self.frame_dos.comparar_sp_vs_pvp(sp,pvp)
            self.frame_dos.create_widgets(self)
            self.show_frame(SecondFrame)
        elif self.current_frame == SecondFrame:
            try:
                oferta.llenar_oferta(self.carpeta_proyecto,self.frame_dos.df_pvp,self.tablacomparativa[2])
                messagebox.showinfo("Finalizado","Oferta finalizada\nRecuerde poner pre-requisitos.")
                exit()
            except Exception as e:
                messagebox.showerror("Error",str(e))


class FirstFrame(ttk.Frame):
    """
    Representa el primer frame de la aplicación donde se gestionan las entradas de la carpeta del proyecto
    y del archivo PVP. Permite al usuario seleccionar la carpeta y el archivo mediante diálogos de archivo,
    validando la presencia de ciertos archivos dentro de la carpeta seleccionada.
    """
    def __init__(self, parent):
        """
        Inicializa el frame, crea e incorpora todos los widgets necesarios.
        
        Args:
            parent (tk.Widget): Widget padre en el que se ubicará este frame.
        """
        super().__init__(parent)
        self.create_widgets(parent)
        self.place_widgets()
        
    def create_widgets(self,parent):
        """
        Crea todos los widgets que se usarán en el frame, como etiquetas, entradas y botones.
        
        Args:
            parent (tk.Widget): Widget padre para usar en callbacks, si es necesario.
        """
        # Creation of GUI components to be used in the frame.
        self.labelCarpeta = ttk.Label(self, text='Seleccionar Carpeta de Proyecto')
        self.labelPVP = ttk.Label(self, text='Seleccione el PVP del Proyecto')
        
        # Initialization of variables for storing user input from file dialogs.
        self.entryCarpetaVar = tk.StringVar()
        self.entryPVPVar = tk.StringVar()
        
        # Entry widgets for displaying the paths, set to readonly to prevent manual user edits.
        self.entryCarpeta = ttk.Entry(self, textvariable=self.entryCarpetaVar, width=152, state='readonly')
        self.entryPVP = ttk.Entry(self, textvariable=self.entryPVPVar, width=152, state='readonly')
        
        # Buttons for triggering the browse dialogs.
        self.buttonCarpeta = ttk.Button(self, text='Examinar', width=25, command=lambda parent=parent:self.browse_project_directory(parent))
        self.buttonPVP = ttk.Button(self, text='Examinar', width=25, command=lambda parent=parent:self.browse_file_pvp(parent))

    def place_widgets(self):
        """
        Organiza los widgets dentro del frame usando el método grid.
        """
        # Configure the grid layout, making sure some rows and columns are weighted to center content.
        self.columnconfigure(0, weight=1)
        self.columnconfigure(3, weight=1)

        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=0)
        self.rowconfigure(3, weight=1)
        self.rowconfigure(4, weight=0)
        self.rowconfigure(5, weight=0)
        self.rowconfigure(6, weight=1)

        # Placement of the widgets using grid layout to achieve desired UI structure.
        self.labelCarpeta.grid(row=1, column=1,sticky='w',pady=(80,2),padx=(20,20))
        self.labelPVP.grid(row=4, column=1,sticky='w',pady=(20,5),padx=(20,20))
        self.entryCarpeta.grid(row=2, column=1, sticky='ew',pady=5,padx=(20,20))
        self.entryPVP.grid(row=5, column=1, sticky='ew',pady=5,padx=(20,20))
        self.buttonCarpeta.grid(row=3, column=1, sticky='w',pady=(5,60),ipady=3,padx=(20,20))
        self.buttonPVP.grid(row=6, column=1, sticky='w',pady=(5,60),ipady=3,padx=(20,20))

    def browse_project_directory(self, parent):
        """
        Abre un diálogo para seleccionar una carpeta y busca un archivo específico dentro de ella.
        
        Args:
            parent (tk.Widget): Widget padre para usar en callbacks.
        """
        carpeta = filedialog.askdirectory()  # Pide al usuario que seleccione una carpeta
        if carpeta:
            archivo_sp = next((archivo for archivo in os.listdir(carpeta) if archivo.startswith('SP')), None)
            if archivo_sp:
                self.ruta_sp = os.path.join(carpeta, archivo_sp)  # Guarda la ruta completa del archivo encontrado
                self.entryCarpetaVar.set(carpeta)  # Muestra la carpeta en la entrada
            else:
                self.entryCarpetaVar.set('')  # Limpia la entrada si no se encuentra el archivo
                messagebox.showerror("¡Error!", "No hay un proyecto en la carpeta seleccionada!")
        else:
            self.entryCarpetaVar.set('')  # Limpia la entrada si el usuario cancela la selección
            messagebox.showerror("¡Error!", "No seleccionó nada!")

        parent.update_buttons()  # Actualiza el estado de los botones en el frame principal

    def browse_file_pvp(self, parent):
        """
        Abre un diálogo para seleccionar un archivo PVP y valida que el archivo tenga el prefijo correcto.
        
        Args:
            parent (tk.Widget): Widget padre para usar en callbacks.
        """
        ruta_pvp = filedialog.askopenfilename()  # Pide al usuario que seleccione un archivo
        archivo_pvp = ruta_pvp.split('/')[-1]  # Obtiene solo el nombre del archivo
        if archivo_pvp.startswith('PVP'):
            self.entryPVPVar.set(ruta_pvp)  # Muestra la ruta en la entrada si es válida
        else:
            self.entryPVPVar.set('')  # Limpia la entrada si el archivo no es válido
            messagebox.showerror("¡Error!", "Eso no es un PVP!")

        parent.update_buttons()  # Actualiza el estado de los botones en el frame principal

    def can_go_to_next_page(self):
        """
        Determina si el usuario puede avanzar al siguiente frame, basado en la validez de las entradas.

        Returns:
            bool: True si ambas entradas tienen rutas válidas, False en caso contrario.
        """
        return True if self.entryPVPVar.get() != '' and self.entryCarpetaVar.get() != '' else False
            
class SecondFrame(ttk.Frame):
    """
    Representa el segundo frame de la aplicación donde se realiza la comparación entre dos archivos,
    y se permite al usuario tomar decisiones basadas en los resultados de esa comparación. También
    proporciona opciones para confirmar la decisión y buscar fichas técnicas adicionales.
    """
    def __init__(self,parent):
        """
        Inicializa el frame, estableciendo la relación con el widget padre.
        
        Args:
            parent (tk.Widget): Widget padre en el que se ubicará este frame.
        """
        super().__init__(parent)

    def create_widgets(self,parent):
        """
        Crea y configura todos los widgets para este frame, incluyendo áreas de visualización de datos y controles.
        
        Args:
            parent (App): La instancia de la aplicación principal que actúa como padre de este frame.
        """
        #Frames
        
        titulo=parent.carpeta_proyecto.split('/')[-1]
        self.frameIzquierdo=ttk.LabelFrame(self,text=titulo,padding=(20, 20))
        self.frameDerecho= ttk.LabelFrame(self,text='',padding=(20, 20))
 

        #TreeView
        self.crear_tabla_comparativa(parent)

        #Frame derecho/Pregunta para continuar
        self.confirmacion_final(parent)
        #Frame Derecho/Boton fichas tecnicas
        self.boton_fichas_tecnicas(parent)

        
        self.place_widgets()

    def place_widgets(self):
        """
        Posiciona los frames y widgets internos dentro del layout del frame principal.
        """
        #SElf frame principal
        self.frameIzquierdo.grid(row=0,column=0,sticky='nsew',pady=(20,20),padx=(50,0))
        self.frameDerecho.grid(row=0,column=1,pady=(20,20),padx=(20,50))
        #frame izquierdo
        self.comparacion.grid(row=0,column=0,columnspan=2,pady=(25,25))        
        
    def comparar_sp_vs_pvp(self,sp,pvp):
        """
        Carga y compara los datos de los archivos SP y PVP, generando un DataFrame de comparación.
        
        Args:
            sp (str): Ruta al archivo SP.
            pvp (str): Ruta al archivo PVP.

        Returns:
            tuple: Contiene la tabla comparativa, totales calculados y la moneda usada.
        """
        #Extrae la moneda del SP
        wb=openpyxl.load_workbook(sp, read_only=True)
        ws=wb.worksheets[0]
        moneda=ws['E17'].value
        wb.close()

        df_sp=oferta.dataframe_sp(sp)
        self.df_pvp,totales=oferta.dataframe_pvp(pvp)
        try:        
            tabla_comparativa=oferta.generar_tabla_comparativa(df_sp,self.df_pvp,moneda)
        except ValueError as e:
            messagebox.showerror("Error fatal",e)
        return (tabla_comparativa,totales,moneda)
    
    def crear_tabla_comparativa(self,parent):
        """
        Crea y configura un widget TreeView para mostrar una tabla comparativa de datos.
    
        Esta función se encarga de inicializar y configurar un TreeView que sirve para presentar los datos comparativos
        entre dos fuentes de información (generalmente archivos SP y PVP). La configuración incluye establecer
        las columnas necesarias, asignar el ancho adecuado para cada columna y llenar el TreeView con los datos 
        obtenidos de una comparación previa.
        
        Args:
            parent (App): La instancia de la aplicación principal que actúa como padre de este frame.
        """
        # Inicializa el TreeView dentro del frame izquierdo y configura para mostrar solo los encabezados.
        self.comparacion=ttk.Treeview(self.frameIzquierdo,show="headings")
        
        # Limpia cualquier entrada anterior en el TreeView para evitar duplicaciones o datos obsoletos.
        for i in self.comparacion.get_children():
            self.comparacion.delete(i) 

        # Obtiene la tabla comparativa de datos, totales y la moneda desde el parent que almacena estos datos después de realizar la comparación.
        tablacomparativa=parent.tablacomparativa
        df=tablacomparativa[0]      # DataFrame con los datos comparativos
        totales=tablacomparativa[1] # Totales calculados de la comparación
        moneda=tablacomparativa[2]  # Moneda utilizada en los cálculos financieros

        


        # Agrega y configura las columnas del TreeView basándose en las columnas del DataFrame.
        columns = list(df.columns)
        self.comparacion['columns'] = columns
        for col in columns:
            self.comparacion.heading(col, text=col)
            self.comparacion.column(col, width=100,anchor='center')

        # Ajustes específicos para algunas columnas, estableciendo un ancho personalizado para mejorar la visibilidad.
        self.comparacion.column('ITEM', width=50)
        self.comparacion.column('NOMBRE', width=60)
        self.comparacion.column('NOMBRE',width=60)
        self.comparacion.column('REFERENCIA',width=78)
        self.comparacion.column('CANTIDAD',width=70)
        self.comparacion.column(4,width=110)
        self.comparacion.column(5,width=150)
        self.comparacion.column(6,width=75)


        # Inserta los datos del DataFrame en el TreeView fila por fila.
        for index, row in df.iterrows():
            self.comparacion.insert("", tk.END, values=list(row))

        # Establece etiquetas con los totales calculados para cada categoría mostrada en el frame izquierdo,
        # proporcionando un resumen visual inmediato al usuario.
        for i, (key, value) in enumerate(totales.items(), 1):
            ttk.Label(self.frameIzquierdo, text=key).grid(row=i, column=0, pady=(0, 10), padx=(10, 10), sticky='e')
            ttk.Label(self.frameIzquierdo, text=moneda + f" {value:.2f}").grid(row=i, column=1, pady=(0, 10), padx=(0, 10), sticky='w')

    def confirmacion_final(self,parent):
        """
        Crea controles para que el usuario confirme si está de acuerdo con los resultados mostrados.
        
        Args:
            parent (App): La instancia de la aplicación principal que actúa como padre de este frame.
        """
        ttk.Label(self.frameDerecho,text='Está de acuerdo con el costeo?').grid(row=0,column=0,sticky='nsew',pady=20)
        
        self.selected_option = tk.IntVar()

        radio1=ttk.Radiobutton(self.frameDerecho,text='SÍ',variable=self.selected_option,value=1,command=parent.update_buttons)
        radio2=ttk.Radiobutton(self.frameDerecho,text='NO',variable=self.selected_option,value=0,command=parent.update_buttons)
        radio1.grid(row=1,column=1,sticky='nsew',pady=6,padx=5)
        radio2.grid(row=1,column=2,sticky='nsew',pady=6,padx=5)

    def canFinish(self):
        """
        Determina si el usuario ha seleccionado la opción de confirmación positiva.
        
        Returns:
            bool: True si el usuario está de acuerdo, False en caso contrario.
        """
        return True if self.selected_option.get()==1 else False
    
    def boton_fichas_tecnicas(self,parent):
        """
        Crea un botón para buscar fichas técnicas Phywe asociadas al proyecto.
        
        Args:
            parent (App): La instancia de la aplicación principal que actúa como padre de este frame.
        """
        boton=ttk.Button(self.frameDerecho,text='Buscar Fichas Técnicas',command=lambda x=parent:self.fichas_tecnicas(x))
        boton.grid(row=2,column=0,columnspan=3,sticky='ns',padx=30,pady=50)

    def fichas_tecnicas(self,parent):
        """
        Busca fichas técnicas en la carpeta especificada del proyecto.
        
        Args:
            parent (App): La instancia de la aplicación principal que actúa como padre de este frame.
        """
        import fichas_tecnicas

        carpeta_fichas=os.path.join(parent.carpeta_proyecto,'FICHAS_TECNICAS')
        encontradas,totales=fichas_tecnicas.main(carpeta_fichas,parent.carpeta_proyecto)
        messagebox.showinfo('Atención',f'Se encontraron {encontradas} fichas técnicas de {totales}.')








