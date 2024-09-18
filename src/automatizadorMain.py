import os
import tkinter as tk
import sys
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import spInterfaz as sp
import informesInterfaz as it
import listaDeChequeoInterfaz as lc
import ofertaInterfaz as of

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
DIRECTORIO=get_resource_path()

class AutomatizadorApp(tk.Tk):
    def __init__(self, titulo, geometria):
        #inicializacion
        super().__init__()
        self.title(titulo)# Establece el título de la ventana
        self.geometry(f'{geometria[0]}x{geometria[1]}')# Configura las dimensiones de la ventana
        self.resizable(False,False)# Permite que la ventana sea redimensionable

        self.widgets()
        
        self.icono_e_imagen()

        self.diccionario_clases_documentos={
            "SP":("Crear Solicitud de Precios (SP)",sp.SPApp),
            "OF":("Crear Oferta",of.OfertaApp),
            "IT":("Crear Informes Técnicos",it.InformesApp),
            "LC":("Crear Listas de Chequeo",lc.listaCApp)
        }
        self.current_frame = None  # Inicialmente no hay frame visible

        # Manejar el evento de cierre de la ventana
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.mainloop()

    def widgets(self):

        self.menuFrame=ttk.Frame(self)
        self.menuFrame.grid(row=0,column=0)

        # Crear un separador vertical entre los frames
        separator = ttk.Separator(self, orient='vertical')
        separator.grid(row=0,column=1,sticky='nsw',padx=(0,0))
        
        self.frameDeTrabajo=ttk.Frame(self)
        self.frameDeTrabajo.grid(row=0,column=2,padx=5,pady=5)

        #Grid configure
        # self.menuFrame.grid_columnconfigure(0,weight=1)
        # self.menuFrame.grid_rowconfigure(0,weight=1)
        # self.menuFrame.grid_rowconfigure(1,weight=1)
        # self.menuFrame.grid_rowconfigure(2,weight=1)
        # self.menuFrame.grid_rowconfigure(3,weight=1)
        # self.menuFrame.grid_rowconfigure(4,weight=1)
 




        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        #Botones del menú
        self.botonHome=ttk.Button(self.menuFrame,text="Inicio", command=self.show_home)
        self.botonHome.grid(row=0,column=0,sticky='nsew',padx=20,pady=(20,50),ipady=5)

        self.botonSolicitudDePrecios=ttk.Button(self.menuFrame,text="SOLICITUD\nDE PRECIOS",
                                                command=lambda x='SP': self.show_frame(x),width=18)
        self.botonSolicitudDePrecios.grid(row=1,column=0,sticky='nsew',padx=20,pady=(20,20),ipady=8)

        self.botonOferta=ttk.Button(self.menuFrame,text="OFERTA",width=18,
                                    command=lambda x='OF': self.show_frame(x))
        self.botonOferta.grid(row=2,column=0,sticky='nsew',padx=20,pady=(20,20),ipady=8)

        self.botonInformeTecnico=ttk.Button(self.menuFrame,text="INFORME\nTÉCNICO",width=18,
                                            command=lambda x='IT': self.show_frame(x))
        self.botonInformeTecnico.grid(row=3,column=0,sticky='nsew',padx=20,pady=(20,20),ipady=8)

        self.botonListaChequeo=ttk.Button(self.menuFrame,text="LISTA DE\nCHEQUEO",width=18,
                                          command=lambda x='LC': self.show_frame(x))
        self.botonListaChequeo.grid(row=4,column=0,sticky='nsew',padx=20,pady=(20,20),ipady=8)

    def show_home(self):

        if self.current_frame:
            mensaje=messagebox.askokcancel("Advertencia","¿Desea borrar el progreso y volver al inicio?")
            if mensaje:
                self.current_frame.grid_forget()
                self.current_frame.destroy()
                self.current_frame = None

    def show_frame(self, documento):
        """
        Cambia el frame visible en la ventana principal a uno especificado.

        Args:
            documento: Frame que se desea mostrar.
        """
        mensaje=True
        if self.current_frame:            
            mensaje=messagebox.askokcancel("Advertencia","¿Desea borrar el progreso y cambiar de módulo?")
            if mensaje: 
                self.current_frame.grid_forget()
                self.current_frame.destroy()

        if mensaje:

            clase_documento_a_mostrar=self.diccionario_clases_documentos[documento]
            
            self.current_frame=clase_documento_a_mostrar[1](self.frameDeTrabajo,clase_documento_a_mostrar[0],self,DIRECTORIO)        
            self.current_frame.grid(row=0,column=0,padx=(5,5 if clase_documento_a_mostrar[1]==sp.SPApp else 80 ),pady=0,sticky='w')

            # Asegúrate de que el frameDeTrabajo esté correctamente configurado para que su contenido se expanda
            self.frameDeTrabajo.grid_columnconfigure(0, weight=0)
            self.frameDeTrabajo.grid_rowconfigure(0, weight=0)
 
    def examinar_buscar_ruta(self):
        filename = filedialog.askdirectory(mustexist=True,parent=self,title="Escoja un directorio de trabajo alternativo")
        if os.path.isdir(filename):
            return filename
        else:
            messagebox.showerror("¡Error!","Debe seleccionar un directorio existente")
            return 0

    def icono_e_imagen(self):
        # Configura el icono de la ventana usando un archivo desde el directorio del script
        self.iconbitmap(os.path.join(DIRECTORIO,'..','img',"imagen.ico"))
        
    def on_closing(self):
        """ Manejar el evento de cierre de la ventana """
        if messagebox.askokcancel("Salir", "¿Seguro que quieres salir?"):
            self.destroy()
            self.quit()
            
AutomatizadorApp("Gestión Documental",(1294,650))