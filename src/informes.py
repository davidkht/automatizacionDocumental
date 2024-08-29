import os
import sys
import math
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image  # Importa Image para añadir imágenes a los archivos Excel
from win32com import client
import pdfManager as pdf

PLANTILLA='plantillaInforme.xlsx'
BASE_DE_DATOS='baseDeDatosParaInforme.xlsx'

def xsl2pdf(file_location):
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False

    workbook=app.Workbooks.open(file_location)
    output = os.path.splitext(file_location)[0]

    # Configuración de impresión
    worksheet = workbook.ActiveSheet
    worksheet.PageSetup.Zoom = False
    worksheet.PageSetup.FitToPagesWide = 1
    worksheet.PageSetup.FitToPagesTall = False  # Permitir que se extienda en múltiples páginas
    ##
    workbook.ActiveSheet.ExportAsFixedFormat(0,output)
    workbook.Close()

# Función para ajustar la altura de las celdas combinadas
def adjust_height(cell_range, text,sheet):
    # Asumir que cada columna tiene un ancho fijo estándar si no se especifica
    default_width = 8.43  # Este es el ancho estándar en Excel para columnas no ajustadas
    column_widths = [sheet.column_dimensions[openpyxl.utils.get_column_letter(cell.column)].width or default_width for cell in cell_range[0]]
    total_width = sum(column_widths)

    # Ajustar para mejor aproximación de caracteres por línea
    estimated_chars_per_line = total_width * 0.9  # Ajustar el factor según sea necesario

    # Calcular el número de líneas necesarias
    estimated_lines_needed = math.ceil(len(text) / estimated_chars_per_line)  # Estimar líneas necesarias basado en el ancho combinado
    
    # Añadir líneas adicionales por cada salto de línea en el texto
    explicit_lines = text.count('\n')

    # Total de líneas necesarias
    total_lines_needed = estimated_lines_needed + explicit_lines

    # Añadir margen adicional para asegurarse de que el texto es visible
    margin_factor = 1.2
    adjusted_lines_needed = total_lines_needed * margin_factor
    
    # Ajustar la altura de la fila
    cell_range[0][0].parent.row_dimensions[cell_range[0][0].row].height = adjusted_lines_needed * 15  # Ajustar la altura

def get_resource_path():
    """ Retorna la ruta absoluta al recurso, para uso en desarrollo o en el ejecutable empaquetado. """
    if getattr(sys, 'frozen', False):
        # Si el programa ha sido empaquetado, el directorio base es el que PyInstaller proporciona
        base_path = sys._MEIPASS
    else:
        # Si estamos en desarrollo, utiliza la ubicación del script
        base_path = os.path.dirname(os.path.realpath(__file__))

    return base_path

script_directory = get_resource_path()
ruta_plantilla=os.path.join(script_directory,'docs',PLANTILLA)
ruta_basedatos=os.path.join(script_directory,'docs',BASE_DE_DATOS)


def llenar_informe(serie,cliente,orden,contrato,direccion,ciudad):


    informe= openpyxl.load_workbook(ruta_plantilla)
    sheet= informe.worksheets[0]

    #ENCABEZADO

    consecutivo= str(serie.iloc[0])
    nombre_equipo=str(serie.iloc[1])

    celda_consecutivo=sheet["B7:E7"]
    celda_consecutivo[0][0].value = "REPORTE TÉCNICO No: " + consecutivo

    celda_nombre=sheet["B10:H10"]
    celda_nombre[0][0].value = "NOMBRE DEL EQUIPO: " + nombre_equipo

    celda_ubicacion=sheet["I10:L10"]
    celda_ubicacion[0][0].value = "UBICACIÓN: " + str(serie.iloc[2])

    celda_marca=sheet["B11:D11"]
    celda_marca[0][0].value = "MARCA: " + str(serie.iloc[3])

    celda_ref=sheet["E11:H11"]
    celda_ref[0][0].value = "REF.: " + str(serie.iloc[4])

    celda_serial=sheet["I11:L11"]
    celda_serial[0][0].value = "SERIAL: " + str(serie.iloc[5])

    celda_cliente=sheet["F7:L7"]
    celda_cliente[0][0].value = "CLIENTE: " + cliente

    celda_orden=sheet["B8:I8"]
    celda_orden[0][0].value = "ORDEN CONTRACTUAL: " + orden

    celda_contrato=sheet["J8:L8"]
    celda_contrato[0][0].value = "CONTRATO INTERNO: " + contrato

    celda_direccion=sheet["B9:H9"]
    celda_direccion[0][0].value = "DIRECCIÓN: " + direccion

    celda_ciudad=sheet["I9:L9"]
    celda_ciudad[0][0].value = "CIUDAD: " + ciudad


    #Parametros del informe
    celda_encendido=sheet["B18"]
    celda_encendido.value = "      Encendido: " + str(serie.iloc[6])

    celda_func=sheet["D18"]
    celda_func.value = "       Funcionamiento: " + str(serie.iloc[7])

    celda_gar=sheet["H18"]
    celda_gar.value = "Sellos de garantía: " + str(serie.iloc[8])

    celda_gar=sheet["K18"]
    celda_gar.value = "    Accesorios: " + str(serie.iloc[9])

    celda_estado=sheet["E26:L26"]
    celda_estado[0][0].value = str(serie.iloc[10])

    #pie de pagina del informe
    celda_mant=sheet["E27:L27"]
    celda_mant[0][0].value = str(serie.iloc[11])

    celda_fecha=sheet["F31:H32"]
    celda_fecha[0][0].value = str(serie.iloc[12])

    celda_h0=sheet["F33:H33"]
    celda_h0[0][0].value = "Hora de Inicio: " + str(serie.iloc[13])

    celda_hf=sheet["F34:H34"]
    celda_hf[0][0].value = "Hora finalización: " + str(serie.iloc[14])

    celda_nombre_cliente=sheet["I33:L33"]
    celda_nombre_cliente[0][0].value = "Nombre: " + str(serie.iloc[15])

    celda_email=sheet["I34:L34"]
    celda_email[0][0].value = "Correo E: " + str(serie.iloc[16])

    celda_tel=sheet["I35:L35"]
    celda_tel[0][0].value = "Número Contacto: " + str(serie.iloc[17])

    #Informe persé
    celda_estado_inicial=sheet["B15:L15"]
    celda_estado_inicial[0][0].value = str(serie.iloc[18])

    celda_descr=sheet["B23:L23"]
    celda_descr[0][0].value = str(serie.iloc[19])

    celda_recomend=sheet["B25:L25"]
    celda_recomend[0][0].value = str(serie.iloc[20])


    # Ajustar la altura de las celdas combinadas específicas
    adjust_height(celda_estado_inicial, str(serie.iloc[18]),sheet)
    adjust_height(celda_descr, str(serie.iloc[19]),sheet)
    adjust_height(celda_recomend, str(serie.iloc[20]),sheet)

    #Crea la imagen de encabezado
    img = Image(os.path.join(script_directory, 'img', 'encabezado.png'))
    sheet.add_image(img, 'B1')
    #Firma de Ingeniero
    firma = Image(os.path.join(script_directory, 'img', 'Firma.png'))
    sheet.add_image(firma, 'C31')


    #Cambia el nombre de la hoja
    sheet.title=consecutivo

    #Crea el arbol de carpetas para el informe
    carpeta=os.path.join(script_directory,'IT',consecutivo+" "+nombre_equipo)
    os.makedirs(carpeta)
    os.makedirs(os.path.join(carpeta,'REGISTRO AUDIOVISUAL'))

    #Guarda el archivo con el nombre correspondiente
    nombrearchivo=consecutivo+" "+nombre_equipo+'.xlsx'
    archivo=os.path.join(carpeta,nombrearchivo)
    informe.save(archivo)

    #Convierte el archivo a pdf
    xsl2pdf(archivo)

def ejecutar_automatizacion_informes(cliente,orden,contrato,direccion,ciudad):
    
    #lee base de datos
    df = pd.read_excel(ruta_basedatos,index_col=0,keep_default_na=False)

    # Asegúrate de que la columna de fechas se lea como datetime
    df['PROX MANT'] = pd.to_datetime(df['PROX MANT'])
    # Formatea la columna de fechas para que muestre solo la fecha en el formato deseado (día-mes-año)
    df['PROX MANT'] = df['PROX MANT'].dt.strftime('%d-%m-%Y')

    # Asegúrate de que la columna de fechas se lea como datetime
    df['FECHA'] = pd.to_datetime(df['FECHA'])
    # Formatea la columna de fechas para que muestre solo la fecha en el formato deseado (día-mes-año)
    df['FECHA'] = df['FECHA'].dt.strftime('%d-%m-%Y')

    for indice, fila in df.iterrows():
        llenar_informe(fila,cliente,orden,contrato,direccion,ciudad)

def buscar_pdf_y_registro(ruta_base):
    resultados = []
    
    # Recorre todos los subdirectorios de la ruta base
    for subdirectorio in os.listdir(ruta_base):
        ruta_subdirectorio = os.path.join(ruta_base, subdirectorio)                        
        if os.path.isdir(ruta_subdirectorio):
            # Buscar el archivo PDF en el subdirectorio
            archivo_pdf = None
            for archivo in os.listdir(ruta_subdirectorio):
                if archivo.endswith('.pdf'):
                    archivo_pdf = archivo
                    break
            
            if archivo_pdf:        
                resultados.append(( archivo_pdf, ruta_subdirectorio))
    
    return resultados

def unir_informe_con_fotos():

    registro_informes=os.path.join(script_directory,'IT')

    rutas_subdirectorios=buscar_pdf_y_registro(registro_informes)
    
    for subdirectorio in rutas_subdirectorios:

        informe= os.path.join(subdirectorio[1],subdirectorio[0])
        registro_audiovisual = os.path.join(subdirectorio[1],'REGISTRO AUDIOVISUAL')
        archivo_de_imagenes  = os.path.join(registro_audiovisual,subdirectorio[0])
        pdf.insert_images_to_pdf(registro_audiovisual,archivo_de_imagenes)
        pdf.merge_pdfs([informe,archivo_de_imagenes] , informe)

