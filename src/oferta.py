import os
import pandas as pd
import numpy as np
import sys
import openpyxl  # Importa openpyxl para trabajar con archivos Excel
from openpyxl.utils.dataframe import dataframe_to_rows  # Función para convertir un DataFrame de pandas en filas para Excel
from openpyxl.drawing.image import Image  # Importa Image para añadir imágenes a los archivos Excel
from datetime import datetime

TRM_ACTUAL_EURO=4300
TRM_ACTUAL_USD=3800

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

def limpiar_dataframe(df):
    """
    Limpia el DataFrame proporcionado eliminando filas y columnas completamente vacías y ajustando los encabezados.
    
    :param df: DataFrame de pandas que será limpiado.
    :return: Tupla que contiene el DataFrame limpio y un diccionario con los totales de algunas columnas específicas.
    """
    df.dropna(axis=0, inplace=True, how='all')  # Elimina filas completamente vacías
    df.dropna(axis=1, inplace=True, how='all')  # Elimina columnas completamente vacías

    # Ajustar el DataFrame seleccionando desde la segunda fila hasta la penúltima
    nuevos_encabezados = df.iloc[0]  # Toma la primera fila para el encabezado
    df = df[1:]  # Elimina la primera fila, manteniendo temporalmente la última fila

    df.columns = nuevos_encabezados  # Establece la primera fila como el encabezado del DataFrame
    fila_TOTALES = df.iloc[-1]  # Ahora fila_TOTALES debería tener los índices correctos

    df = df[:-1]  # Ahora elimina la última fila después de capturar los totales

    # Resetear el índice para reflejar los cambios
    df.reset_index(drop=True, inplace=True)

    # Crear un diccionario con los totales específicos
    totales = {
        'SUBTOTAL': fila_TOTALES['SUBTOTAL'],
        'IVA': fila_TOTALES['IVA'],
        'TOTAL INCLUIDO IVA': fila_TOTALES['TOTAL INCLUIDO IVA']
    }

    return df, totales
 
def convertir_moneda(fila, moneda_objetivo):
    """
    Convierte el valor unitario de compra de una fila del DataFrame a la moneda objetivo.
    
    :param fila: Un DataFrame que representa la fila a convertir. Debe contener las columnas 'MONEDA' y 'VALOR UNITARIO COMPRA'.
    :param moneda_objetivo: La moneda a la que se desea convertir el valor ('COP', 'EUR', 'USD').
    :return: El valor convertido a la moneda objetivo.
    """
    tasas_cambio = {
        'COP': {'COP': 1, 'USD': TRM_ACTUAL_USD, 'EUR': TRM_ACTUAL_EURO},
        'USD': {'COP': 1/TRM_ACTUAL_USD, 'USD': 1, 'EUR': TRM_ACTUAL_EURO/TRM_ACTUAL_USD},
        'EUR': {'COP': 1/TRM_ACTUAL_EURO, 'USD': TRM_ACTUAL_USD/TRM_ACTUAL_EURO, 'EUR': 1},
    }
    
    moneda_origen = fila['MONEDA']
    precio_origen = fila['VALOR UNITARIO COMPRA']
    tasa_cambio = tasas_cambio[moneda_objetivo].get(moneda_origen, 1)  # Default a 1 si la moneda no se encuentra
    
    return precio_origen * tasa_cambio

def dataframe_pvp(ruta_pvp):
    """
    Busca y lee el archivo Excel PVP en el directorio especificado, limpiando el DataFrame resultante.
    
    :param ruta_pvp: Ruta del archivo PVP.
    :return: Un DataFrame limpio y un diccionario con los totales de ciertas columnas.
    """

    pvp = pd.read_excel(ruta_pvp, header=None)  # Leer sin encabezado
    return limpiar_dataframe(pvp)

def dataframe_sp(ruta_sp):
    """
    Busca y lee el archivo Excel SP en el directorio especificado.
    
    :param ruta_sp: Ruta del archivo SP en el proyecto
    :return: Un DataFrame creado a partir del archivo Excel SP encontrado.
    """
    return pd.read_excel(ruta_sp, skiprows=18, sheet_name='SOLICITUD')

def generar_tabla_comparativa(SP, PVP, moneda_pvp):
    """
    Genera una tabla comparativa entre dos DataFrames (SP y PVP) según la moneda especificada.
    
    :param SP: DataFrame de SP que contiene información sobre productos.
    :param PVP: DataFrame de PVP que contiene información sobre precios de venta.
    :param moneda_pvp: Un string que indica la moneda del PVP (1: COP, 2: EURO, 3: USD).
    :return: Un DataFrame que contiene la tabla comparativa generada.
    """


    df_comparacion = pd.DataFrame()
    # Comparaciones directas entre columnas de ambos DataFrames
    df_comparacion['ITEM']=SP.index

    try:
        df_comparacion['NOMBRE'] = SP['DESCRIPCION'] == PVP['DESCRIPCION']
    except ValueError as e:
        raise ValueError("PVP no tiene el mismo nombre de columnas del SP, o no tienen la misma longitud.\ndf_comparacion['NOMBRE'] = SP['DESCRIPCION'] == PVP['DESCRIPCION']")
    
    df_comparacion['REFERENCIA'] = SP['REFERENCIA'] == PVP['REFERENCIA']
    df_comparacion['CANTIDAD'] = SP['CANTIDAD'] == PVP['CANTIDAD']

    # Uso de la función apply con lambda para convertir los precios según la moneda elegida
    if moneda_pvp == 'COP':
        df_comparacion['PRECIO SP EN COP'] = SP.apply(lambda fila: convertir_moneda(fila, 'COP'), axis=1)
    elif moneda_pvp == 'EUR':
        df_comparacion['PRECIO SP EN EUR'] = SP.apply(lambda fila: convertir_moneda(fila, 'EUR'), axis=1)
    elif moneda_pvp == 'USD':
        df_comparacion['PRECIO SP EN USD'] = SP.apply(lambda fila: convertir_moneda(fila, 'USD'), axis=1)
    
    PVP['SUBTOTAL UNITARIO'] = pd.to_numeric(PVP['SUBTOTAL UNITARIO'], errors='coerce')
    df_comparacion['SUBTOTAL PVP'] = PVP['SUBTOTAL UNITARIO'].round(3)
    

    df_comparacion['SUBTOTAL PVP'] = df_comparacion['SUBTOTAL PVP'].round(3)
    columna_precio_sp = 'PRECIO SP EN ' + moneda_pvp
    df_comparacion[columna_precio_sp] = df_comparacion[columna_precio_sp].replace(0, np.nan)

    df_comparacion['AUMENTO'] = df_comparacion['SUBTOTAL PVP'] / df_comparacion[columna_precio_sp]
    df_comparacion['AUMENTO'] = df_comparacion['AUMENTO'].fillna(0)
    df_comparacion['AUMENTO'] = df_comparacion['AUMENTO'].infer_objects(copy=False)
    df_comparacion['AUMENTO'] = df_comparacion['AUMENTO']

    return df_comparacion.round(3)

def llenar_oferta(directorio, dfpvp,moneda):
    """
    Llena y actualiza un archivo Excel de oferta con los datos proporcionados en el DataFrame dfpvp.
    
    :param directorio: Directorio donde se encuentra el archivo de oferta a actualizar.
    :param dfpvp: DataFrame que contiene los datos a insertar en el archivo de oferta.
    :moneda: Tipo de moneda del PVP
    """

    archivo_of = next((archivo for archivo in os.listdir(directorio) if archivo.startswith('OFERTA')), None)
    ruta_of = os.path.join(directorio, archivo_of)
    wb_OF = openpyxl.load_workbook(ruta_of)  # Loads the Excel workbook.
    hoja_destino= wb_OF.worksheets[0]

    # Iterates through the rows of the DataFrame and inserts them into the Excel sheet.
    for r_idx, row in enumerate(dataframe_to_rows(dfpvp.loc[:,['DESCRIPCION','REFERENCIA','CANTIDAD','SUBTOTAL UNITARIO']], index=False, header=False), 21):
        for c_idx, value in enumerate(row, 3):
            hoja_destino.cell(row=r_idx, column=c_idx, value=value)  # Inserts each value into the corresponding cell.

    # Obtener la fecha actual en el formato día/mes/año
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    hoja_destino['I17'] = f"FECHA: {fecha_actual}"


    # Valores a buscar en la columna 'G'
    valores_a_buscar = ['SUBTOTAL', 'IVA 19%', 'TOTAL']
    filas_encontradas = {}

    # Itera sobre la columna 'G'
    for fila in range(1, hoja_destino.max_row + 1):
        celda = hoja_destino[f'F{fila}']
        if celda.value in valores_a_buscar:
            filas_encontradas[celda.value] = fila

    # Itera sobre la columna 'C'
    for fila in range(1, hoja_destino.max_row + 1):
        celda = hoja_destino[f'C{fila}']
        if celda.value == 'MONEDA':
            moneda_celda_fila = fila
    
    hoja_destino[f'D{moneda_celda_fila}']=moneda

    currency_formats = {
    'COP': '[$COP] #,##0.00',
    'USD': '$#,##0.00',
    'EUR': '€#,##0.00'
    }
    currency_format = currency_formats.get(moneda, '[$COP] #,##0.00')  # Default a COP si la moneda no está definida
    # Aplicar el formato al rango de celdas F21:G75
    for row in hoja_destino['F21:G75']:
        for cell in row:
            cell.number_format = currency_format

    # Aplicar el mismo formato a celdas específicas adicionales
    additional_cells = [f'G{filas_encontradas['IVA 19%']}', f'G{filas_encontradas['SUBTOTAL']}', f'G{filas_encontradas['TOTAL']}']
    for cell_address in additional_cells:
        hoja_destino[cell_address].number_format = currency_format


    img = Image(os.path.join(script_directory,'..','docs','second.png'))
    hoja_destino.add_image(img, 'B1')

    wb_OF.save(ruta_of)