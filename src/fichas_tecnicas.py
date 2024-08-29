import os
import shutil
import pandas as pd

# Configure paths (adjust as needed) 
nombre_csv = "PHYWEQUOTE.csv"  # The name of the CSV file to be processed.
ruta_fichas = "G:/Mi unidad/Fichas_Tecnicas_Phywe"  # Path to the folder containing the technical sheets in Spanish.


def rutas_y_dataframe(ruta_destino,ruta_base):
    """
    :ruta_destino: Ruta en donde guardar las fichas
    :ruta_base: Ruta del proyecto donde se encuentra el csv
    """
    # Create the destination folder if it does not exist
    if not os.path.exists(ruta_destino):
        os.makedirs(ruta_destino)  # This ensures that the script can run without manual directory creation.

    # Attempt to read the CSV file
    try:
        # The CSV file is expected to use ';' as a separator. The regex replace function is used to remove all whitespace, which could be important for matching file names accurately.
        df = pd.read_csv(os.path.join(ruta_base, nombre_csv), sep=';').replace(r'\s+', '', regex=True)
    except FileNotFoundError:
        # Handles the case where the CSV file is not found in the expected directory, providing a clear error message.
        raise Exception("Archivo .csv no encontrado")

    return df

def buscar_y_copiar(nombre, ruta_origen, ruta_destino):
    for elemento in os.listdir(ruta_origen):
        if elemento.startswith(nombre):
            # Constructs the full source and destination paths for the file or directory to be copied.
            origen_completo = os.path.join(ruta_origen, elemento)
            destino_completo = os.path.join(ruta_destino, elemento)
            # Checks if the destination already exists to avoid overwriting without notice.
            if os.path.exists(destino_completo):
                # Deletes the existing file or directory to make room for the new copy. This could be problematic if there's a need to preserve existing files.
                if os.path.isfile(destino_completo):
                    os.remove(destino_completo)  # Deletes file.
                else:
                    shutil.rmtree(destino_completo)  # Deletes directory.
        
            # Determines whether to copy a file or a directory and performs the appropriate action.
            if os.path.isfile(origen_completo):
                shutil.copy(origen_completo, destino_completo)
            elif os.path.isdir(origen_completo):
                shutil.copytree(origen_completo, destino_completo)
            

            return True  # Indicates a successful copy operation.
    return False  # Indicates that the specified sheet was not found.

def main(ruta_destino,ruta_base):
    """
    :ruta_destino: Ruta en donde guardar las fichas
    :ruta_base: Ruta del proyecto donde se encuentra el csv
    """

    df=rutas_y_dataframe(ruta_destino,ruta_base)

    total_fichas=len(df)
    no_encontradas=[]
    #Iterate over the dataframe's first column, processing each sheet.
    for ficha in df.iloc[:, 0]:

        # Busca la ficha y la copia. Si no existe la ficha encontrada = false
        encontrada=buscar_y_copiar(ficha, ruta_fichas, ruta_destino)
        # Provides feedback if a sheet is specified in the CSV but not found in the source directory.
        if not encontrada:
            no_encontradas.append(ficha)

    return total_fichas-len(no_encontradas),total_fichas

