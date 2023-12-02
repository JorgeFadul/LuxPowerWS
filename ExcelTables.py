import os
import pandas as pd
import xlrd

# Ruta de la carpeta que contiene los archivos Excel
path = "C:/Users/jorge/OneDrive/Escritorio/Archivos CorpTesla/LuxPowerWS/TablasExcel/MZ Venao Solar - Octubre"

# Extraer el nombre de la carpeta
nombre_carpeta = os.path.basename(path)

# Obtener la ruta del directorio superior (carpeta padre)
carpeta_padre = os.path.abspath(os.path.join(path, os.pardir))

def extract(carpeta_input):

    # Lista para almacenar los datos de cada hoja de cada archivo
    datos = []

    # Iterar sobre cada archivo en la carpeta
    for archivo in os.listdir(carpeta_input):
        if archivo.endswith('.xls'):  # Asegurarse de que solo se procesen archivos Excel
            ruta_archivo = os.path.join(carpeta_input, archivo)

            # Leer el archivo Excel con xlrd
            libro_excel = xlrd.open_workbook(ruta_archivo)
            
            # Iterar sobre cada hoja en el archivo en sentido inverso
            for nombre_hoja in reversed(libro_excel.sheet_names()):
                hoja = libro_excel.sheet_by_name(nombre_hoja)
                
                # Obtener los datos de las columnas especificadas
                columnas_seleccionadas = [1, 2, 6, 7, 11, 12, 13, 20, 26, 27, 28]  # Índices de columna en base a la posición (empezando desde 0)
                datos_hoja = [hoja.col_values(col, start_rowx=1) for col in columnas_seleccionadas]
                
                # Crear un DataFrame con los datos seleccionados
                df_seleccionado = pd.DataFrame(
                    dict(zip(['Time', 'Status', 'vBat', 'soc', 'pCharge', 'pDisCharge', 'vacr', 'vepsr', 'pToGrid', 'pToUser', 'pLoad'], datos_hoja))
                )
                
                # Verificar si hay datos en el DataFrame seleccionado
                if not df_seleccionado.empty:
                    # Agregar los datos a la lista
                    datos.append(df_seleccionado)
                else:
                    print(f"No hay datos en la hoja '{nombre_hoja}' del archivo: {archivo}")
    if datos:
        # Combinar todos los datos en un solo DataFrame
        datos_combinados = pd.concat(datos, ignore_index=True)
        return datos_combinados
    else:
        return "No se encontraron datos"

def transform(datos=extract(path)):
    if datos.empty:
        print("No se encontraron datos en ninguno de los archivos.")
        return

    # Transformar la columna 'Time' a formato de fecha y hora
    datos['Time'] = pd.to_datetime(datos['Time'], errors='coerce')

    # Transformar las columnas 'vBat', 'soc', 'vacr' y 'vepsr' a números
    columnas_a_numeros = ['vBat', 'vacr', 'vepsr']
    for columna in columnas_a_numeros:
        datos[columna] = pd.to_numeric(datos[columna], errors='coerce')

    # Transformar la columna 'soc' a números (porcentajes almacenados como texto)
    datos['soc'] = datos['soc'].str.rstrip('%').astype('float') / 100.0

    return datos

def dashboard(writer, datos):
    # Crear una hoja vacía llamada 'Dashboard'
    workbook = writer.book
    dashboard_sheet = workbook.add_worksheet('Dashboard')

    # Obtener la columna 'Time' de la hoja 'Datos' como cadena
    columna_time_datos = datos['Time'].astype(str)

    # Escribir el encabezado en la hoja 'Dashboard'
    dashboard_sheet.write(0, 0, 'Time')

    # Escribir la columna 'Time' en la hoja 'Dashboard'
    for i, valor in enumerate(columna_time_datos, start=1):
        dashboard_sheet.write(i, 0, valor)

    # 1. Configurar la celda B1 con la fórmula "=F3"
    dashboard_sheet.write_formula('B1', '=$F$3')

    # 2. Combinar y aplicar formato a las celdas F3, G3, F4 y G4
    formato_negrita = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 22})
    dashboard_sheet.merge_range('F3:G4', '', formato_negrita)

    # 3. Configurar validación de datos en F3
    lista_desplegable = '=Datos!$B$1:$K$1'  # Ajusta esto según la ubicación de tus datos
    dashboard_sheet.data_validation('F3', {'validate': 'list',
                                           'source': lista_desplegable,
                                           'input_message': 'Elija un valor de la lista desplegable'})

    # 4. Combinar y aplicar formato a las celdas I3:L4
    dashboard_sheet.merge_range('I3:L4', '', formato_negrita)



def load(datos=transform()):
    if not datos.empty:
        # Construir el nombre del archivo de salida
        nombre_archivo_output = f'Datos - {nombre_carpeta}.xlsx'

        # Ruta del directorio superior (carpeta padre)
        carpeta_output = os.path.join(carpeta_padre, nombre_archivo_output)

        # Crear un escritor de Excel
        with pd.ExcelWriter(carpeta_output, engine='xlsxwriter') as writer:
            # Guardar el DataFrame combinado en la hoja 'Datos'
            datos.to_excel(writer, sheet_name='Datos', index=False)

            # Llamar a la función para crear la hoja 'Dashboard'
            dashboard(writer, datos)

        print(f"Los datos se han extraído y guardado en {carpeta_output}")
    else:
        print("No se encontraron datos en ninguno de los archivos.")

    

if __name__ == "__main__":
    datos_combinados = extract(path)
    
    if datos_combinados.empty:
        print("No se encontraron datos en ninguno de los archivos.")
    else:
        datos_transformados = transform(datos_combinados)
        load(datos_transformados)
