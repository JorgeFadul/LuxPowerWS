import os
import pandas as pd
import xlrd
import secret
import lxp_wscrap




# Extraer el nombre de la carpeta
carpeta_padre = secret.carpeta_descargas


def extract(carpeta_input, cliente):
    
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
                    print(f"datos agregados desde {archivo}")
                    
                    
                else:
                    print(f"No hay datos en la hoja '{nombre_hoja}' del archivo: {archivo}")
    if not len(datos) == 0:
        datos_combinados = pd.concat(datos, ignore_index=True)
        print(f"Todos los datos de {cliente} han sido añadidos")
        return datos_combinados

def transform(datos_t):
    
    if datos_t.empty:
        print("TRANSFORM No se encontraron datos en ninguno de los archivos.")
        return

    # Transformar la columna 'Time' a formato de fecha y hora
    datos_t['Time'] = pd.to_datetime(datos_t['Time'], errors='coerce')

    # Transformar las columnas 'vBat', 'soc', 'vacr' y 'vepsr' a números
    columnas_a_numeros = ['vBat', "pCharge", "pDisCharge", 'vacr', 'vepsr', "pToUser", "pLoad"]
    for columna in columnas_a_numeros:
        datos_t[columna] = pd.to_numeric(datos_t[columna], errors='coerce')

    # Transformar la columna 'soc' a números (porcentajes almacenados como texto)
    datos_t['soc'] = datos_t['soc'].str.rstrip('%').astype('float') / 100.0

    return datos_t

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



def load(datos_l, cliente):
    if not datos_l.empty:
        # Construir el nombre del archivo de salida
        nombre_archivo_output = f'Datos - {cliente}.xlsx'

        # Ruta del directorio superior (carpeta padre)
        carpeta_output = os.path.join(carpeta_padre, cliente, nombre_archivo_output)

        # Crear un escritor de Excel
        with pd.ExcelWriter(carpeta_output, engine='xlsxwriter') as writer:
            # Guardar el DataFrame combinado en la hoja 'Datos'
            datos_l.to_excel(writer, sheet_name='Datos', index=False)

            # # Llamar a la función para crear la hoja 'Dashboard'
            # dashboard(writer, datos_l)
        
        

        print(f"Los datos se han extraído y guardado en {carpeta_output}")
    else:
        print("LOAD: No se encontraron datos en ninguno de los archivos.")

    

def excelETL():
    for cliente in lxp_wscrap.client_list:
        datos_combinados = pd.DataFrame(extract(carpeta_input= os.path.join(carpeta_padre,cliente), cliente=cliente))
        
        if datos_combinados.empty:
            print("ETL No se encontraron datos en ninguno de los archivos.")
        else:
            datos_transformados = transform(datos_combinados)
            load(datos_l= datos_transformados, cliente=cliente)
        
        for archivo in os.listdir(os.path.join(secret.carpeta_descargas,cliente)):
            if archivo.endswith('.xls'):
                delete_path = os.path.join(secret.carpeta_descargas, cliente, archivo)
                os.remove(delete_path)

if __name__ == "__main__":
    excelETL()
