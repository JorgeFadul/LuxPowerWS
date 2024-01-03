import math
from matplotlib.dates import DateFormatter
import pandas as pd
import matplotlib.pyplot as plt
import os
import secret
import lxp_wscrap

carpeta_padre = secret.carpeta_descargas


def dataframe_creación(cliente):
    nombre_archivo_output = f'Datos - {cliente}.xlsx'
    archivo = os.path.join(carpeta_padre, cliente, nombre_archivo_output)

    df = pd.read_excel(archivo)
    df['Time'] = pd.to_datetime(df['Time'])

    # Inicializa un diccionario para almacenar las filas de cada fecha
    dataframes_por_fecha = {}

    # Itera sobre las filas del DataFrame
    for index, row in df.iterrows():
        fecha_actual = row['Time'].date()

        # Create the list if it doesn't exist for the current date
        if fecha_actual not in dataframes_por_fecha:
            dataframes_por_fecha[fecha_actual] = []

        # Now the list is guaranteed to exist, so append safely
        dataframes_por_fecha[fecha_actual].append(row[['Time', 'pLoad']])

    # Ahora, dataframes_por_fecha contiene listas de filas separadas por fecha sin duplicados

    # Concatena las listas de filas en DataFrames
    for fecha, filas in dataframes_por_fecha.items():
        dataframes_por_fecha[fecha] = pd.DataFrame(filas, columns=['Time', 'pLoad']).reset_index(drop=True).drop_duplicates()
    
    return dataframes_por_fecha


def dataframe_tiempo():
    # Crea un rango de tiempo de 00:00:00 a 23:59:59 con intervalos de 5 minutos
    tiempo = pd.date_range(start='00:00:00', end='23:59:59', freq='5T')

    # Crea un dataframe de tiempo de referencia
    df_tiempo = pd.DataFrame(tiempo, columns=['tiempo'])

    # Convierte la columna 'tiempo' a un objeto pd.Timedelta
    df_tiempo['tiempo'] = df_tiempo['tiempo'].dt.time

    # Elimina la columna de índice de filas
    df_tiempo = df_tiempo.reset_index(drop=True)
    # print(df_tiempo)
    return df_tiempo


def dataframe_cliente(df_tiempo, df, cliente):
    df_combinado = pd.DataFrame()

    # Agrega la columna 'Time' del dataframe de tiempo de referencia
    df_combinado['Time'] = df_tiempo['tiempo']


    for fecha, dataframes in df.items():
        
        column_pLoad = 'pLoad ' + str(fecha)
        df_combinado[column_pLoad] = math.nan
        
        for indx, hora, valor in zip(dataframes.index.to_list(), dataframes["Time"].dt.time, dataframes["pLoad"]):
            # print(valor)
            mintime_list = []
            

            for indexref, timeref in zip(df_tiempo.index.to_list(), df_tiempo["tiempo"]):

                diff_time = abs((timeref.hour * 60 + timeref.minute) - (hora.hour * 60 + hora.minute))
                # print(indexref, diff_time)
                
                if not math.isnan(diff_time):
                    mintime_list.append([indexref, diff_time])
            
            # print(mintime_list)
            
            
            if not mintime_list == []:
                min_elem = min(mintime_list, key= lambda x: x[1])
                # print(min_elem)
                min_index = min_elem[0] if len(min_elem) == 2 else None
                # print(min_index)
                df_combinado.at[min_index, column_pLoad] = valor
                df_tiempo["tiempo"].drop(min_index)
                dataframes["Time"].drop(indx)



    df_combinado["Prom pLoad"] = df_combinado.loc[:, df_combinado.columns != "Time"].mean(axis=1)
    df_combinado['Time'] = pd.to_datetime(df_combinado["Time"], format="%H:%M:%S")

    print(df_combinado)
    df_combinado.to_csv(os.path.join(carpeta_padre,cliente,f"Consumo Promedio - {cliente}.csv"))
    return df_combinado


def grafica_dataframe(dataframe):
    plt.figure(figsize=(10, 6))
    plt.plot(dataframe["Time"], dataframe['Prom pLoad'], color='blue', marker='.')
    date_form = DateFormatter("%H,%M,%S")
    plt.gca().xaxis.set_major_formatter(date_form)
    plt.gcf().autofmt_xdate()
    plt.title('Gráfica de Time vs Prom pLoad')
    plt.xlabel('Time')
    plt.ylabel('Prom pLoad')
    plt.show()

def main_consumo():
    for cliente in lxp_wscrap.client_list:
        dataframe_cliente(df_tiempo=dataframe_tiempo(), df= dataframe_creación(cliente), cliente=cliente)

