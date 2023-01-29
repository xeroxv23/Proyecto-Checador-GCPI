# Importar librerias
import pandas as pd
from datetime import datetime
import time

# Leer el archivo de excel y seleccionar los datos
df = pd.read_excel('/home/xeroxv23/Documents/proyecto_checador/enero_2023.xlsx', sheet_name=0, usecols=[0,1,2,3,4],names=['CODIGO','NOMBRE','FECHA','INGRESO','SALIDA'],header=None)
df.drop([0], axis=0, inplace=True)

# Sustituir las filas con valores nulos en 0
df['INGRESO'].fillna(0, inplace=True)
df['SALIDA'].fillna(0, inplace=True)

# Convertir las columna 'time_str' a un objeto de fecha y hora
df["INGRESO"] = pd.to_datetime(df["INGRESO"],format='%H:%M',errors='coerce')
df["SALIDA"] = pd.to_datetime(df["SALIDA"],format='%H:%M',errors='coerce')

# Funcion para obtener las horas laboradas
import datetime

def obtener_horas_laboradas(row):
    if pd.isna(row['SALIDA']) or pd.isna(row['INGRESO']):
        return 'NO CHECO'
    else:
        duration = row['SALIDA'] - row['INGRESO']
        hours, remainder = divmod(duration.total_seconds(), 3600)
        minutes, seconds = divmod(remainder, 60)
        return datetime.time(int(hours), int(minutes), int(seconds))

df['HORAS_LABORADAS'] = df.apply(obtener_horas_laboradas, axis=1)

# Regresar el formato en HORA Y MINUTO
df['INGRESO'] = pd.to_datetime(df['INGRESO']).dt.time
df['SALIDA'] = pd.to_datetime(df['SALIDA']).dt.time

# Crear un objeto de escritura de Excel
writer = pd.ExcelWriter('/home/xeroxv23/Documents/proyecto_checador/ENERO_PERSONAL/checador_oficina.xlsx', engine='xlsxwriter')

# Escribir el dataframe en la primera hoja
df.to_excel(writer, sheet_name='DESGLOSE_HORAS_DIAS', index=False)

# Con el metodo .groupby estamos creando un nuevo dataframe donde los index(columnas) seran CODIGO Y HORAS LABORADAS y con el metodo .sum, se sumarian las horas cuando hay coincidencia de codigo de trabajador. El metodo pivot_table crea la tabla dinamica en el orden que se es requerido, dando como valores una sumatoria de las horas y como index el codigo

df = df[df['HORAS_LABORADAS']!='NO CHECO']
df['HORAS_LABORADAS'] = pd.to_datetime(df['HORAS_LABORADAS'], format='%H:%M:%S')
df['HORAS_LABORADAS'] = df['HORAS_LABORADAS'].dt.hour * 60 + df['HORAS_LABORADAS'].dt.minute
df['HORAS_LABORADAS_MIN'] = df['HORAS_LABORADAS'] / 60
df_agrupado = df.groupby(['CODIGO','NOMBRE']).sum()
df_agrupado = df_agrupado.drop('HORAS_LABORADAS', axis=1)

def redondear_horas(x):
    if x % 1 >= 0.60:
        return int(x) + 1 + (x % 1 - 0.60)
    else:
        return round(x, 2)

df_agrupado['HORAS_LABORADAS_MIN'] = df_agrupado['HORAS_LABORADAS_MIN'].apply(redondear_horas)

# Escribir la tabla dinámica en la segunda hoja
df_agrupado.to_excel(writer, sheet_name='ACUMULADO_HORAS', index=True)

# Guardar el archivo
writer.save()
print('GUARDAMOS EL ARCHIVO EXCEL')