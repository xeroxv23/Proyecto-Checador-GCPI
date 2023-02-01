import pandas as pd

# Leer archivo excel
df = pd.read_excel('/home/xeroxv23/Documents/Proyectos GCPI/proyecto_checador/enero_2023.xlsx', header=None, usecols='A:G', names=["CODIGO", "NOMBRE", "FECHA", "INGRESO", "SALIDA", "SAL_COM","ING_COM"])

# Eliminar las filas con valores nulos
df['INGRESO'].fillna(0, inplace=True)
df['SALIDA'].fillna(0, inplace=True)
df['SAL_COM'].fillna(0, inplace=True)
df['ING_COM'].fillna(0, inplace=True)

# Convertir las columna 'time_str' a un objeto de fecha y hora
df["INGRESO"] = pd.to_datetime(df["INGRESO"],format='%H:%M',errors='coerce')
df["SALIDA"] = pd.to_datetime(df["SALIDA"],format='%H:%M',errors='coerce')
df["SAL_COM"] = pd.to_datetime(df["SAL_COM"],format='%H:%M',errors='coerce')
df["ING_COM"] = pd.to_datetime(df["ING_COM"],format='%H:%M',errors='coerce')

# HORAS LABORADAS - FUNCION PARA AGREGAR EL VALOR CUANDO EL RESULTADO ES 0
def obtener_horas_laboradas(row):
    if row['SALIDA'] == 0 or row['INGRESO'] == 0:
        return 'NO CHECO'
    else:
        return (row['SALIDA'] - row['INGRESO']).total_seconds() / 3600
df['HORAS_LABORADAS'] = df.apply(obtener_horas_laboradas, axis=1)

# HORAS DE COMIDA - FUNCION PARA AGREGAR VALOR CUANDO EL RESULTADO ES 0
def obtener_horas_comida(row):
    if row['SAL_COM'] == 0 or row['ING_COM'] == 0:
        return -.30
    else:
        return (row['ING_COM'] - row['SAL_COM']).total_seconds() / 3600
df['TIEMPO_COMIDA'] = df.apply(obtener_horas_comida, axis=1)

"""# ELIMINAR LA COLUMNA TIEMPO COMIDA PARA SU VALOR AGREGARLO A TIEMPO
df['HORAS_LABORADAS'] = (df['HORAS_LABORADAS'] - df['TIEMPO_COMIDA'])
del(df['TIEMPO_COMIDA'])"""

# Guardar archivo como csv
df.to_csv('/home/xeroxv23/Documents/Proyectos GCPI/proyecto_checador/ENERO_PERSONAL/checador_oficina.csv', index=False)
print('ARCHIVO CREADO')


df['INGRESO'] = pd.to_datetime(df['INGRESO']).dt.time
df['SALIDA'] = pd.to_datetime(df['SALIDA']).dt.time
df['SAL_COM'] = pd.to_datetime(df['SAL_COM']).dt.time
df['ING_COM'] = pd.to_datetime(df['ING_COM']).dt.time

# Crear un objeto de escritura de Excel
writer = pd.ExcelWriter('/home/xeroxv23/Documents/Proyectos GCPI/proyecto_checador/ENERO_PERSONAL/checador_oficina.xlsx', engine='xlsxwriter')

# Escribir el dataframe en la primera hoja
df.to_excel(writer, sheet_name='DESGLOSE_HORAS_DIAS', index=False)



# Con el metodo .groupby estamos creando un nuevo dataframe donde los index(columnas) seran CODIGO Y HORAS LABORADAS y con el metodo .sum, se sumarian las horas cuando hay coincidencia de codigo de trabajador. El metodo pivot_table crea la tabla dinamica en el orden que se es requerido, dando como valores una sumatoria de las horas y como index el codigo
df_agrupado = df.groupby('CODIGO')['HORAS_LABORADAS'].sum()
tabla_dinamica = df.pivot_table(values='HORAS_LABORADAS', index='CODIGO', aggfunc='sum')

# Escribir la tabla dinámica en la segunda hoja
tabla_dinamica.to_excel(writer, sheet_name='Sheet2', index=True)

# Guardar el archivo
writer.save()
print('GUARDAMOS EL ARCHIVO EXCEL')