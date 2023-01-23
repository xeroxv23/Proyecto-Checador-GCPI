# IMPORTAMOS LAS LIBRERIAS NECESARIAS PARA EL PROYECTO
import pandas as pd
import datetime

"""Con el metodo read_excel se crea con pandas un data frame que llevara el nombre df
se escoge a traves del metodo sheet_name el numero de hoja (partiendo 0 como numero 1), no tendra headers, con el metodo skiprows se ignoran las primeras 5 filas del archivo original y con el metodo usecols (A:F) se escogen solo los datos de las columnas mencionadas en excel"""

df = pd.read_excel('/home/xeroxv23/Documents/proyecto_checador/SEMANA_02_2023.xls', sheet_name=3, header=None, skiprows=4, usecols='A:F')

# Con el metodo loc se puede seleccionar rango de datos (A5:F829), el cual sera modificado mas adelante para solo tomar las filas de excel con valores 
df = df.loc[4:828]

# Con el metodo drop podemos Eliminar las columnas B y C que contienen datos que no necesitamos, podriamos eliminar esta linea de codigo mas adelante al simplemente seleccionar las columnas con usecols (A,D,E Y F) ... cabe mencionar que las variables en [[1, 2]] toman en cuenta que la fila A representa el indice 0

df = df.drop(df.columns[[1, 2]], axis=1)

# El metodo columns asigna nombres a los indices de las cabeceras o columnas

df.columns = ['CODIGO','FECHA','INGRESO','SALIDA']

# El metodo assign agrega columnas, que en este caso el primer parametro(HORAS_LABORADAS) es el nombre de la columna y contendra los siguientes valores(=0)

df = df.assign(HORAS_LABORADAS=0)

# El metodo fillna es utilizado para rellenar los valores faltantes (NaN) con algun valor, el metodo inplace es utilizado para indicar si se debe modificar el DATRAFRAME original o devolver una copia de los cambios 

df['INGRESO'].fillna(0, inplace=True)
df['SALIDA'].fillna(0, inplace=True)

# El metodo de pandas to_datetime se utiliza para convertir los valores de una columna en formato de tiempo y fecha, en este caso se le asigna el valor '%H:%M' y errors='coerce' significa que los valores que no puedan ser convertidos a datetime seran convertidos a NaT para no generar error

df['INGRESO'] = pd.to_datetime(df['INGRESO'], format='%H:%M', errors='coerce')
df['SALIDA'] = pd.to_datetime(df['SALIDA'], format='%H:%M', errors='coerce')

# La siguiente funcion, obtendra el resultado de horas laboradas restando como valores la SALIDA - INGRESO en formato de segundos y sera dividido para dar un numero entero, representando las horas, para esto se agrego una condicional donde si el valor no se encontraba o resultaba ser 0, se arrojaria el mensaje FALTA/NO CHECO, porque aunque un trabajador marcara su hora de ingreso, no se tendria informacion con la hora de salida.

def obtener_horas_laboradas(row):
    if row['SALIDA'] == 0 or row['INGRESO'] == 0:
        return 'FALTA/NO CHECO'
    else:
        return (row['SALIDA'] - row['INGRESO']).total_seconds() / 3600

df['HORAS_LABORADAS'] = df.apply(obtener_horas_laboradas, axis=1)
def obtener_dia_semana(fecha):
    fecha = datetime.datetime.strptime(fecha, "%Y-%m-%d")
    return fecha.strftime("%A")

df['HORAS_LABORADAS'].fillna('FALTA/NO CHECO', inplace=True)

df['DIA_SEMANA'] = df['FECHA'].apply(obtener_dia_semana)

columns = ['FECHA','DIA_SEMANA','INGRESO','SALIDA','CODIGO','HORAS_LABORADAS']
df = df.reindex(columns=columns)
df.insert(0, 'CODIGO', df.pop('CODIGO'))

df = df.sort_values(by=['CODIGO','FECHA'], 
                    axis=0, 
                    ascending=[True,True],
                    inplace=False)

# Guardar archivo como csv
df.to_csv('/home/xeroxv23/Documents/proyecto_checador/SEMANA_02/SEMANA_02_2023.csv', index=False)
print("Hemos guardado el archivo")

df.to_excel("/home/xeroxv23/Documents/proyecto_checador/SEMANA_02/SEMANA_2_TABLA_HORAS.xlsx", index=False)

