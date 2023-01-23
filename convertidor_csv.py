import pandas as pd
import datetime

# Leer archivo xls
df = pd.read_excel('/home/xeroxv23/Documents/proyecto_checador/SEMANA_02_2023.xls', sheet_name=3, header=None, skiprows=4, usecols='A:F')

# Seleccionar rango A5:F829
df = df.loc[4:828]

# Eliminar columnas B y C
df = df.drop(df.columns[[1, 2]], axis=1)

df.columns = ['CODIGO','FECHA','INGRESO','SALIDA']
df = df.assign(HORAS_LABORADAS=0)
df['INGRESO'].fillna(0, inplace=True)
df['SALIDA'].fillna(0, inplace=True)
df['INGRESO'] = pd.to_datetime(df['INGRESO'], format='%H:%M', errors='coerce')
df['SALIDA'] = pd.to_datetime(df['SALIDA'], format='%H:%M', errors='coerce')

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

