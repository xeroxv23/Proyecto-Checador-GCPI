import pandas as pd

df = pd.read_csv('/home/xeroxv23/Documents/proyecto_checador/SEMANA_02_2023.csv', header=None)

# Agregar columna vacía al final del DataFrame
df = df.assign(columna_vacia='')
print('Se agrego la nueva columna')
