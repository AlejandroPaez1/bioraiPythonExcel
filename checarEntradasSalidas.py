import pandas as pd
import json
from datetime import datetime

# Cargar datos desde el archivo JSON con encoding utf-8
with open("16-22abril.json", "r", encoding="utf-8") as file:
    data_json = json.load(file)

# Convertir el JSON a un DataFrame de Pandas
df = pd.DataFrame(data_json['data'])

# Convertir las columnas de fecha y hora a formato datetime
df['Checada'] = pd.to_datetime(df['Checada'])

# Ordenar por PIN y Checada
df = df.sort_values(by=['PIN', 'Checada'])

# Calcular fecha y hora de entrada y salida para cada día
df['Fecha'] = df['Checada'].dt.date
df['Hora_entrada'] = df.groupby(['PIN', 'Fecha'])['Checada'].transform('min')
df['Hora_salida'] = df.groupby(['PIN', 'Fecha'])['Checada'].transform('max')

# Identificar y marcar registros donde solo se realizó una checada en un día
df.loc[df['Hora_entrada'] == df['Hora_salida'], 'Hora_salida'] = 'no checó'

# Eliminar duplicados
df = df.drop_duplicates(subset=['PIN', 'Fecha'])

# Seleccionar las columnas requeridas
df = df[['PIN', 'Nombre de empleado', 'Hora_entrada', 'Hora_salida']]

# Renombrar las columnas
df = df.rename(columns={'Hora_entrada': 'Fecha ingreso', 'Hora_salida': 'Fecha final'})

# Guardar el DataFrame en un archivo Excel
excel_file = "registros_empleados3.xlsx"
df.to_excel(excel_file, index=False)

print("Archivo de Excel generado correctamente:", excel_file)
