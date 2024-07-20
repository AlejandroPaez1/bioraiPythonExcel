import pandas as pd

def determinar_formato(checada_str):
    if '/' in checada_str:
        if len(checada_str) == 16:
            return "%d/%m/%Y %H:%M"
        elif len(checada_str) == 19:
            return "%d/%m/%Y %H:%M:%S"
    elif '-' in checada_str:
        if len(checada_str) == 16:
            return "%d-%m-%Y %H:%M"
        elif len(checada_str) == 19:
            return "%d-%m-%Y %H:%M:%S"
    raise ValueError("Formato de fecha y hora no reconocido")

# Cargar el CSV
ruta_csv = 'faltas.csv'  # Cambia esta línea con la ruta correcta
df = pd.read_csv(ruta_csv)

# Convertir la columna 'Checada' a datetime usando el validador
df['Checada'] = df['Checada'].apply(lambda x: pd.to_datetime(x, format=determinar_formato(x)))

# Extraer la fecha y la hora de la columna 'Checada'
df['Fecha'] = df['Checada'].dt.date
df['Hora'] = df['Checada'].dt.time

# Agrupar por fecha y PIN y calcular la hora mínima y máxima
result = df.groupby(['Fecha', 'PIN']).agg(
    Dispositivo=('Dispositivo', 'first'),
    Nombre=('Nombre de empleado', 'first'),
    Entrada=('Checada', 'min'),
    Salida=('Checada', 'max')
).reset_index()

# Formatear las columnas de entrada y salida como horas
result['Entrada'] = result['Entrada'].dt.time
result['Salida'] = result['Salida'].dt.time

# Si entrada y salida son iguales, poner None en salida
result['Salida'] = result.apply(lambda row: None if row['Entrada'] == row['Salida'] else row['Salida'], axis=1)

# Guardar el resultado en un archivo Excel
result.to_excel('resultados2.xlsx', index=False, columns=['PIN', 'Dispositivo', 'Nombre', 'Fecha', 'Entrada', 'Salida'])
