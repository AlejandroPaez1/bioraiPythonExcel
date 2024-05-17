import csv
from collections import defaultdict
from datetime import datetime
import openpyxl

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
    else:
        raise ValueError("Formato de fecha y hora no reconocido")

def buscar_dias_checados_por_empleado(archivo_checadores):
    with open(archivo_checadores, 'r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        data_checadores = list(csv_reader)

    # Crear un diccionario para almacenar las checadas por PIN y fecha
    checadas_por_pin_y_fecha = defaultdict(lambda: defaultdict(list))
    for empleado in data_checadores:
        pin = empleado["PIN"]
        checada_str = empleado["Checada"]
        # Convertir la cadena de fecha y hora en un objeto datetime
        formato = determinar_formato(checada_str)
        checada_dt = datetime.strptime(checada_str, formato)
        # Almacenar la checada en el diccionario
        checadas_por_pin_y_fecha[pin][checada_dt.date()].append(checada_dt)

    return checadas_por_pin_y_fecha, data_checadores

# Llamada a la función para buscar los días checados por empleado
checadas_por_empleado_y_fecha, data_checadores = buscar_dias_checados_por_empleado("archivo.csv")

def buscar_nombre_dispositivo_por_pin(pin):
    for empleado in data_checadores:
        if empleado['PIN'] == pin:
            return empleado['Nombre de empleado'], empleado['Dispositivo']
    return None, None

# Crear un nuevo archivo de Excel
wb = openpyxl.Workbook()
ws = wb.active

# Escribir encabezados al archivo de Excel
encabezados = ["PIN", "NOMBRE", "DISPOSITIVO", "Fecha"]
ws.append(encabezados)
# Escribir datos en el archivo de Excel
for pin, checadas_por_fecha in checadas_por_empleado_y_fecha.items():
    nombre_empleado, dispositivo = buscar_nombre_dispositivo_por_pin(pin)
    for fecha, checadas in checadas_por_fecha.items():
        # Ordenar las checadas por hora
        checadas.sort()
        # Escribir la información en el archivo Excel
        fila = [int(pin), nombre_empleado, dispositivo, fecha.strftime("%d/%m/%Y")]
        # Agregar las horas de checada en una sola cadena separada por espacio
        # horas_checadas = (checada.strftime("%H:%M:%S") for checada in checadas)
        # fila.append( (checada.strftime("%H:%M:%S") for checada in checadas))
        fila.extend(checada.strftime("%H:%M:%S") for checada in checadas)

        ws.append(fila)


# Guardar el archivo de Excel
nombre_archivo = "test223sintiuuupov4.xlsx"
wb.save(nombre_archivo)
