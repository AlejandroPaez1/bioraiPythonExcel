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
def buscar_nombre_por_pin(pin):
    nombre_empleado = None
    for empleado in data_checadores:
        if empleado['PIN'] == pin:
            nombre_empleado = empleado['Nombre de empleado']
            break  # Si encuentras el empleado, puedes salir del bucle
    return nombre_empleado

def buscar_dispositivo_por_pin(pin):
    nombre_empleado = None
    for empleado in data_checadores:
        if empleado['PIN'] == pin:
            nombre_empleado = empleado['Dispositivo']
            break  # Si encuentras el empleado, puedes salir del bucle
    return nombre_empleado

# Crear un nuevo archivo de Excel
wb = openpyxl.Workbook()
ws = wb.active

# Escribir encabezados al archivo de Excel
# Escribir datos en el archivo de Excel
# Obtener todas las fechas únicas de las checadas
fechas = sorted(set(checada for checadas_por_fecha in checadas_por_empleado_y_fecha.values() for checada in checadas_por_fecha.keys()))

# Escribir encabezados al archivo de Excel
encabezados = ["PIN", "Nombre", "Dispositivo"]
encabezados.extend([fecha.strftime("%d/%m/%Y") for fecha in fechas])
ws.append(encabezados)
# Escribir datos en el archivo de Excel
for pin, checadas_por_fecha in checadas_por_empleado_y_fecha.items():
    nombre_empleado = None
    dispositivo = None
    for fecha in fechas:
        if fecha in checadas_por_fecha:
            checadas = checadas_por_fecha[fecha]
            if not nombre_empleado:
                nombre_empleado = buscar_nombre_por_pin(pin)
                dispositivo = buscar_dispositivo_por_pin(pin)
            # Obtener todas las horas únicas de checadas en el día
            horas = sorted(set(checada.hour for checada in checadas))
            # Iterar sobre las horas de checada
            for hora in horas:
                # Filtrar las checadas de esta hora
                checadas_en_hora = [checada for checada in checadas if checada.hour == hora]
                # Dividir las checadas en dos grupos
                num_checadas = len(checadas_en_hora)
                group_size = num_checadas // 2
                grupo_1 = checadas_en_hora[:group_size]
                grupo_2 = checadas_en_hora[group_size:]
                # Escribir fila para el primer grupo de checadas
                for checada in sorted(grupo_1):
                    fila = [int(pin), nombre_empleado, dispositivo] + [""] * len(fechas)
                    fila[encabezados.index(fecha.strftime("%d/%m/%Y"))] = checada.strftime("%H:%M:%S")
                    ws.append(fila)
                # Escribir fila para el segundo grupo de checadas
                for checada in sorted(grupo_2):
                    fila = [int(pin), nombre_empleado, dispositivo] + [""] * len(fechas)
                    fila[encabezados.index(fecha.strftime("%d/%m/%Y"))] = checada.strftime("%H:%M:%S")
                    ws.append(fila)


# Guardar el archivo de Excel
nombre_archivo = "test223sintiuuupo.xlsx"
wb.save(nombre_archivo)