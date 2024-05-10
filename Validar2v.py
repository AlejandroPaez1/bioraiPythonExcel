import csv
import os
from tkinter import Tk, filedialog, messagebox
from datetime import datetime, timedelta
import openpyxl
from collections import defaultdict

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

def obtener_nombre_archivo(nombre_base, extension):
    contador = 1
    while True:
        nombre_archivo = f"{nombre_base}_{contador}.{extension}"
        if not os.path.exists(nombre_archivo):
            return nombre_archivo
        contador += 1

def buscar_no_checadores(archivo_checadores):
    with open(archivo_checadores, 'r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        data_checadores = list(csv_reader)

    # Crear un diccionario para almacenar las fechas de las checadas por PIN
    checadas_por_pin = defaultdict(list)
    for empleado in data_checadores:
        pin = empleado["PIN"]
        nombre = empleado["Nombre de empleado"]
        dispositivo = empleado["Dispositivo"]
        checada_str = empleado["Checada"]
        # Convertir la cadena de fecha y hora en un objeto datetime
        formato = determinar_formato(checada_str)
        checada_dt = datetime.strptime(checada_str, formato)
        checadas_por_pin[pin].append((checada_dt, nombre, dispositivo))

    # Crear un diccionario para almacenar los días en que no se checó por PIN
    dias_no_checados_por_pin = defaultdict(list)
    for pin, checadas in checadas_por_pin.items():
        # Obtener el conjunto único de fechas (días) en que se checó
        dias_checados = set(checada.date() for checada, _, _ in checadas)
        if not dias_checados:
            continue  # Saltar si no hay registros para este PIN
        # Obtener la fecha mínima y máxima de las checadas
        primer_dia = min(checada.date() for checada, _, _ in checadas)
        ultimo_dia = max(checada.date() for checada, _, _ in checadas)
        primer_dia -= timedelta(days=1)
        ultimo_dia += timedelta(days=1)
        # Ajustar la fecha mínima y máxima para incluir el día siguiente al mínimo
        todos_los_dias = [primer_dia + timedelta(days=d) for d in range((ultimo_dia - primer_dia).days + 1)]

        # Generar todos los días entre la fecha mínima y máxima
        # todos_los_dias = [primer_dia + timedelta(days=d) for d in range((ultimo_dia - primer_dia).days)]
        # Determinar los días en que no se checó, excluyendo los domingos
        dias_no_checados = [dia for dia in todos_los_dias if dia not in dias_checados and dia.weekday() != 6]
        # Almacenar los días no checados en el diccionario
        dias_no_checados_por_pin[pin] = (dias_no_checados, [nombre for _, nombre, _ in checadas], [dispositivo for _, _, dispositivo in checadas])

    return dias_no_checados_por_pin, data_checadores

def leer_csv(nombre_archivo):
    diccionario = {}
    with open(nombre_archivo, 'r', encoding='utf-8') as archivo:
        lector = csv.DictReader(archivo)
        for fila in lector:
            pin = fila['PIN']
            diccionario[pin] = fila
    return diccionario

# Crear ventana de Tkinter para seleccionar el archivo
root = Tk()
root.withdraw()  # Ocultar la ventana principal

# Solicitar al usuario que seleccione el archivo CSV
archivo_checadores = filedialog.askopenfilename(title="Seleccione el archivo CSV", filetypes=[("Archivos CSV", "*.csv")])

# Buscar días en que no se checó por empleado
dias_no_checados_por_empleado, data_checadores = buscar_no_checadores(archivo_checadores)

# Crear un nuevo archivo de Excel
wb = openpyxl.Workbook()
ws = wb.active

# Encabezados
ws['A1'] = "PIN"
ws['B1'] = "Nombre"
ws['C1'] = "Dispositivo"
# Escribir encabezados de fechas dinámicamente
fechas = sorted(set(fecha for dias_no_checados, _, _ in dias_no_checados_por_empleado.values() for fecha in dias_no_checados))
for idx, fecha in enumerate(fechas, start=1):
    ws.cell(row=1, column=3+idx).value = f"{fecha}"

# Escribir datos en el archivo de Excel
row = 2
for pin, (dias_no_checados, nombres, dispositivos) in dias_no_checados_por_empleado.items():
    if dias_no_checados:  # Verificar si hay días no checados
        # Escribir el PIN, nombre del empleado y dispositivo
        ws.cell(row=row, column=1).value = int(pin)  # Convertir PIN a entero
        ws.cell(row=row, column=2).value = nombres[0]  # Suponiendo que el nombre siempre estará en la primera posición
        ws.cell(row=row, column=3).value = dispositivos[0]  # Suponiendo que el dispositivo siempre estará en la primera posición

        # Escribir las fechas en las columnas correspondientes
        for idx, fecha in enumerate(fechas, start=1):
            if fecha in dias_no_checados:
                ws.cell(row=row, column=3+idx).value = fecha.strftime("%d/%m/%Y")
            else:
                ws.cell(row=row, column=3+idx).value = "A"
        row += 1

# Leer el archivo 'todos.csv'
archivo_csv = "todos.csv"
diccionario_resultante = leer_csv(archivo_csv)

# Obtener el PIN de los registros ya presentes en el archivo CSV
pines_presentes = set(diccionario_resultante.keys())

# Agregar los registros que no están presentes en el archivo CSV al final del archivo Excel
for pin, (dias_no_checados, nombres, dispositivos) in dias_no_checados_por_empleado.items():
    if pin not in pines_presentes:
        # Agregar datos al archivo CSV
        diccionario_resultante[pin] = {'PIN': pin, 'Nombre de empleado': nombres[0], 'Dispositivo': dispositivos[0]}
        for fecha in dias_no_checados:
            diccionario_resultante[pin][fecha.strftime("%d/%m/%Y")] = "A"

# Actualizar el archivo CSV con los nuevos datos
# Actualizar el archivo CSV con los nuevos datos
with open(archivo_csv, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['Tenant', 'Dispositivo', 'Num de empleado', 'PIN', 'Nombre de empleado', 'Departamento', 'Tipo Nómina', 'Checada', 'Verificacion', 'Entregado a RRHH', 'Mensaje de entrega', 'Estatus de entrega', 'Tipo de empleado', 'Tipo Checada', 'Temperatura'] + fechas
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in diccionario_resultante.values():
        # Crear un nuevo diccionario con las claves específicas para los encabezados existentes
        new_row = {key: row.get(key, '') for key in fieldnames}
        writer.writerow(new_row)

# Eliminar la primera columna de fechas
ws.delete_cols(4)
# Eliminar la última columna
ws.delete_cols(ws.max_column)

# Guardar el archivo de Excel
nombre_base = "Resultadodecdsv_"
extension = "xlsx"
nombre_archivo = obtener_nombre_archivo(nombre_base, extension)
wb.save(nombre_archivo)

# Indicar al usuario que se ha creado el archivo
print(f"Se ha creado el archivo '{nombre_archivo}' con éxito.")
messagebox.showinfo("Archivo creado", f"Se ha creado el archivo '{nombre_archivo}' con éxito.")