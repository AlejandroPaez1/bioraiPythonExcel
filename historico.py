import validarFaltasSemanalCSV
from datetime import datetime

def generar_historico(wb, dias_no_checados_por_empleado, data_checadores):
    ws_faltantes = wb.create_sheet(title="Fechas de falta")

    # Escribir encabezados en la nueva hoja
    ws_faltantes.append(["Dispositivo", "PIN", "Nombre de empleado", "Fecha que Falto"])

    # Escribir datos de empleados faltantes en la nueva hoja
    for pin, dias_no_checados in dias_no_checados_por_empleado.items():
        if dias_no_checados:
            nombre_empleado = validarFaltasSemanalCSV.buscar_nombre_por_pin(pin, data_checadores)
            dispositivo_empleado = validarFaltasSemanalCSV.buscar_dispositivo_por_pin(pin, data_checadores)
            for fecha in dias_no_checados:
                # Verificar si el día es domingo
                if fecha.weekday() != 6:  # 0 es lunes, 1 martes, ..., 6 domingo
                    ws_faltantes.append([dispositivo_empleado, int(pin), nombre_empleado, fecha.strftime("%d/%m/%Y")])

    return wb
