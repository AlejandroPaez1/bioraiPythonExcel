# Código en Python para Analizar Faltas Semanales

Este proyecto en Python te permite analizar las faltas semanales de los empleados a partir de datos de checado. Sigue los siguientes pasos para utilizarlo:
## debe de haber un archivo llamado todos.csv 
lo que hace que el sistema busque cuales personas no tienen registros en ese rango de fechas 


## Pasos para CSV
El formato csv debe ser 


Tenant,Dispositivo,"Num de empleado",PIN,"Nombre de empleado",Departamento,"Tipo Nómina",Checada,Verificacion,"Entregado a RRHH","Mensaje de entrega","Estatus de entrega","Tipo de empleado","Tipo Checada",Temperatura
"Acme S.A. de C.V.","direccion",XXXX,XXXX,"Nombre apellido apellido",Ndepartamento,semana,"DD-MM-YYYY HH:MM:SS",tipo,"DD-MM-YYYY HH:MM:SS","TEXTO",Correcto,Externo,Device,'-


### 1. Selecciona el Método de Descarga en Formato CSV

- Inicia seleccionando el método de descarga en formato CSV desde el biorai. 
----------

![Método de Descarga en Formato CSV](https://github.com/AlejandroPaez1/bioraiPythonExcel/blob/main/assets/metodoDescarga.png)

tip: 
* Puedes descargar varios cvs y con excel pegarlo uno abajo del otro
* lo que genera un .xlsx de las fechas ingresadas
 
### 2. Ejecuta el Comando
para generar un archivo .exe

`pyi-makespec .\validarFaltasSemanalCSV.py` 

`pyinstaller .\validarFaltasSemanalCSV.py`


