# R.StormMarker
Herramienta para la marcación y análisis de tormentas a partir de eventos de precipitación obtenidos de pluviometros o pluviografos.

Compatible con: Microsoft Office 2019.

Nota 1: Para el correcto funconamiento de la aplicación, antes de pegar los datos en la App, asegurese de indexar previamente los registros de 1 a n (Columna H de la hoja de Datos) ordenando los datos por fecha y hora. Desactive todos los filtros de datos antes de dar clic en EJECUTAR.

Nota 2: No se recomienda definir el número de ceros intermedios mayor a 1 sí ejecutó previamente la función de eliminación de registros con ceros sucesivos, debido a que no se mantiene la continuidad de fechas y horas en los registros. 


PRUEBAS DE EJECUCIÓN
-------------------

Procesador Intel Core I5-8300H, RAM 16gb: 100K registros pueden ser limpiados y procesados en 9 minutos.

Procesador AMD Ryzen 7 2700 8 Cores 3.2GHz, RAM 32gb: 100K registros pueden ser limpiados y procesados en 9 minutos.


NOVEDADES
----------------------

v.20191029
----------

Se ha incluído en la hoja "DatosEjemplo" datos que pueden ser utilizados para realizar pruebas de funcionamiento.

En la hoja "DatosEjemplo" se ha incluído un ejemplo de la formulación requerida para segmentar el campo de la fecha del registro de datos (Formato AAAA/mm/dd HH:MM:SS) a columnas independientes de año, mes, día y hora en formato de texto con relleno de ceros en 4 caracteres.


v.20191028
----------

Títulos principales y secundarios en gráficas son aignados automáticamente a partir de los datos generales registrados para la estación.

Optimizado el algoritmo de eliminación de registros con ceros consecutivos, reduciendo el tamaño del archivo y simplificando el análisis posterior de los datos. Ejecutar este procedimiento no es obligatorio, sin embargo se recomienda su utilización para reducir el tamaño del archivo y los tiempos de análisis.

Tabla y gráfica dinámica para análisis de duraciones.

Grafica general con # de tormenta asignado, Precipitación Total e Intensidad.

Link R (Celda A1) en las hojas del libro para volver a la hoja MAIN.

Resúmen de tormentas encontradas y tiempo empleado en el procesamiento.
