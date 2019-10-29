# R.StormMarker
Herramienta para la marcación y análisis de tormentas a partir de eventos de precipitación obtenidos de pluviometros o pluviografos.

Compatible con: Microsoft Office 2019.

Nota 1: Para el correcto funconamiento de la aplicación, asegurese de indexar los registros de 1 a n (Columna H de la hoja de Datos) ordenando los datos por fecha y hora. Desactive todos los filtros de datos antes de dar clic en EJECUTAR.

Nota 2: No se recomienda definir el número de ceros intermedios si se ejecutó previamente la función de eliminación de registros con ceros sucesivos. 


NOVEDADES EN VERSIONES
----------------------

v.20191029
----------

Se ha incluído en la hoja "DatosEjemplo" datos que pueden ser utilizados para realizar pruebas de funcionamiento.

En la hoja "DatosEjemplo" se ha incluído un ejemplo de la formulación requerida para segmentar el campo de la fecha del registro de datos (Formato AAAA/mm/dd HH:MM:SS) a columnas independientes de año, mes, día y hora en formato de texto con relleno de ceros en 4 caracteres.


v.20191028
----------

Títulos principales y secundarios en gráficas son aignados automáticamente a partir de los datos generales registrados para la estación.

Optimizado el algoritmo de eliminación de registros con ceros consecutivos, reduciendo el tamaño del archivo y simplificando el análisis posterior de los datos. (Prueba de ejecución: 100K registros pueden ser limpiados y marcado en 9 minutos. Procesador Intel Core I5-8300H, RAM 16gb). Ejecutar este procedimiento no es obligatorio, sin embargo se recomienda su utilización para reducir el tamaño del archivo y los tiempos de análisis.

Tabla y gráfica dinámica para análisis de duraciones.

Grafica general con # de tormenta asignado, Precipitación Total e Intensidad.

Link R (Celda A1) en las hojas del libro para volver a la hoja MAIN.

Resúmen de tormentas encontradas y tiempo empleado en el procesamiento.
