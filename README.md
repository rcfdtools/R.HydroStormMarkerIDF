# R.HydroStormMarker
Herramienta para la marcación y análisis de tormentas a partir de eventos de precipitación obtenidos de pluviometros o pluviografos.

Compatible con: Microsoft Office 2019.

En hidrológía, el estudio de la precipitación a partir de datos de tormentas registrados en estaciones de precipitación, permite analizar su comportamiento, duración, intensidad y patrón temporal.

R.HydroStormMarker, es una herramienta computacional que permite identificar y marcar los pulsos asociados a un mismo evento de lluvia, permitiendo conocer el valor total acumulado, duración e intensidad. Las tormentas identificadas pueden ser utilizadas para la construcción de curvas de Intesidad - Duración - Frecuencia ó IDF.

Los pulsos de la precipitación pueden contener ceros intermedios en los cuales el sensor de captura no registra los cambios en la precipitación, razón por la cual la App permite incluir hasta 3 ceros consecutivos por cada evento. 

Nota 1: Para el correcto funconamiento de la aplicación, antes de pegar los datos en la App, asegurese de indexar previamente los registros de 1 a n (Columna H de la hoja de Datos) ordenando los datos por fecha y hora. Desactive todos los filtros de datos antes de dar clic en EJECUTAR.

Nota 2: No se recomienda definir el número de ceros intermedios mayor a 1 sí ejecutó previamente la función de eliminación de registros con ceros sucesivos, debido a que no se mantiene la continuidad de fechas y horas en los registros. 

Nota 3: Para registros con frecuencia >= 30 minutos se recomienda utilizar solo 1 cero consecutivo.


PRUEBAS DE EJECUCIÓN
-------------------

LIMPIANDO REGISTROS CON DATOS EN CERO Y PROCESANDO LA SERIE

Intel Core I5-8300H, 4 Cores, RAM 16gb: 100K registros en 9 minutos.

AMD Ryzen 7 2700, 8 Cores 3.2GHz, RAM 32gb: 100K registros en 9 minutos.

SOLO PROCESANDO LA SERIE SIN LIMPIEZA DE REGISTROS EN CERO

AMD Ryzen 7 2700, 8 Cores 3.2GHz, RAM 32gb: 

  100K registros en 0.42 minutos.
  
  406K registros en 1.47 minutos.


NOVEDADES
----------------------

v.20191030
----------

Se incorporó nueva gráfica de todos los datos ingresados a partir del Índice consecutivo y el valor o dato registrado. 

En la hoja de Setup se incorporó un campo para rótulo personalizado a mostrar en gráficas.

Gráficas TormentaResumenGrafico y IndiceDatoGrafico son solo de lectura. Se pueden filtrar los datos a visualizar desde la hoja Datos y TormentaResumen.




v.20191029
----------

Hoja de Conceptos Generales y Diagrama de Flujo General.

Mejoras de rendimiento.

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
