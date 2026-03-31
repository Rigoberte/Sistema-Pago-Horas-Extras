# Sistema De Pago De Horas Extras

Esta aplicacion toma registros de ingreso y egreso de personal, los cruza con los datos de cada empleado y calcula, por cada turno:

- horas trabajadas totales,
- horas normales diurnas,
- horas extras normales,
- horas nocturnas.

El objetivo es estandarizar el calculo de horas para facilitar liquidaciones y controles.

## Interfaz grafica

La aplicacion se ejecuta con una interfaz que permite completar el workflow:

1. Seleccionar un Excel exportado desde Qontact.
2. Ver una previsualizacion del calculo en una tabla.
3. Editar valores habilitados antes de guardar.
4. Cargar el lote de forma temporal en el historico (estado `NO_CONFIRMADO`).
5. Confirmar el lote cargado (estado `CONFIRMADO`).

Para iniciar la interfaz:

```bash
python src/main.py
```

## Como funciona

La aplicacion se utiliza con el siguiente flujo de trabajo:

1. Se configura cada empleado con su jornada (hs jornal), su sueldo y el valor de la hora extra.
2. De forma opcional, se cargan los feriados para que sean considerados en el calculo.
3. Para registrar horas extras, el usuario selecciona un reporte exportado desde Qontact.
4. La aplicacion valida que todos los empleados del reporte tengan configuracion completa.
5. Si falta informacion, la aplicacion solicita completarla antes de continuar.
6. Con los datos completos, genera y muestra en pantalla el detalle calculado.
7. El usuario puede revisar y editar ese detalle antes de la confirmacion final. Tambien puede agregar comentarios por turno si lo desea.
8. Una vez confirmado, se pueden generar reportes por empleado y por periodo de dias.

## Regla De Horas Nocturnas

Se considera hora nocturna todo tramo trabajado:

- desde las 21:00 en adelante, o
- antes de las 06:00.

Esto cubre correctamente casos como:

- ingresos de madrugada (por ejemplo 01:00),
- turnos normales sin nocturnidad,
- egresos despues de las 21:00,
- ingresos a las 21:00 o mas tarde.

Importante: una hora nocturna no se cuenta tambien como extra normal. Las extras normales se aplican solo a tiempo diurno excedente.

## Estructura Esperada De Archivos

### ReporteQontact.xlsx
Columnas requeridas:

- NOMBRE Y APELLIDO
- DESDE
- HASTA
- INGRESO
- EGRESO

## Salida

La aplicacion genera un archivo Excel con estas columnas principales:

- NOMBRE_Y_APELLIDO
- INGRESO
- EGRESO
- COMENTARIOS
- HORAS_TRABAJADAS
- HORAS_NORMALES_DIURNAS
- HORAS_EXTRAS_NORMALES
- HORAS_NOCTURNAS
- IMPORTE_JORNAL
- IMPORTE_HORAS_EXTRAS
- IMPORTE_HORAS_NOCTURNAS
- IMPORTE_TOTAL