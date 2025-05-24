# Módulo VBA para Cálculos Solares

**Autor**: Pablo Álvarez Fernández
**Versión**: 1.1
**Fecha de Revisión**: 04/02/2025

Este proyecto contiene un módulo VBA LUZ_SOLAR.bas diseñado para realizar cálculos astronómicos relacionados con la posición del sol, incluyendo la declinación solar, la horas de amanecer (orto), atardecer (ocaso) y la duración del día.
El módulo permite seleccionar la formula de aproximación para el calculo de la declinación solar, Cooper o Spencer(mas precisas, recomendada por defecto)

---

## Contenido del Módulo

El módulo incluye las siguientes funciones:

* `DeclinacionSolar(fecha As Variant, Optional ecuacion As Integer)`: Calcula la declinación solar.
* `HoraOrto(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer, Optional ecuacion As Integer)`: Devuelve la hora del amanecer.
* `HoraOcaso(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer, Optional ecuacion As Integer)`: Devuelve la hora del atardecer.
* `HorasLuzDia(dia As Variant, latitud As Double, Optional ecuacion As Integer)`: Calcula la duración de la luz diurna.
* `diaJuliano(fecha As Variant)`: Convierte una fecha a día juliano del año.
* `AnguloDiario(diaJuliano As Double, bisiesto As Boolean, Optional hora As Double)`: Calcula el ángulo diario en radianes.
* `EcuacionTiempo(fecha As Variant)`: Devuelve la corrección de la ecuación del tiempo.
* `HorasDeg2Reloj(HoraDeg As Double, fecha As Variant, longitud As Double, Optional timezone As Double)`: Convierte horas solares en grados a hora local.
* `IsLeapYear(year As Integer)`: Determina si un año es bisiesto.

---

## Dependencias

Este módulo requiere el módulo VBA 'TRIGONOMETRIA.BAS' que contiene las siguientes funciones auxiliares:

* `pi()`: Devuelve el valor de Pi.
* `Deg2Rad(Deg As Double)`: Convierte grados a radianes.
* `Rad2Deg(Rad As Double)`: Convierte radianes a grados.
* `ArcCos(Rad As Double)`: Calcula el arcocoseno.
* `ArcSin(Rad As Double)`: Calcula el arcoseno.
* `ArcTan(Rad As Double)`: Calcula la arcotangente.

---

### Ejemplo de uso en Excel:

```excel
=HoraOrto(HOY(); 42,5463; -6,59083,1,1)  ' Hora del amanecer en Ponferrada para hoy en formato decimal, utilizando la formula de Spencer
=HorasLuzDia(HOY(); 42,5463) ' Horas de luz en Ponferrada hoy
=Rad2Deg(DeclinacionSolar(HOY();1)) 'Declinación solar hoy en grados