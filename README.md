# LuzDia-VBA: Librería de Cálculos Astronómicos para Alumbrado Público Inteligente y Sostenible

![VBA](https://img.shields.io/badge/Language-VBA-purple?style=flat-square)
![Microsoft Excel](https://img.shields.io/badge/Platform-Excel-green?style=flat-square)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

## Descripción

**LuzDia-VBA** es una librería de funciones en Visual Basic for Applications (VBA) diseñada para ingenieros y técnicos de alumbrado público. Su objetivo es **simplificar y automatizar los cálculos astronómicos fundamentales** para poder realizar calculos de eficiencia energética y una gestión óptima de cualquier red de iluminación.

Esta herramienta permite determinar con **precisión las horas de orto y ocaso solar**, así como la **duración exacta de la noche y el día** para cualquier latitud y fecha. Estas funciones sirven como base para implementarse en cálculos más complejos para la programación óptima de los sistemas de alumbrado, contribuyendo al ahorro energético y a la planificación inteligente de nuestras ciudades.

El módulo permite calcular la declinación solar, mediante dos formula de aproximación, Cooper o Spencer(mas precisas, recomendada por defecto)

## ¿Por qué LuzDia-VBA?

En mi experiencia como ingeniero en alumbrado público, he constatado la necesidad de herramientas prácticas y accesibles que vayan más allá de la mera luminotecnia. **LuzDia-VBA nace de esa necesidad:** una solución pragmática para un problema recurrente: el **cálculo preciso de las horas de funcionamiento de las instalaciones**, esencial para la eficiencia operativa y el ahorro energético.

Mi visión es **democratizar el acceso a cálculos esenciales** para que más profesionales puedan diseñar y gestionar redes de iluminación de manera **verdaderamente eficiente y sostenible**. Esta herramienta es un paso hacia la optimización energética y la contribución a ciudades más inteligentes y respetuosas con el entorno.

## Características Clave

* **Cálculo Preciso:** Determina la hora de salida y puesta de sol (orto y ocaso) para cualquier latitud y fecha.
* **Duración de la Noche/Día:** Obtiene las horas de oscuridad y luz solar para optimizar los ciclos de encendido/apagado y la planificación energética.
* **Análisis Diario:** Permite generar datos para un día específico, fácilmente implementables para calcular un año completo, facilitando la planificación a largo plazo.
* **Integración Sencilla:** Funciones VBA fáciles de integrar en cualquier hoja de cálculo de Microsoft Excel.

## Cómo Usar

1.  **Descarga los archivos `LuzDia.bas` y `Trigonometria.bas`** desde este repositorio.
2.  **Abre tu archivo de Microsoft Excel.**
3.  **Accede al Editor de VBA** (Alt + F11). En tu proyecto de VBA (ej. `VBAProject (TuLibro.xlsm)`), haz clic derecho en "Módulos" -> "Importar Archivo..." y selecciona `LuzDia.bas` y `Trigonometria.bas`.
4.  **Integra las funciones** en tus propias hojas de cálculo o macros VBA llamándolas directamente en tus fórmulas.


## Potencial Futuro y Hoja de Ruta

**LuzDia-VBA** es el inicio de una serie de herramientas pensadas para el ingeniero de alumbrado. Aunque esta versión está en VBA, estoy activamente trabajando en la **portabilidad y expansión de estas y otras funcionalidades a librerías de código abierto en Python.**

Mi objetivo es crear un ecosistema de herramientas interoperables que continúen democratizando el conocimiento técnico y fomentando soluciones innovadoras para el alumbrado público inteligente y sostenible a escala global.

## Licencia

Este proyecto está licenciado bajo la [**Licencia MIT**](LICENSE.md). Eres libre de usar, modificar y distribuir el código, siempre y cuando se incluya la atribución original.


## Contacto

Para preguntas, sugerencias o colaboraciones, no dudes en contactarme:

* **Pablo Fernández**
* **pabalvfer@gmail.com**

### Ejemplo de uso en Excel:

```excel
=HoraOrto(HOY(); 42,5463; -6,59083; 1; 1)  ' Hora del amanecer en Ponferrada para hoy en formato decimal, utilizando la formula de Spencer
=HorasLuzDia(HOY(); 42,5463) ' Horas de luz en Ponferrada hoy
=Rad2Deg(DeclinacionSolar(HOY();1)) 'Declinación solar hoy en grados