# Plan: correcciones para la v1.1

## Qué
Tres fallos encontrados en una revisión del repo, comprobados contra el código:

1. **Meses sin ventas desaparecen del gráfico** (p. ej. agosto cerrado por vacaciones).
   Se rellenan con 0 para que el gráfico no se salte meses.
2. **Cantidades con decimales se recortaban** (`2,5` pasaba a `2`, sin avisar).
   Decisión: se **aceptan** decimales (opción A). Hay negocios que venden por kilos,
   metros o litros; descartarlas les haría perder ventas reales.
3. **Pie de página**: con un nombre de Excel largo, "Fuente" y "Generado el" se pisaban.
   Se recorta el nombre con "…".

Además, el README gana una sección "Preguntas para el cliente" (decimales y duplicados).

Descartado de la revisión: "README desactualizado sobre devoluciones". Ya estaba corregido en el PR #7.

## Cómo
Una rama y un test por fallo. Después, Release v1.1 con el `.exe` nuevo.
