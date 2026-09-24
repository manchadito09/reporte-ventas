# Plan: probar y hacer robusto el informe con Excels de otras empresas

## Qué
Generar Excels de 4 empresas ficticias (ropa, bebidas, electrónica y un Excel "hecho a mano"),
ejecutar el informe con cada uno y arreglar lo que falle.

## Por qué
Hasta ahora solo se había probado con nuestro propio Excel, hecho a medida del programa.
Los Excels de clientes reales vienen distintos: más categorías, nombres largos, periodos que
cruzan de año, precios tecleados como texto y devoluciones.

## Resultado de la prueba
- Electrónica: el PDF se rompía (texto pisado, tabla fuera de la página).
- Excel "a mano": se perdía el 40 % de las ventas **sin avisar** (precios en texto, devoluciones).
- Ropa: los meses de 2025 salían sin año.

## Arreglos acordados
Precios en texto, devoluciones restadas del total (opción A), textos recortados con "…",
categorías pequeñas en "Otras", meses con año cuando cruza de año y aviso si se descarta
más del 5 % de las filas. Un test por arreglo.
